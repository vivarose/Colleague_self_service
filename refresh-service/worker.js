export function estDate(now = Date.now()) {
  return new Date(now - 5 * 60 * 60 * 1000).toISOString().slice(0, 10);
}

function json(value, status = 200) {
  return Response.json(value, { status });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    if (!['/refresh', '/status', '/claim', '/complete'].includes(url.pathname)) {
      return json({ error: 'Not found' }, 404);
    }
    const callback = ['/claim', '/complete'].includes(url.pathname);
    if (callback) {
      if (request.method !== 'POST' ||
          request.headers.get('Authorization') !== `Bearer ${env.CALLBACK_SECRET}` ||
          !env.CALLBACK_SECRET) {
        return json({ error: 'Unauthorized' }, 401);
      }
    } else {
      if (request.headers.get('Origin') !== env.ALLOWED_ORIGIN) {
        return json({ error: 'Origin not allowed' }, 403);
      }
      if (request.method === 'OPTIONS') {
        return new Response(null, { headers: cors(env.ALLOWED_ORIGIN) });
      }
      if (request.method !== (url.pathname === '/refresh' ? 'POST' : 'GET')) {
        return json({ error: 'Method not allowed' }, 405);
      }
    }
    const stub = env.REFRESH_GATE.get(env.REFRESH_GATE.idFromName('catalog'));
    const result = await stub.fetch(request);
    const response = new Response(result.body, result);
    response.headers.set('Cache-Control', 'no-store');
    if (!callback) {
      for (const [key, value] of Object.entries(cors(env.ALLOWED_ORIGIN))) {
        response.headers.set(key, value);
      }
    }
    return response;
  }
};

function cors(origin) {
  return {
    'Access-Control-Allow-Origin': origin,
    'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
    Vary: 'Origin'
  };
}

export class RefreshGate {
  constructor(state, env) {
    this.state = state;
    this.env = env;
  }

  async github(path, options = {}) {
    const response = await fetch(
      `https://api.github.com/repos/${this.env.GITHUB_REPOSITORY}/${path}`,
      {
        ...options,
        headers: {
          Authorization: `Bearer ${this.env.GITHUB_TOKEN}`,
          Accept: 'application/vnd.github+json',
          'X-GitHub-Api-Version': '2022-11-28',
          'User-Agent': 'Hamilton-Catalog-Refresh',
          'Content-Type': 'application/json'
        }
      }
    );
    if (!response.ok) throw new Error(`GitHub API returned HTTP ${response.status}`);
    return response.status === 204 ? null : response.json();
  }

  async fetch(request) {
    // Keep storage checks and dispatch decisions atomic across all visitors.
    return this.state.blockConcurrencyWhile(async () => {
      const path = new URL(request.url).pathname;
      if (path === '/claim' || path === '/complete') {
        return this.callback(path, await request.json());
      }
      const date = estDate();
      const key = `day:${date}`;
      let day = await this.state.storage.get(key);
      const active = await this.state.storage.get('active');
      if (active) {
        return json({ date, status: 'pending', request_id: active.request_id,
          refresh_date: active.date, message: 'A refresh is queued or running.' });
      }
      if (day?.status === 'succeeded' || path === '/status') {
        return json(day || { date, status: 'not_requested' });
      }
      // Retrying a failed date creates a new attempt; successful dates are immutable.
      day = { date, status: 'pending', request_id: crypto.randomUUID(),
        started_at: new Date().toISOString() };
      await this.state.storage.put(key, day);
      await this.state.storage.put('active', day);
      await this.state.storage.setAlarm(Date.now() + 60_000);
      try {
        await this.github(`actions/workflows/${this.env.GITHUB_WORKFLOW}/dispatches`, {
          method: 'POST',
          body: JSON.stringify({
            ref: this.env.GITHUB_REF,
            inputs: { refresh_date: date, request_id: day.request_id }
          })
        });
      } catch (error) {
        // A lost response may still mean GitHub accepted the dispatch. Reconcile
        // before retrying; /claim ensures duplicate runs cannot download sources.
        console.error('Refresh dispatch uncertain:', error);
        day.message = 'Dispatch could not be confirmed; checking GitHub before retrying.';
        await this.state.storage.put(key, day);
      }
      return json(day, 202);
    });
  }

  async callback(path, body) {
    const active = await this.state.storage.get('active');
    if (!active || active.request_id !== body.request_id || active.date !== body.refresh_date) {
      return json({ proceed: false, error: 'This refresh attempt is no longer active.' }, 409);
    }
    if (path === '/claim') {
      if (active.claimed || active.date !== estDate()) {
        return json({ proceed: false });
      }
      active.claimed = true;
      active.run_id = String(body.run_id);
      await this.state.storage.put('active', active);
      await this.state.storage.put(`day:${active.date}`, active);
      return json({ proceed: true });
    }
    if (!active.claimed || String(body.run_id) !== active.run_id ||
        !['succeeded', 'failed'].includes(body.status)) {
      return json({ error: 'Invalid completion' }, 409);
    }
    await this.finish(active, body.status);
    return json({ status: body.status });
  }

  async finish(active, status) {
    await this.state.storage.put(`day:${active.date}`, {
      date: active.date, request_id: active.request_id, status,
      completed_at: new Date().toISOString(),
      message: status === 'failed' ? 'Refresh failed; showing last good data. A later visit may retry.' : undefined
    });
    await this.state.storage.delete('active');
    await this.state.storage.deleteAlarm();
  }

  async alarm() {
    return this.state.blockConcurrencyWhile(async () => {
      const active = await this.state.storage.get('active');
      if (!active) return;
      try {
        let run;
        if (active.run_id) {
          run = await this.github(`actions/runs/${active.run_id}`);
        } else {
          const runs = await this.github(
            `actions/workflows/${this.env.GITHUB_WORKFLOW}/runs?event=workflow_dispatch&per_page=100`
          );
          run = runs.workflow_runs.find(item => item.display_title === `Catalog refresh ${active.request_id}`);
        }
        if (run?.status === 'completed') {
          // Only /complete after deployment can mark success. A lost callback
          // is recovered from the workflow's successful conclusion.
          await this.finish(active, active.claimed && run.conclusion === 'success' ? 'succeeded' : 'failed');
          return;
        }
        if (!run && !active.claimed &&
            Date.now() - Date.parse(active.started_at) > 10 * 60_000) {
          await this.finish(active, 'failed');
          return;
        }
      } catch (error) {
        console.error('Refresh reconciliation failed:', error);
      }
      await this.state.storage.setAlarm(Date.now() + 60_000);
    });
  }
}
