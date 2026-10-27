import assert from 'node:assert/strict';
import { test } from 'node:test';
import worker, { estDate, RefreshGate } from './worker.js';

function harness() {
  const values = new Map();
  const calls = [];
  let queue = Promise.resolve();
  const state = {
    storage: {
      get: async key => structuredClone(values.get(key)),
      put: async (key, value) => { values.set(key, structuredClone(value)); },
      delete: async key => values.delete(key),
      setAlarm: async () => {},
      deleteAlarm: async () => {}
    },
    blockConcurrencyWhile: callback => {
      const next = queue.then(callback);
      queue = next.catch(() => {});
      return next;
    }
  };
  const gate = new RefreshGate(state, {
    GITHUB_WORKFLOW: 'refresh-catalog.yml', GITHUB_REF: 'main'
  });
  gate.github = async (path, options) => {
    calls.push({ path, options });
    return null;
  };
  const request = async (path = '/refresh', body) => {
    const response = await gate.fetch(new Request(`https://worker.test${path}`, {
      method: path === '/status' ? 'GET' : 'POST',
      ...(body ? { body: JSON.stringify(body) } : {})
    }));
    return { status: response.status, body: await response.json() };
  };
  return { gate, values, calls, request };
}

test('fixed EST changes date at 05:00 UTC even during daylight saving time', () => {
  assert.equal(estDate(Date.parse('2026-07-01T04:59:59Z')), '2026-06-30');
  assert.equal(estDate(Date.parse('2026-07-01T05:00:00Z')), '2026-07-01');
  assert.equal(estDate(Date.parse('2026-12-01T04:59:59Z')), '2026-11-30');
});

test('concurrent visitors dispatch once, claim once, and never redownload after success', async () => {
  const { request, calls } = harness();
  const visitors = await Promise.all(Array.from({ length: 20 }, () => request()));
  assert.equal(calls.length, 1);
  assert.equal(new Set(visitors.map(value => value.body.request_id)).size, 1);
  const first = visitors[0].body;
  const body = { request_id: first.request_id, refresh_date: first.date, run_id: '123' };
  assert.equal((await request('/claim', body)).body.proceed, true);
  assert.equal((await request('/claim', body)).body.proceed, false);
  assert.equal((await request('/complete', { ...body, status: 'succeeded' })).status, 200);
  assert.equal((await request()).body.status, 'succeeded');
  assert.equal(calls.length, 1);
});

test('failure permits a new attempt; stale callbacks cannot release its lock', async () => {
  const { request, calls } = harness();
  const first = (await request()).body;
  const body = { request_id: first.request_id, refresh_date: first.date, run_id: '123' };
  await request('/claim', body);
  await request('/complete', { ...body, status: 'failed' });
  const retry = (await request()).body;
  assert.notEqual(retry.request_id, first.request_id);
  assert.equal(calls.length, 2);
  assert.equal((await request('/complete', { ...body, status: 'succeeded' })).status, 409);
  assert.equal((await request()).body.request_id, retry.request_id);
});

test('new EST date starts a refresh only on demand', async () => {
  const { request, values, calls } = harness();
  values.set('day:2000-01-01', { date: '2000-01-01', status: 'succeeded' });
  assert.equal((await request('/status')).body.status, 'not_requested');
  assert.equal(calls.length, 0);
  assert.equal((await request()).body.status, 'pending');
  assert.equal(calls.length, 1);
});

test('lost dispatch response holds the lock until reconciliation', async () => {
  const { gate, request, calls } = harness();
  gate.github = async () => { calls.push({}); throw new Error('Connection lost'); };
  await request();
  await request();
  assert.equal(calls.length, 1);
});

test('alarm recovers completion when callback is lost', async () => {
  const { gate, request } = harness();
  const first = (await request()).body;
  await request('/claim', {
    request_id: first.request_id, refresh_date: first.date, run_id: '123'
  });
  gate.github = async () => ({ status: 'completed', conclusion: 'success' });
  await gate.alarm();
  assert.equal((await request()).body.status, 'succeeded');
});

test('an expired queued attempt cannot claim and a running attempt never times out', async () => {
  const { gate, request, values } = harness();
  const old = {
    date: '2000-01-01', request_id: 'old', status: 'pending',
    started_at: '2000-01-01T00:00:00Z'
  };
  values.set('active', old);
  assert.equal((await request('/claim', {
    request_id: 'old', refresh_date: old.date, run_id: '1'
  })).body.proceed, false);
  values.set('active', { ...old, claimed: true, run_id: '1' });
  gate.github = async () => ({ status: 'in_progress' });
  await gate.alarm();
  assert.equal((await request()).body.request_id, 'old');
});

test('public endpoints enforce origin and callbacks require the backend secret', async () => {
  const env = { ALLOWED_ORIGIN: 'https://crgibbons.github.io', CALLBACK_SECRET: 'test-secret' };
  const blocked = await worker.fetch(new Request('https://worker.test/refresh', {
    method: 'POST', headers: { Origin: 'https://other.example' }
  }), env);
  assert.equal(blocked.status, 403);
  const callback = await worker.fetch(new Request('https://worker.test/complete', {
    method: 'POST'
  }), env);
  assert.equal(callback.status, 401);
});
