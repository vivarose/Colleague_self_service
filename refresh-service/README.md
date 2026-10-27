# On-demand daily catalog refresh

The static Pages app calls a Cloudflare Worker when opened. One Durable Object
stores the shared state for **all visitors**. Dates use fixed EST (UTC-05:00)
year-round, not `America/New_York`; midnight is always 05:00 UTC.

There is no scheduled source download. The first request for a date dispatches
the GitHub workflow. Other requests reuse the pending attempt or its successful
result. Failed attempts may be retried on a later visit. A workflow must claim
the active request before downloading anything, so duplicate dispatches and
workflow reruns cannot download twice for the same attempt. An attempt queued
past its date is rejected before downloading.

If GitHub's dispatch response or a completion callback is lost, a Durable Object
alarm reconciles the workflow status. It never abandons a claimed, running
workflow simply because time elapsed. Failed/cancelled runs release the gate.
An unconfirmed dispatch without a matching run is released after ten minutes;
any subsequently starting stale run fails its claim. These alarms only check
GitHub; they do not download college data.

## Deployment

1. Merge this integration into the branch that will serve the site (`main` by
   default). The workflow must exist on the repository's default branch for
   `workflow_dispatch`. Enable Actions and allow it to write repository contents.
2. In GitHub **Settings > Pages**, change the build source to **GitHub Actions**.
   The workflow deploys with `actions/deploy-pages`; branch builds are not used.
   Ensure the `github-pages` environment permits deployments from the target
   branch without manual approval if you want unattended refreshes.
3. Use a Cloudflare account with Workers/Durable Objects enabled. From this
   directory, run `npm ci` and `npx wrangler login`. Review
   [wrangler.jsonc](wrangler.jsonc):
   - `ALLOWED_ORIGIN`: exact Pages origin, without a path or trailing slash.
   - `GITHUB_REPOSITORY`: owner/repository.
   - `GITHUB_REF`: publication branch.
   - `GITHUB_WORKFLOW`: `refresh-catalog.yml`.
4. Create a fine-grained GitHub access token scoped to this repository with
   **Actions: read and write**. Store it only as a Worker secret:

   ```sh
   npx wrangler secret put GITHUB_TOKEN
   ```

5. Generate a strong random callback secret locally (at least 32 random bytes).
   Store the same value as Worker secret `CALLBACK_SECRET` and GitHub repository
   Actions secret `REFRESH_CALLBACK_SECRET`. Do not commit or paste secrets into
   chat, configuration files, command arguments, or the Pages client.

   ```sh
   npx wrangler secret put CALLBACK_SECRET
   npx wrangler deploy
   ```

6. Set GitHub repository Actions variable `REFRESH_SERVICE_URL` to the deployed
   Worker's HTTPS origin, for example `https://hamilton-catalog-refresh.ACCOUNT.workers.dev`.
   Set `service_url` in [refresh-config.json](../refresh-config.json) to that same
   origin and commit it. The checked-in empty value deliberately disables the
   client until a backend is deployed.
7. Publish the initial Pages artifact **before** relying on an app visit to
   trigger anything. In an authenticated terminal, run the workflow with a
   request created by the gate:

   ```sh
   curl --fail-with-body -X POST \
     -H 'Origin: https://crgibbons.github.io' \
     https://YOUR-WORKER.workers.dev/refresh
   ```

   This is the same public request made by opening the app. It is not a bypass of
   the per-date claim. Check the Actions run and Pages deployment for success.

## Visitor behavior

The app remains usable while refresh runs and displays its status. It polls
`/status` without triggering more downloads. Once publication succeeds, it offers
a cache-busting reload link; it does not discard an active course selection
automatically. On a workflow error, it explicitly says it is showing last
published data. A later app visit can retry. If a prior date's refresh was still
running at midnight, the current date starts only after that attempt finishes
and a current-date request is made.

The workflow persists the spreadsheet, page, provenance, and immutable snapshots
to the configured branch before deploying Pages. A successful date is recorded
only after deployment. A deployment failure may therefore leave a newer data
commit without a successful publication; a later request can retry that date.

## Source validation

[refresh_catalogue.py](../refresh_catalogue.py) downloads:

- The explicitly configured Spring 2027 schedule.
- The registrar's stable seat-allocation link, whose resolved PDF must identify
  Spring 2027.
- Public SmartCatalog subject listings discovered from its site map, then the
  course pages needed for the schedule. The catalog is the current edition,
  **not a verified Spring 2027 edition**.

Downloads are serial and paced for catalog pages. Empty/malformed responses,
duplicate identities, missing prerequisite markup, a schedule shrinking by over
10%, fewer than 200 seat rows, or catalog coverage below 60% stop publication.
These are conservative review guards, not proof of feed completeness. Review
and update term-specific limits and URLs for future terms. Catalog pages absent
from discovery are logged and receive an explicit missing-entry note rather
than retaining stale descriptions. Source allocation discrepancies are warnings;
original values and `OPEN` semantics are preserved.

Seat keys use `Course Number+` from the schedule and retain `W` on both sides.
This fixes the legacy W-suffix mismatch without collapsing distinct `101` and
`101W` sections. Section `01`, numeric `1`, and spreadsheet `1.0` normalize alike;
lab suffixes remain distinct.

Every source's retrieval time, resolved URL, and SHA-256 are saved in
`source-provenance.json`. Retrieval time is not a source modification time.
Network/parsing/build checks happen in staging before published files change.
No live upstream calls occur in unit tests.

## Validation

Workflow callbacks identify themselves as `Hamilton-Catalog-Refresh`; Cloudflare
may reject Python's default urllib user agent before requests reach the Worker.

From the repository root:

```sh
python -m pip install -r requirements-refresh.txt
python -m unittest discover -s tests -v
node --test tests/corrigendum.test.js refresh-service/worker.test.js
```

Tests simulate durable storage and request serialization. They do not prove
account permissions, Cloudflare deployment, or GitHub Pages environment settings.
After deployment, verify one dispatch for concurrent visits, no new dispatch
after success on that date, retry after a failed run, and the published reload
link. GitHub API failures are logged in Workers and remain pending until
reconciliation can determine whether a run is active.

Run `npm ci && npm run check` in this directory for a Wrangler deployment
dry-run. This bundles and checks bindings without uploading or deploying.
