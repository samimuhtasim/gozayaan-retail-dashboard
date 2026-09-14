# v9 stability patch — deploy notes

## Root cause

`/api/health` reported:

```
DS1Banani CSS: 404 Client Error ... script.googleusercontent.com/macros/echo?user_content_key=...
Target Helper: 404 Client Error ... script.googleusercontent.com/macros/echo?user_content_key=...
```

Apps Script answers `/exec` with a 302 to a single-use
`script.googleusercontent.com` URL. `warm()` fired 13 tabs at 8 concurrent
workers; under that load Google invalidated some content keys before
`requests` followed the redirect, giving a 404 on the **second hop**.

11 of 13 tabs succeeded. `warm()` then raised, `refresh()` set
`self.source = None`, and every endpoint returned `AttributeError` →
plain-text 500 → the frontend's `JSON.parse` failed with
`Unexpected token 'I'`. Two flaky HTTP hops took the whole dashboard down.

## What changed in `app.py`

| # | Change | Fixes |
|---|--------|-------|
| 1 | `_fetch` retries 3× with backoff (`_fetch_once` holds the original logic) | The transient 404 itself |
| 2 | `WARM_WORKERS` 8 → 2 | Reduces how often Google throttles |
| 3 | `warm()` returns errors instead of raising | Stops discarding 11 good tabs over 2 bad ones |
| 4 | `refresh()` keeps the last good source on failure; `PARTIAL` / `STALE` naming | Stale numbers with a warning instead of a dead dashboard |
| 5 | `tab()` raises `HTTPException(503, detail=last_error)` when source is `None` | Frontend gets JSON with the real reason, not a parse error |
| 6 | `auto_refresh_loop` sleeps first and uses `asyncio.to_thread` | The duplicate boot refresh, and the blocked event loop / port bind |
| 7 | `month_end_snapshot_loop` uses `asyncio.to_thread` | Same blocking issue, once a minute |
| 8 | `degraded` flag on `/api/health` and `/api/source` | Lets the UI show a staleness banner |
| 9 | Opt-in `_fetch_batch` behind `GOZAAYAN_BATCH_FETCH` | The structural fix — see below |

Missing tabs are still fetched lazily by `sheet()` on first use, so a
partial warm costs one slow request, not a broken page.

## Deploy order

**Step 1 — ship `app.py` now.** No env changes needed. Expected result:
the 404s retry away, boot time drops from ~180s to ~60s, and a future
partial failure degrades one section instead of the whole app.

**Step 2 — batch the Apps Script (this week).** Open the Apps Script
editor, replace `doGet` with `AppsScript_Code.gs`, confirm
`SPREADSHEET_ID` matches, then **Deploy > New deployment > Web app**:

- Execute as: **Me**
- Who has access: **Anyone** (otherwise Google serves an HTML login page)

Update `GOZAAYAN_SHEETS_API_URL` on Render to the new `/exec` URL, then
set `GOZAAYAN_BATCH_FETCH=1`. 13 requests become 1; the redirect race
cannot happen. Startup should land in single-digit seconds.

If the batch endpoint misbehaves, the adapter falls back to per-tab
fetching automatically and records it in `last_error`. To roll back
fully, set `GOZAAYAN_BATCH_FETCH=0` — no redeploy required.

## Env vars (all optional, defaults shown)

```
GOZAAYAN_WARM_WORKERS=2
GOZAAYAN_FETCH_ATTEMPTS=3
GOZAAYAN_FETCH_BACKOFF=1.5
GOZAAYAN_BATCH_FETCH=0
```

Also: `refresh_minutes` is currently **15**, not the 30 in the handoff.
At 15 you re-roll the throttling dice twice as often. Put it back to 30.

## Still to do in `static/app.js` (not supplied, so not patched)

**1. Check the status before parsing.** This is what turned a clean 503
into `Unexpected token 'I'`:

```js
const res = await fetch(url);
if (!res.ok) {
  let msg = `Server returned ${res.status}`;
  try { msg = (await res.json()).detail || msg; } catch (_) {}
  throw new Error(msg);
}
const data = await res.json();
```

**2. Render the staleness banner.** `/api/health` now returns `degraded`
and a `source` of `LIVE GOOGLE SHEET (PARTIAL)` or `(STALE)`. Show an
amber "data as of <last_refresh>" strip rather than the red blocking
error — the numbers are real, just not fresh.

**3. Verify which progress object the branch bars read.** `api_branch`
returns both `progress` (driven by the selected end date) and
`locked_month_progress` (driven by today in Dhaka). Feedback point C
requires the month-locked one. The backend supports it; confirm the
frontend actually uses it.

## Verification

`test_patches.py` loads the adapter classes against a mocked Apps Script
and covers all of it — including a reproduction of the exact two-tab 404.
19/19 passing. Run with:

```
GOZAAYAN_FETCH_BACKOFF=0.01 python3 test_patches.py
```

## Unrelated but worth 10 minutes

`render.yaml` has no `rootDir`. The README says it must be
`gozayaan_retail_dashboard_v8`. It works because the dashboard UI has it
set, but anyone redeploying from the blueprint gets a broken service. Add:

```yaml
    rootDir: gozayaan_retail_dashboard_v8
```

And confirm Render's Root Directory points at wherever this v9 `app.py`
actually lives — if v9 went into a new folder and Render still builds
`v8/`, none of this ships.
