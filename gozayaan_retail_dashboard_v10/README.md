# GoZayaan Retail Dashboard

FastAPI + vanilla HTML/CSS/JS retail performance dashboard, reading a
live Google Sheet through a Google Apps Script JSON API. Deployed on
Render.

## Structure

Single root — no versioned subfolders. Everything Render needs to build
and run is right here:

```
app.py              — backend: FastAPI app, all business logic
static/index.html   — page shell
static/app.js        — frontend logic
static/styles.css    — styling incl. dark mode
render.yaml           — Render service definition
requirements.txt      — pinned Python dependencies
RENDER_ENV.txt         — reference for the env vars set in Render's dashboard
.python-version        — 3.12.8
```

Previous deploys used `gozayaan_retail_dashboard_v8/` and `_v9/`
subfolders with an incrementing version number. That pattern caused the
repo to accumulate multiple divergent copies of the same files — this
structure replaces all of that. Going forward, this root **is** the
app; there is no next version-numbered folder. Commit changes directly
here.

## Live source

The app reads a live Google Sheet through a Google Apps Script web app,
configured via:

```
GOZAAYAN_SHEETS_API_URL   — the Apps Script /exec URL (see RENDER_ENV.txt)
GOZAAYAN_REFRESH_MINUTES  — auto-refresh interval, default 30
GOZAAYAN_BATCH_FETCH      — set to 1 once the Apps Script has the
                             batch (?sheets=...) handler deployed;
                             falls back to per-tab fetching if unset
                             or if the batch call fails
```

Live mode is strict: if the Apps Script API is unavailable and there's
no previously-cached data to fall back on, the app reports the outage
rather than showing stale numbers unlabeled. If it *does* have
previously-cached data, it keeps serving that with a `(STALE)` /
`degraded: true` flag rather than going down over a transient blip.

No XLSX snapshot files are part of this deployment — production always
runs in live mode. If you want a local offline copy of the workbook for
testing, keep it in a local `data/` folder (already gitignored) rather
than committing it; it adds several MB to every clone for zero
production benefit.

## Pages

1. **Dashboard** — date-locked, reads the live Dashboard tab directly.
2. **Banani / Chattogram / HQ / Motijheel** — Start Date + End Date.
3. **Daily GMV** — company-wide GMV by day over a selected range.
4. **WBR (Weekly Business Review)** — defaults to the most recently
   completed Monday–Sunday week; that locked week only advances to a
   newly-completed week the following Monday at 5 PM Dhaka time
   (`GOZAAYAN_WBR_LOCK_HOUR`, default 17), giving a data-entry buffer.
   A specific week can still be reviewed manually via `?week_end=`.

## Other features

- USD labelling throughout.
- PDF and XLSX export (Dashboard, WBR, branch pages — not yet wired for
  Daily GMV).
- URL-preserved filters.
- Month-end snapshot generation.
- Dark mode (follows OS preference by default; toggle overrides and is
  remembered per-browser).
- GoZayaan brand palette, including a dark-mode-safe variant of every
  surface color, not just the base palette.
- Per-refresh-cycle backend caching (repeated computation within one
  30-minute window is served from cache, not recomputed) and a 60s
  client-side cache for perceived navigation speed.

## Render

**Root Directory:** leave blank — this repo root is the whole app now.
If Render's dashboard still has it set to `gozayaan_retail_dashboard_v8`
or `_v9` from a previous deploy, change it to blank/`.` and redeploy.

**Build Command:** `pip install -r requirements.txt`

**Start Command:** `uvicorn app:app --host 0.0.0.0 --port $PORT`

**Python:** `3.12.8` via `.python-version`.

**Environment variables:** see `RENDER_ENV.txt` for the values; set them
in Render's dashboard, not committed as real secrets elsewhere.

## Target Helper — action needed before Q4

`TARGET_MONTH_COLS` in `app.py` only has July/August/September mapped,
matching what's currently in the Target Helper sheet. Requesting a
target for an unmapped month raises a clear `409` rather than silently
reusing a previous month's numbers. Before October: add the next column
block to Target Helper (pattern: +5 columns per month — Oct would be
`S:V`, Nov `X:AA`, Dec `AC:AF`) and add the entry to `TARGET_MONTH_COLS`.
