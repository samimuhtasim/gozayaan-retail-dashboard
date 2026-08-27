
# GoZayaan Retail Dashboard v7

v7 adds the requested:
1. PDF and XLSX export for every page.
2. URL-preserved filters and date ranges.
3. Month-end snapshots.

It remains FastAPI + HTML/CSS/JS.

## URL state

Examples:
- `/ ?page=dashboard`
- `/?page=banani&start=2026-08-01&end=2026-08-25`
- `/?page=wbr&week_end=2026-08-30`

Copying a URL reproduces the same filter state.

## Exports

Each page has:
- XLSX export
- PDF export

The export endpoints are:
- `/api/export/xlsx?...`
- `/api/export/pdf?...`

## Month-end snapshots

A month-end snapshot contains:
- `snapshot.json`
- dashboard close XLSX/PDF
- Banani XLSX/PDF
- Chattogram XLSX/PDF
- HQ XLSX/PDF
- Motijheel XLSX/PDF
- WBR XLSX/PDF

Automatic snapshot logic checks for the final day of the month around 23:55 Asia/Dhaka whenever the service process is alive.

Manual snapshot:
`POST /api/snapshots/month-end?snapshot_date=2026-08-31`

List:
`GET /api/snapshots`

Download:
`GET /api/snapshots/2026-08/banani.pdf`

## Important Render note

A Render free instance can sleep. Therefore a strict 23:55 automatic snapshot cannot be guaranteed on the free tier. For guaranteed month-end capture, use an always-on service or an external scheduler that calls:
`POST /api/snapshots/month-end?snapshot_date=YYYY-MM-DD`
at month-end.

## Data source

The application first attempts the configured published Google Sheet URL, then falls back to the bundled XLSX snapshots.

Set:
`GOZAAYAN_PUBLISHED_URL`
`GOZAAYAN_GID_MAP`

The app has no dependency on `Dashboard API`.

## Deployment

Build:
`pip install -r requirements.txt`

Start:
`uvicorn app:app --host 0.0.0.0 --port $PORT`

Python recommendation:
`3.12.x`

## Validation done

The v7 package was smoke-tested against the provided Dashboard and CSS Helper snapshots, including:
- API startup
- Dashboard endpoint
- Branch endpoint
- WBR endpoint
- XLSX export
- PDF export
- URL query parsing in frontend
- month-end snapshot creation


## Render

`render.yaml` is included for a simple Render setup. If the v7 folder is the repository root, Render can use the blueprint directly. If v7 is inside a subfolder, keep that subfolder as Render's Root Directory and use the build/start commands shown in `render.yaml`.

## Persistence caveat

The month-end snapshot feature writes files to the app filesystem. On hosting with ephemeral storage, those files can disappear on redeploy/restart. For a durable archive, point `SNAPSHOT_DIR` at persistent storage or add object storage later. The capture logic itself is ready.


## v7.1 fixes

- Removed duplicate refresh-button rendering.
- Live Google Sheet mode is now strict: if `GOZAAYAN_PUBLISHED_URL` is set and the live source cannot be read, the app does NOT fall back to XLSX snapshots.
- `/api/health` now returns `ok: false` and the source error in live failure mode.
- The frontend shows an explicit LIVE SOURCE UNAVAILABLE state instead of silently showing stale snapshot data.
- Local development without `GOZAAYAN_PUBLISHED_URL` still uses bundled snapshots.
