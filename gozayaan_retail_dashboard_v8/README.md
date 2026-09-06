# GoZayaan Retail Dashboard v8

v8 switches the live source from the broken Google Published CSV endpoint to the working Google Apps Script JSON API.

## Live source

The app uses:
`GOZAAYAN_SHEETS_API_URL`

Default value is the deployed Apps Script web-app URL supplied for this project.

The Python app requests each required sheet as:
`?sheet=Dashboard`
`?sheet=All Tour`
etc.

No Google Cloud project, billing account, service account, published CSV, or GID map is required for the live adapter.

## Source refresh

- Automatic Python source refresh every 30 minutes.
- Large manual `REFRESH DATA` button.
- Live mode is strict: if the Apps Script API is unavailable, the app does not silently fall back to the bundled XLSX snapshot.
- The bundled XLSX files are retained only for local/offline testing when `GOZAAYAN_SHEETS_API_URL` is explicitly set empty.

## Pages

1. Dashboard — date locked, reads the live Dashboard tab.
2. Banani — Start Date + End Date.
3. Chattogram — Start Date + End Date.
4. HQ — Start Date + End Date.
5. Motijheel — Start Date + End Date.
6. WBR — Week Ending.

## Other features retained

- USD labelling.
- PDF and XLSX export.
- URL-preserved filters.
- Month-end snapshot generation.
- GoZayaan brand palette.

## Render

Root Directory:
`gozayaan_retail_dashboard_v8`

Build Command:
`pip install -r requirements.txt`

Start Command:
`uvicorn app:app --host 0.0.0.0 --port $PORT`

Python:
`3.12.8` via `.python-version`.

Recommended Render environment variable:
`GOZAAYAN_SHEETS_API_URL=https://script.google.com/macros/s/AKfycbxzvfVm3zjo8ADZBCxE7jcSnKIaISEmRzbsIfWiIAUsiHHTDnLU6uhsYlhrfYMjVYLT/exec`

Optional:
`GOZAAYAN_REFRESH_MINUTES=30`

Remove old v7 variables:
`GOZAAYAN_PUBLISHED_URL`
`GOZAAYAN_GID_MAP`

## Testing

The package was smoke-tested using an Apps Script-compatible mock backed by the provided workbook snapshots. Verified:
- live adapter JSON parsing
- source initialization
- Dashboard endpoint
- Banani date-range calculation
- WBR calculation
- H2 performance calculation
- PDF export
- XLSX export
- month-end snapshot creation

The real Apps Script URL cannot be network-tested from this environment, but it has already been manually verified in the user's browser with successful JSON responses for `action=health` and `sheet=Dashboard`.
