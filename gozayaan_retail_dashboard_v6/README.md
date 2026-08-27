# GoZayaan Retail Dashboard v6

v6 replaces Streamlit with a standard FastAPI + HTML/CSS/JS web application.

## Architecture

Google Sheet -> Python data/calculation engine -> FastAPI -> browser

The Google Sheet remains the source system.

The application supports:

- 6 pages
- Dashboard with locked date
- Banani / Chattogram / HQ / Motijheel with Start Date + End Date
- WBR with Week Ending
- USD labels
- GoZayaan brand palette
- automatic source refresh every 30 minutes
- large manual `REFRESH DATA` button
- last refresh / source status indicator
- XLSX snapshot fallback for local testing

## Live Google Sheet

The default `GOZAAYAN_PUBLISHED_URL` is the published Google Sheet URL provided for this project.

For production, set `GOZAAYAN_GID_MAP` if automatic tab discovery fails. Example:

GOZAAYAN_GID_MAP={"Dashboard":"0","WBR Format":"123456789",...}

The app first tries the live published Sheet. If that fails, it falls back to the bundled snapshots.

## Deployment on Render

Create a GitHub repository with:

app.py
requirements.txt
static/index.html
static/app.js
static/styles.css
data/Dashboard (1).xlsx
data/CSS Helper (1).xlsx

Render settings:

Build command:
pip install -r requirements.txt

Start command:
uvicorn app:app --host 0.0.0.0 --port $PORT

Environment variables:

GOZAAYAN_PUBLISHED_URL=<published Google Sheet URL>
GOZAAYAN_REFRESH_MINUTES=30

If automatic tab discovery cannot identify tabs, also add:
GOZAAYAN_GID_MAP=<JSON mapping>

No Streamlit account is required.

## Local test

Windows:
python -m venv .venv
.venv\\Scripts\\activate
pip install -r requirements.txt
uvicorn app:app --reload

Then open:
http://127.0.0.1:8000

## Important calculation notes

The locked Dashboard page reads the Dashboard sheet's A4:J9 block as its source of truth.

Branch pages use the imported DS1 CSS / DS1 FIT layers that the existing Dashboard workbook consumes from CSS Helper / FIT sources.

WBR intentionally calculates the four displayed product GMVs and therefore does not reproduce the existing WBR J6 total defect that omits Hotel.

The source WBR carries Hotel Net Revenue as a fixed 83.5 value; v6 preserves that behavior.

All dashboard monetary values are USD.

## Data integrity

The first production milestone should be reconciliation of every v6 endpoint against the live Sheet for:
- 25 Aug 2026 branch dashboards
- 24-30 Aug 2026 WBR
- H2 dashboard
before opening the public URL widely.
