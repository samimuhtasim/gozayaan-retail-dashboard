"""
Full-stack check: boots the real app.py (network mocked, SOURCE swapped
for synthetic data) and drives it through FastAPI's TestClient — actual
HTTP requests through actual ASGI routing — including fetching the real
static/index.html, app.js, styles.css this app will serve, and hitting
/api/daily-gmv the way the browser actually will.
"""
import sys, types, io as _io
import numpy as np
import pandas as pd
from datetime import date, datetime
from zoneinfo import ZoneInfo

BDT = ZoneInfo("Asia/Dhaka")
results = []
def check(label, cond):
    results.append((label, cond))
    print(("PASS  " if cond else "FAIL  ") + label)

src = open("app.py").read()
mod = types.ModuleType("appmod")
mod.__dict__["__file__"] = "app.py"
mod.__dict__["__name__"] = "appmod"

import requests as _requests
def _boot_get(*a, **k):
    class R:
        status_code=200; text=""
        def raise_for_status(self): pass
        def json(self): return {"ok": True, "rows": [["hdr"], ["a"]]}
    return R()
_orig = _requests.get
_requests.get = _boot_get
try:
    exec(compile(src, "app.py", "exec"), mod.__dict__)
finally:
    _requests.get = _orig

check("app module boots", mod.app is not None)

# --- same synthetic dataset construction as test_v10.py -----------------
def wide_df(nrows=115, ncols=45):
    return pd.DataFrame(np.zeros((nrows, ncols)), dtype=object)
def set_cell(df, row_1based, col_letter, value):
    df.iat[row_1based - 2, mod.col_idx(col_letter)] = value
def set_date(df, idx0, col_letter, d):
    df.iat[idx0, mod.col_idx(col_letter)] = pd.Timestamp(d)

target_df = wide_df()
branch_base = {"Banani": 19, "Chattogram": 45, "HQ": 73, "Motijheel": 102}
for branch, base in branch_base.items():
    for month, (c0, c1) in [(7, ("D", "E")), (8, ("I", "J")), (9, ("N", "O"))]:
        s = {7: 100, 8: 1000, 9: 10000}[month]
        for off, mult in [(0,1),(1,2),(2,3),(3,4),(4,5)]:
            set_cell(target_df, base+off, c0, s*mult); set_cell(target_df, base+off, c1, s*mult)

D1, D2, D3 = date(2026, 9, 7), date(2026, 9, 10), date(2026, 9, 13)

def make_css(flight, hotel, tour):
    df = wide_df(10, 20)
    set_date(df, 0, "A", D1)
    set_cell(df, 2, "F", flight); set_cell(df, 2, "G", hotel); set_cell(df, 2, "H", tour)
    set_cell(df, 2, "L", flight*0.1); set_cell(df, 2, "M", hotel*0.1); set_cell(df, 2, "N", tour*0.1)
    set_cell(df, 2, "O", 2); set_cell(df, 2, "P", 1)
    return df

def make_fit(flight, hotel, tour, visa, footfall, visa_col):
    df = wide_df(10, 45)
    set_date(df, 0, "A", D2)
    set_cell(df, 2, "J", flight); set_cell(df, 2, "M", hotel); set_cell(df, 2, "N", tour)
    set_cell(df, 2, visa_col, visa); set_cell(df, 2, "Q", footfall)
    set_cell(df, 2, "R", 3); set_cell(df, 2, "V", 1)
    return df

all_tour = wide_df(10, 20); set_date(all_tour, 0, "A", D3)
set_cell(all_tour, 2, "O", 5000); set_cell(all_tour, 2, "L", 10); set_cell(all_tour, 2, "I", 300)
all_visa = wide_df(10, 20); set_date(all_visa, 0, "A", D3)
set_cell(all_visa, 2, "D", 900); set_cell(all_visa, 2, "F", 4); set_cell(all_visa, 2, "L", 65.6)
dv = wide_df(10, 10)
set_cell(dv, 2, "A", "x"); set_cell(dv, 3, "A", "y"); set_cell(dv, 2, "B", "x")
dash_df = wide_df(30, 12)
set_cell(dash_df, 4, "A", "Product")
for r, name in zip([5,6,7,8,9], ["Flight","Hotel","Tour","Visa","Overall — All Products"]):
    set_cell(dash_df, r, "A", name)
set_cell(dash_df, 5, "B", 1000000); set_cell(dash_df, 5, "H", 2000000)
set_cell(dash_df, 6, "B", 40000);   set_cell(dash_df, 6, "H", 90000)
set_cell(dash_df, 7, "B", 800000);  set_cell(dash_df, 7, "H", 1900000)
set_cell(dash_df, 8, "B", 35000);   set_cell(dash_df, 8, "H", 75000)

TABS = {
    "DS1Banani CSS": make_css(1000,100,200), "DS1CTG CSS": make_css(1500,150,250),
    "DS1HQ CSS": make_css(500,50,80), "DS1Mot CSS": make_css(800,90,120),
    "DS1Banani FIT": make_fit(400,40,90,60,25,"AM"), "DS1CTG FIT": make_fit(300,30,70,45,18,"AJ"),
    "DS1Mot FIT": make_fit(200,20,40,30,12,"AJ"),
    "All Tour": all_tour, "All Visa": all_visa, "Target Helper": target_df,
    "DV": dv, "Dashboard": dash_df, "WBR Format": wide_df(10,10),
}
mod.SOURCE.source = object()
mod.SOURCE.last_refresh = datetime(2026, 9, 16, 2, 8, tzinfo=BDT)
mod.SOURCE.tab = lambda name: TABS[name]
for fn in [mod.branch_period, mod.h2_performance, mod.dashboard_h2,
           mod.wbr_targets, mod.branch_month_targets, mod.daily_company_gmv]:
    if hasattr(fn, "cache_clear"): fn.cache_clear()

# --- real ASGI HTTP client ------------------------------------------------
from fastapi.testclient import TestClient
client = TestClient(mod.app)

r = client.get("/")
check("GET / returns 200", r.status_code == 200)
check("index.html mentions Daily GMV nav entry", 'data-page="daily-gmv"' in r.text)
check("index.html has the theme toggle button", 'id="themeToggle"' in r.text)

r = client.get("/static/app.js")
check("GET /static/app.js returns 200", r.status_code == 200)
check("served app.js has the daily-gmv label", '"daily-gmv"' in r.text)
check("served app.js has the sidebar no-hover fix", "no-hover" in r.text)
check("served app.js has theme toggle logic", "THEME_KEY" in r.text)

r = client.get("/static/styles.css")
check("GET /static/styles.css returns 200", r.status_code == 200)
check("served styles.css has dark mode block", "prefers-color-scheme:dark" in r.text)
check("served styles.css fixes the sidebar-hidden overlap", "padding-left:64px" in r.text)

r = client.get("/api/health")
check("GET /api/health returns 200", r.status_code == 200)
check("health reports ok", r.json()["ok"] is True)

r = client.get("/api/dashboard")
check("GET /api/dashboard returns 200", r.status_code == 200)
d = r.json()
check("dashboard rows present and header not duplicated", d["rows"][0][0] == "Product" and sum(1 for row in d["rows"] if row and row[0]=="Product")==1)

r = client.get("/api/branch/banani", params={"start": "2026-09-01", "end": "2026-09-15"})
check("GET /api/branch/banani returns 200", r.status_code == 200)
bd = r.json()
check("branch response has footfall_avg_daily", "footfall_avg_daily" in bd["progress"])

r = client.get("/api/wbr")  # no week_end -> should use the locked default now
check("GET /api/wbr with no week_end returns 200 (used to be a required param)", r.status_code == 200)
wd = r.json()
check("wbr response carries manual_override flag", "manual_override" in wd and wd["manual_override"] is False)
check("wbr response carries next_rollover_at", "next_rollover_at" in wd)

r = client.get("/api/daily-gmv", params={"start": "2026-09-01", "end": "2026-09-15"})
check("GET /api/daily-gmv returns 200", r.status_code == 200)
gd = r.json()
check("daily-gmv has 15 rows for a 15-day range", len(gd["rows"]) == 15)
check("daily-gmv row shape matches what app.js's renderDailyGmv expects",
      set(gd["rows"][0].keys()) >= {"date","flight","hotel","tour","visa","total"})
check("daily-gmv total_gmv / average_daily_gmv present", "total_gmv" in gd and "average_daily_gmv" in gd)

r = client.get("/api/export/xlsx", params={"page": "dashboard"})
check("GET /api/export/xlsx (dashboard) returns 200", r.status_code == 200)
check("xlsx export content-type looks right", "spreadsheet" in r.headers.get("content-type","") or "excel" in r.headers.get("content-type","").lower() or len(r.content) > 1000)

r = client.get("/api/export/pdf", params={"page": "wbr", "end": "2026-09-13"})
check("GET /api/export/pdf (wbr) returns 200", r.status_code == 200)
check("pdf export bytes look like a real PDF", r.content[:4] == b"%PDF")

print()
failed = [l for l, c in results if not c]
print(f"FINAL {len(results)-len(failed)}/{len(results)} passed")
if failed:
    print("FAILED:")
    for l in failed: print("  -", l)
sys.exit(1 if failed else 0)
