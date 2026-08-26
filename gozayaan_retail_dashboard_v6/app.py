
from __future__ import annotations

import asyncio
import io
import json
import os
import re
import threading
import time
from contextlib import asynccontextmanager
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Dict, Optional

import numpy as np
import pandas as pd
import requests
from fastapi import FastAPI, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
import uvicorn


# ============================================================
# Configuration
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"

SNAPSHOT_DASH = DATA_DIR / "Dashboard (1).xlsx"
SNAPSHOT_CSS = DATA_DIR / "CSS Helper (1).xlsx"

PUBLISHED_URL = os.getenv(
    "GOZAAYAN_PUBLISHED_URL",
    "https://docs.google.com/spreadsheets/d/e/2PACX-1vTv8b75NS5qseGpuVvNHvCckH7luGvMhbMS_fSB8f3rf6VOvPKt7_5n9rzcvFWBrfozh66N06_cFHIU/pubhtml",
).strip()

GID_MAP_ENV = os.getenv("GOZAAYAN_GID_MAP", "").strip()

REFRESH_MINUTES = int(os.getenv("GOZAAYAN_REFRESH_MINUTES", "30"))

BRANCH_SOURCE = {
    "Banani": {
        "css": "DS1Banani CSS",
        "fit": "DS1Banani FIT",
        "visa": "AM",
        "dv": "A",
    },
    "Chattogram": {
        "css": "DS1CTG CSS",
        "fit": "DS1CTG FIT",
        "visa": "AJ",
        "dv": "B",
    },
    "HQ": {
        "css": "DS1HQ CSS",
        "fit": None,
        "visa": None,
        "dv": "D",
    },
    "Motijheel": {
        "css": "DS1Mot CSS",
        "fit": "DS1Mot FIT",
        "visa": "AJ",
        "dv": "C",
    },
}

TARGET_ROWS = {
    "Banani": 19,
    "Chattogram": 45,
    "HQ": 73,
    "Motijheel": 102,
}


# ============================================================
# Utilities
# ============================================================

def num(value) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float, np.number)):
        if np.isfinite(float(value)):
            return float(value)
        return 0.0
    s = str(value).strip().replace(",", "")
    if not s or s.lower() in {"nan", "none", "n/a", "na", "-"}:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def fmt_date(d: date) -> str:
    return d.isoformat()


def parse_date(value: str) -> date:
    try:
        return date.fromisoformat(value)
    except Exception as exc:
        raise ValueError(f"Invalid date: {value}") from exc


def month_start(d: date) -> date:
    return d.replace(day=1)


def month_end(d: date) -> date:
    nxt = (d.replace(day=28) + timedelta(days=4)).replace(day=1)
    return nxt - timedelta(days=1)


def q3_start() -> date:
    return date(2026, 7, 1)


def q3_end() -> date:
    return date(2026, 9, 30)


def q4_start() -> date:
    return date(2026, 10, 1)


def q4_end() -> date:
    return date(2026, 12, 31)


def h2_start() -> date:
    return date(2026, 7, 1)


def h2_end() -> date:
    return date(2026, 12, 31)


def pct(actual, target):
    target = num(target)
    return (num(actual) / target) if target else None


def change(cur, prev):
    prev = num(prev)
    return ((num(cur) - prev) / prev) if prev else None


def col_idx(letter: str) -> int:
    n = 0
    for ch in letter.upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def normalise(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [
        str(c).strip() if str(c).strip() else f"col_{i}"
        for i, c in enumerate(df.columns)
    ]
    return df


def get_col(df: pd.DataFrame, letter: str) -> pd.Series:
    idx = col_idx(letter)
    if idx >= df.shape[1]:
        return pd.Series([np.nan] * len(df), index=df.index)
    return df.iloc[:, idx]


def dates(df: pd.DataFrame, letter: str = "A") -> pd.Series:
    return pd.to_datetime(get_col(df, letter), errors="coerce").dt.date


def sumifs(df: pd.DataFrame, value_col: str, date_col: str, start: date, end: date) -> float:
    if df.empty:
        return 0.0
    d = dates(df, date_col)
    v = pd.to_numeric(get_col(df, value_col), errors="coerce").fillna(0)
    mask = (d >= start) & (d <= end)
    return float(v.loc[mask].sum())


def count_nonempty(df: pd.DataFrame, letter: str) -> int:
    if df.empty:
        return 0
    s = get_col(df, letter)
    return int(
        s.astype(str)
        .str.strip()
        .replace("nan", "")
        .ne("")
        .sum()
    )


# ============================================================
# Source adapters
# ============================================================

class SnapshotSource:
    def __init__(self, dashboard_path: Path, css_path: Path):
        self.dashboard_path = dashboard_path
        self.css_path = css_path
        self._dash = pd.ExcelFile(dashboard_path)
        self._css = pd.ExcelFile(css_path)
        self._dash_cache: Dict[str, pd.DataFrame] = {}
        self._css_cache: Dict[str, pd.DataFrame] = {}

    def dash(self, sheet: str) -> pd.DataFrame:
        if sheet not in self._dash_cache:
            self._dash_cache[sheet] = normalise(
                pd.read_excel(self._dash, sheet_name=sheet)
            )
        return self._dash_cache[sheet].copy()

    def css(self, sheet: str) -> pd.DataFrame:
        if sheet not in self._css_cache:
            self._css_cache[sheet] = normalise(
                pd.read_excel(self._css, sheet_name=sheet)
            )
        return self._css_cache[sheet].copy()


class PublishedSource:
    def __init__(self, published_url: str):
        self.published_url = published_url
        self._gid_map: Optional[Dict[str, str]] = None
        self._cache: Dict[str, pd.DataFrame] = {}

    def _discover_gids(self) -> Dict[str, str]:
        if self._gid_map is not None:
            return self._gid_map

        if GID_MAP_ENV:
            data = json.loads(GID_MAP_ENV)
            self._gid_map = {str(k): str(v) for k, v in data.items()}
            return self._gid_map

        r = requests.get(self.published_url, timeout=30)
        r.raise_for_status()
        html = r.text

        found: Dict[str, str] = {}

        # Published Sheets commonly embeds tab definitions in JS.
        patterns = [
            r'"gid":"(\d+)".{0,1000}?"name":"([^"]+)"',
            r'"name":"([^"]+)".{0,1000}?"gid":"(\d+)"',
            r'gid=(\d+)[^<]{0,500}>([^<]{1,120})<',
            r'>([^<]{1,120})</a>[^<]{0,250}gid=(\d+)',
        ]

        for pattern in patterns:
            for m in re.finditer(pattern, html, re.I | re.S):
                a, b = m.groups()
                if a.isdigit():
                    gid, name = a, b
                else:
                    name, gid = a, b
                name = re.sub(r"<[^>]+>", "", name)
                name = re.sub(r"\s+", " ", name).strip()
                if name and gid.isdigit():
                    found[name] = gid

        if not found:
            raise RuntimeError(
                "Could not discover published Google Sheet tabs. "
                "Set GOZAAYAN_GID_MAP in the deployment secrets."
            )

        self._gid_map = found
        return found

    def sheet(self, name: str) -> pd.DataFrame:
        if name in self._cache:
            return self._cache[name].copy()

        gids = self._discover_gids()
        if name not in gids:
            raise KeyError(
                f"Published tab '{name}' was not found. "
                f"Available: {sorted(gids)}"
            )

        gid = gids[name]
        pub_base = self.published_url.split("/pubhtml")[0] + "/pub"
        url = f"{pub_base}?gid={gid}&single=true&output=csv"

        r = requests.get(url, timeout=45)
        r.raise_for_status()

        df = pd.read_csv(io.BytesIO(r.content))
        df = normalise(df)
        self._cache[name] = df
        return df.copy()


class SourceManager:
    """
    Keeps one source object in memory and refreshes it explicitly.

    Live mode is preferred. Snapshot mode is the fallback for local testing.
    """

    REQUIRED_TABS = [
        "Dashboard",
        "Target Helper",
        "DV",
        "DS1Banani FIT",
        "DS1CTG FIT",
        "DS1Mot FIT",
        "DS1Banani CSS",
        "DS1CTG CSS",
        "DS1HQ CSS",
        "DS1Mot CSS",
        "All Tour",
        "All Visa",
        "WBR Format",
    ]

    def __init__(self):
        self.lock = threading.RLock()
        self.source = None
        self.source_name = "UNINITIALIZED"
        self.last_refresh: Optional[datetime] = None
        self.last_error: Optional[str] = None
        self.refresh()

    def refresh(self):
        with self.lock:
            if PUBLISHED_URL:
                try:
                    live = PublishedSource(PUBLISHED_URL)
                    for tab in self.REQUIRED_TABS:
                        live.sheet(tab)
                    self.source = live
                    self.source_name = "LIVE GOOGLE SHEET"
                    self.last_error = None
                    self.last_refresh = datetime.now()
                    return
                except Exception as exc:
                    self.last_error = str(exc)

            if not SNAPSHOT_DASH.exists() or not SNAPSHOT_CSS.exists():
                raise RuntimeError(
                    "Live Google Sheet could not be loaded and bundled snapshots are missing."
                )

            self.source = SnapshotSource(
                SNAPSHOT_DASH,
                SNAPSHOT_CSS,
            )
            self.source_name = "BUNDLED SNAPSHOT"
            self.last_refresh = datetime.now()

    def tab(self, name: str) -> pd.DataFrame:
        with self.lock:
            if isinstance(self.source, SnapshotSource):
                return self.source.dash(name)
            return self.source.sheet(name)


SOURCE = SourceManager()


# ============================================================
# Business calculations
# ============================================================

def target_value(df: pd.DataFrame, row_1based: int, col: str) -> float:
    # pd.read_excel() consumes the first worksheet row as the header,
    # so workbook row N maps to dataframe row N-2.
    idx = row_1based - 2
    if idx < 0 or idx >= len(df):
        return 0.0
    return num(df.iloc[idx, col_idx(col)])


def dashboard_h2() -> dict:
    """
    The Dashboard sheet is the source of truth for the locked H2 page.
    We do not derive these values independently.
    """
    df = SOURCE.tab("Dashboard")

    # Actual workbook block: A4:J9.
    block = df.iloc[2:8, :10].fillna("")

    rows = []
    for row in block.itertuples(index=False, name=None):
        rows.append(list(row))

    return {
        "page": "dashboard",
        "currency": "USD",
        "title": "H-2 Achievement Dashboard — All Products",
        "rows": rows,
    }


def branch_month_targets(branch: str, end: date) -> dict:
    t = SOURCE.tab("Target Helper")
    row = TARGET_ROWS[branch]

    cols = (
        ("D", "E") if end.month == 7 else
        ("I", "J") if end.month == 8 else
        ("N", "O") if end.month == 9 else
        ("I", "J")
    )

    return {
        "Flight": target_value(t, row, cols[0]) + target_value(t, row, cols[1]),
        "Hotel": target_value(t, row + 1, cols[0]) + target_value(t, row + 1, cols[1]),
        "Tour": (
            target_value(t, row + 2, cols[0])
            + target_value(t, row + 2, cols[1])
            + target_value(t, row + 3, cols[0])
            + target_value(t, row + 3, cols[1])
        ),
        "Visa": target_value(t, row + 4, cols[0]) + target_value(t, row + 4, cols[1]),
    }


def branch_period(branch: str, start: date, end: date) -> dict:
    cfg = BRANCH_SOURCE[branch]

    css = SOURCE.tab(cfg["css"])
    fit = SOURCE.tab(cfg["fit"]) if cfg["fit"] else None

    # CSS Helper-fed DS1 CSS
    css_flight = sumifs(css, "F", "A", start, end)
    css_hotel = sumifs(css, "G", "A", start, end)
    css_tour = sumifs(css, "H", "A", start, end)

    css_flight_recv = sumifs(css, "L", "A", start, end)
    css_hotel_recv = sumifs(css, "M", "A", start, end)
    css_tour_recv = sumifs(css, "N", "A", start, end)

    css_flight_bookings = sumifs(css, "O", "A", start, end)
    css_hotel_bookings = sumifs(css, "P", "A", start, end)

    # Three FIT options are imported to the Dashboard workbook.
    retail_flight = retail_hotel = retail_tour = retail_visa = footfall = 0.0
    fit_flight_bookings = fit_hotel_bookings = 0.0

    if fit is not None:
        retail_flight = sumifs(fit, "J", "A", start, end)
        retail_hotel = sumifs(fit, "M", "A", start, end)
        retail_tour = sumifs(fit, "N", "A", start, end)
        retail_visa = sumifs(fit, cfg["visa"], "A", start, end)
        footfall = sumifs(fit, "Q", "A", start, end)

        fit_flight_bookings = sumifs(fit, "R", "A", start, end)
        fit_hotel_bookings = sumifs(fit, "V", "A", start, end)

    # HQ has no dedicated FIT sheet in the current workbook.
    # Its retail Tour and Visa values are residuals from the global
    # All Tour / All Visa figures after the other branches' contributions.
    if branch == "HQ":
        all_tour = SOURCE.tab("All Tour")
        all_visa = SOURCE.tab("All Visa")

        other_tour = sum(
            branch_period(b, start, end)["products"]["Tour"]["gmv"]
            for b in ["Banani", "Chattogram", "Motijheel"]
        )
        other_visa = sum(
            branch_period(b, start, end)["products"]["Visa"]["gmv"]
            for b in ["Banani", "Chattogram", "Motijheel"]
        )

        retail_tour = (
            sumifs(all_tour, "O", "A", start, end)
            - other_tour
            - css_tour
        )

        retail_visa = (
            sumifs(all_visa, "D", "A", start, end)
            - other_visa
        )

    products = {
        "Flight": {
            "css": css_flight,
            "retail": retail_flight,
            "gmv": css_flight + retail_flight,
            "receivable": css_flight_recv,
            "bookings": css_flight_bookings + fit_flight_bookings,
        },
        "Tour": {
            "css": css_tour,
            "retail": retail_tour,
            "gmv": css_tour + retail_tour,
            "receivable": css_tour_recv,
            "bookings": 0.0,
        },
        "Visa": {
            "css": 0.0,
            "retail": retail_visa,
            "gmv": retail_visa,
            "receivable": 0.0,
            "bookings": 0.0,
        },
        "Hotel": {
            "css": css_hotel,
            "retail": retail_hotel,
            "gmv": css_hotel + retail_hotel,
            "receivable": css_hotel_recv,
            "bookings": css_hotel_bookings + fit_hotel_bookings,
        },
        "Vouchers": {
            "css": 0.0,
            "retail": 0.0,
            "gmv": 0.0,
            "receivable": 0.0,
            "bookings": 0.0,
        },
        "Others": {
            "css": 0.0,
            "retail": 0.0,
            "gmv": 0.0,
            "receivable": 0.0,
            "bookings": 0.0,
        },
    }

    total = sum(p["gmv"] for p in products.values())

    targets = branch_month_targets(branch, end)

    return {
        "branch": branch,
        "start": start.isoformat(),
        "end": end.isoformat(),
        "currency": "USD",
        "products": products,
        "total_gmv": total,
        "targets": targets,
        "footfall": footfall,
    }


def branch_progress(branch: str, as_of: date) -> dict:
    mtd = branch_period(branch, month_start(as_of), as_of)
    targets = mtd["targets"]

    days_elapsed = as_of.day
    days_in_month = month_end(as_of).day

    projected = (
        mtd["total_gmv"] / days_elapsed * days_in_month
        if days_elapsed else 0
    )

    current_run_rate = (
        mtd["total_gmv"] / days_elapsed
        if days_elapsed else 0
    )

    monthly_target = sum(targets.values())
    remaining_days = days_in_month - days_elapsed

    required = (
        max(
            0.0,
            (
                monthly_target -
                mtd["total_gmv"]
            ) / remaining_days
        )
        if remaining_days > 0
        else 0
    )

    cfg = BRANCH_SOURCE[branch]
    dv = SOURCE.tab("DV")
    active_css = count_nonempty(dv, cfg["dv"])

    result = {
        **mtd,
        "projected": projected,
        "current_run_rate": current_run_rate,
        "required_run_rate": required,
        "gap": current_run_rate - required,
        "active_css": active_css,
        "pipeline_number": 5 if branch != "HQ" else 0,
        "pipeline_worth": 25000 if branch != "HQ" else 0,
    }

    return result


def h2_performance() -> list[dict]:
    """
    H2 is intentionally derived from the same data families:
      Flight/Hotel = branch CSS + branch FIT
      Tour/Visa = All Tour / All Visa
    Targets come from the locked Dashboard values.
    """
    dash = SOURCE.tab("Dashboard")

    q3_target = {
        "Flight": target_value(dash, 5, "B"),
        "Hotel": target_value(dash, 6, "B"),
        "Tour": target_value(dash, 7, "B"),
        "Visa": target_value(dash, 8, "B"),
    }

    h2_target = {
        "Flight": target_value(dash, 5, "H"),
        "Hotel": target_value(dash, 6, "H"),
        "Tour": target_value(dash, 7, "H"),
        "Visa": target_value(dash, 8, "H"),
    }

    q3 = {}
    q4 = {}

    for product in ["Flight", "Hotel"]:
        q3[product] = sum(
            branch_period(
                b,
                q3_start(),
                q3_end()
            )["products"][product]["gmv"]
            for b in BRANCH_SOURCE
        )

        q4[product] = sum(
            branch_period(
                b,
                q4_start(),
                q4_end()
            )["products"][product]["gmv"]
            for b in BRANCH_SOURCE
        )

    tour = SOURCE.tab("All Tour")
    visa = SOURCE.tab("All Visa")

    q3["Tour"] = sumifs(tour, "O", "A", q3_start(), q3_end())
    q4["Tour"] = sumifs(tour, "O", "A", q4_start(), q4_end())

    q3["Visa"] = sumifs(visa, "D", "A", q3_start(), q3_end())
    q4["Visa"] = sumifs(visa, "D", "A", q4_start(), q4_end())

    rows = []

    for product in ["Flight", "Hotel", "Tour", "Visa"]:
        q4_target = h2_target[product] - q3_target[product]

        q3_actual = q3[product]
        q4_actual = q4[product]
        h2_actual = q3_actual + q4_actual

        rows.append({
            "product": product,
            "q3_target": q3_target[product],
            "q3_actual": q3_actual,
            "q3_achievement": pct(q3_actual, q3_target[product]),
            "q4_target": q4_target,
            "q4_actual": q4_actual,
            "q4_achievement": pct(q4_actual, q4_target),
            "h2_target": h2_target[product],
            "h2_actual": h2_actual,
            "h2_achievement": pct(h2_actual, h2_target[product]),
        })

    total = {
        "product": "Overall — All Products",
        "q3_target": sum(r["q3_target"] for r in rows),
        "q3_actual": sum(r["q3_actual"] for r in rows),
        "q4_target": sum(r["q4_target"] for r in rows),
        "q4_actual": sum(r["q4_actual"] for r in rows),
        "h2_target": sum(r["h2_target"] for r in rows),
        "h2_actual": sum(r["h2_actual"] for r in rows),
    }

    total["q3_achievement"] = pct(
        total["q3_actual"],
        total["q3_target"]
    )

    total["q4_achievement"] = pct(
        total["q4_actual"],
        total["q4_target"]
    )

    total["h2_achievement"] = pct(
        total["h2_actual"],
        total["h2_target"]
    )

    return rows + [total]


def wbr_targets() -> dict:
    t = SOURCE.tab("Target Helper")

    def v(row, col):
        return target_value(t, row, col)

    # August WBR target model from the current workbook.
    return {
        "Flight": (
            v(19, "I") + v(19, "J") +
            v(45, "I") + v(45, "J") +
            v(73, "I") + v(73, "J") +
            v(102, "I") + v(102, "J")
        ) / 4,

        "Tour": (
            v(21, "I") + v(21, "J") +
            v(22, "I") + v(22, "J") +
            v(47, "I") + v(47, "J") +
            v(48, "I") + v(48, "J") +
            v(75, "I") + v(75, "J") +
            v(76, "I") + v(76, "J") +
            v(104, "I") + v(104, "J") +
            v(105, "I") + v(105, "J")
        ) / 4,

        "Visa": (
            v(23, "I") + v(23, "J") +
            v(49, "I") + v(49, "J") +
            v(77, "I") + v(77, "J") +
            v(106, "I") + v(106, "J")
        ) / 4,

        "Hotel": (
            v(20, "I") + v(20, "J") +
            v(46, "I") + v(46, "J") +
            v(74, "I") + v(74, "J") +
            v(103, "I") + v(103, "J")
        ) / 4,
    }


def wbr_metrics(week_end: date) -> dict:
    monday = week_end - timedelta(days=week_end.weekday())
    sunday = monday + timedelta(days=6)

    prev_monday = monday - timedelta(days=7)
    prev_sunday = sunday - timedelta(days=7)

    def period(start, end):
        flight = sum(
            branch_period(
                b,
                start,
                end
            )["products"]["Flight"]["gmv"]
            for b in BRANCH_SOURCE
        )

        hotel = sum(
            branch_period(
                b,
                start,
                end
            )["products"]["Hotel"]["gmv"]
            for b in BRANCH_SOURCE
        )

        tour = sumifs(
            SOURCE.tab("All Tour"),
            "O",
            "A",
            start,
            end
        )

        visa = sumifs(
            SOURCE.tab("All Visa"),
            "D",
            "A",
            start,
            end
        )

        # Bookings:
        flight_bookings = sum(
            sum(
                [
                    sumifs(SOURCE.tab(cfg["css"]), "O", "A", start, end),
                    sumifs(
                        SOURCE.tab(cfg["fit"]), "R", "A", start, end
                    ) if cfg["fit"] else 0,
                ]
            )
            for cfg in BRANCH_SOURCE.values()
        )

        hotel_bookings = sum(
            sum(
                [
                    sumifs(SOURCE.tab(cfg["css"]), "P", "A", start, end),
                    sumifs(
                        SOURCE.tab(cfg["fit"]), "V", "A", start, end
                    ) if cfg["fit"] else 0,
                ]
            )
            for cfg in BRANCH_SOURCE.values()
        )

        tour_bookings = sumifs(
            SOURCE.tab("All Tour"),
            "L",
            "A",
            start,
            end
        )

        visa_bookings = sumifs(
            SOURCE.tab("All Visa"),
            "F",
            "A",
            start,
            end
        )

        # Receivable
        flight_recv = sum(
            sumifs(
                SOURCE.tab(cfg["css"]),
                "L",
                "A",
                start,
                end
            )
            for cfg in BRANCH_SOURCE.values()
        )

        tour_recv = sum(
            sumifs(
                SOURCE.tab(cfg["css"]),
                "N",
                "A",
                start,
                end
            )
            for cfg in BRANCH_SOURCE.values()
        )

        hotel_recv = sum(
            sumifs(
                SOURCE.tab(cfg["css"]),
                "M",
                "A",
                start,
                end
            )
            for cfg in BRANCH_SOURCE.values()
        )

        # Net revenue follows the current WBR workbook sources.
        tour_net = sumifs(
            SOURCE.tab("All Tour"),
            "I",
            "A",
            start,
            end
        )

        visa_net = sumifs(
            SOURCE.tab("All Visa"),
            "L",
            "A",
            start,
            end
        )

        # WBR source has a fixed Hotel net revenue value.
        hotel_net = 83.5

        return {
            "flight": flight,
            "tour": tour,
            "visa": visa,
            "hotel": hotel,
            "gmv": flight + tour + visa + hotel,

            "flight_bookings": flight_bookings,
            "tour_bookings": tour_bookings,
            "visa_bookings": visa_bookings,
            "hotel_bookings": hotel_bookings,
            "bookings": (
                flight_bookings +
                tour_bookings +
                visa_bookings +
                hotel_bookings
            ),

            "flight_recv": flight_recv,
            "tour_recv": tour_recv,
            "hotel_recv": hotel_recv,
            "receivable": (
                flight_recv +
                tour_recv +
                hotel_recv
            ),

            "tour_net": tour_net,
            "visa_net": visa_net,
            "hotel_net": hotel_net,
            "net_revenue": (
                tour_net +
                visa_net +
                hotel_net
            ),
        }

    cur = period(monday, sunday)
    prev = period(prev_monday, prev_sunday)

    return {
        "start": monday.isoformat(),
        "end": sunday.isoformat(),

        "summary": {
            "gmv": {
                "Flight": cur["flight"],
                "Tour": cur["tour"],
                "Visa": cur["visa"],
                "Hotel": cur["hotel"],
                "Total": cur["gmv"],
                "WoW": change(cur["gmv"], prev["gmv"]),
            },

            "bookings": {
                "Flight": cur["flight_bookings"],
                "Tour": cur["tour_bookings"],
                "Visa": cur["visa_bookings"],
                "Hotel": cur["hotel_bookings"],
                "Total": cur["bookings"],
                "WoW": change(cur["bookings"], prev["bookings"]),
            },

            "receivable": {
                "Flight": cur["flight_recv"],
                "Tour": cur["tour_recv"],
                "Visa": 0.0,
                "Hotel": cur["hotel_recv"],
                "Total": cur["receivable"],
                "Ratio": pct(cur["receivable"], cur["gmv"]),
            },

            "net_revenue": {
                "Flight": 0.0,
                "Tour": cur["tour_net"],
                "Visa": cur["visa_net"],
                "Hotel": cur["hotel_net"],
                "Total": cur["net_revenue"],
                "Ratio": pct(cur["net_revenue"], cur["gmv"]),
            },
        },

        "weekly_targets": wbr_targets(),

        "h2": h2_performance(),

        "footfall": sum(
            branch_period(b, monday, sunday)["footfall"]
            for b in BRANCH_SOURCE
        ),

        "previous": {
            "gmv": prev["gmv"],
            "bookings": prev["bookings"],
            "receivable": prev["receivable"],
            "net_revenue": prev["net_revenue"],
        },

        "narrative": wbr_narrative(),
    }


def wbr_narrative() -> dict:
    df = SOURCE.tab("WBR Format")

    def rows(start, end, cols):
        result = []
        for r in range(start, end + 1):
            values = []
            for c in cols:
                idx = col_idx(c)
                row_idx = r - 2
                if (
                    0 <= row_idx < len(df)
                    and idx < df.shape[1]
                ):
                    v = df.iloc[row_idx, idx]
                    if pd.notna(v) and str(v).strip():
                        values.append(str(v))
            if values:
                result.append(values)
        return result

    return {
        "analysis": rows(31, 36, ["A", "C"]),
        "challenges": rows(39, 43, ["A", "B", "G", "I"]),
        "actions": rows(46, 51, ["A", "B", "G", "H", "I", "K"]),
    }


# ============================================================
# API
# ============================================================

@asynccontextmanager
async def lifespan(app: FastAPI):
    task = asyncio.create_task(auto_refresh_loop())
    yield
    task.cancel()
    try:
        await task
    except asyncio.CancelledError:
        pass


async def auto_refresh_loop():
    while True:
        await asyncio.sleep(REFRESH_MINUTES * 60)
        try:
            SOURCE.refresh()
        except Exception as exc:
            SOURCE.last_error = str(exc)


app = FastAPI(
    title="GoZayaan Retail Dashboard",
    lifespan=lifespan,
)

app.mount(
    "/static",
    StaticFiles(directory=str(BASE_DIR / "static")),
    name="static",
)


@app.get("/")
def index():
    return FileResponse(BASE_DIR / "static" / "index.html")


@app.get("/api/health")
def health():
    return {
        "source": SOURCE.source_name,
        "last_refresh": SOURCE.last_refresh.isoformat()
        if SOURCE.last_refresh else None,
        "last_error": SOURCE.last_error,
        "refresh_minutes": REFRESH_MINUTES,
        "currency": "USD",
    }


@app.post("/api/refresh")
def refresh():
    started = time.perf_counter()

    try:
        SOURCE.refresh()
    except Exception as exc:
        raise HTTPException(
            status_code=502,
            detail=str(exc),
        )

    return {
        "ok": True,
        "source": SOURCE.source_name,
        "refreshed_at": SOURCE.last_refresh.isoformat()
        if SOURCE.last_refresh else None,
        "seconds": round(
            time.perf_counter() - started,
            2
        ),
    }


@app.get("/api/dashboard")
def api_dashboard():
    data = dashboard_h2()
    return JSONResponse({
        **data,
        "source": SOURCE.source_name,
        "last_refresh": SOURCE.last_refresh.isoformat()
        if SOURCE.last_refresh else None,
    })


@app.get("/api/h2")
def api_h2():
    return {
        "page": "h2",
        "currency": "USD",
        "rows": h2_performance(),
        "source": SOURCE.source_name,
        "last_refresh": SOURCE.last_refresh.isoformat()
        if SOURCE.last_refresh else None,
    }


@app.get("/api/branch/{branch}")
def api_branch(
    branch: str,
    start: str,
    end: str,
):
    canonical = {
        "banani": "Banani",
        "chattogram": "Chattogram",
        "hq": "HQ",
        "motijheel": "Motijheel",
    }.get(branch.lower())

    if not canonical:
        raise HTTPException(
            status_code=404,
            detail="Unknown branch.",
        )

    try:
        start_d = parse_date(start)
        end_d = parse_date(end)
    except ValueError as exc:
        raise HTTPException(
            status_code=400,
            detail=str(exc),
        )

    if start_d > end_d:
        raise HTTPException(
            status_code=400,
            detail="Start date cannot be after end date.",
        )

    data = branch_period(
        canonical,
        start_d,
        end_d,
    )

    # Operational/month progress intentionally uses the selected end date.
    progress = branch_progress(
        canonical,
        end_d,
    )

    return {
        "page": "branch",
        "branch": canonical,
        "currency": "USD",
        "start": start,
        "end": end,
        "total_gmv": data["total_gmv"],
        "products": data["products"],
        "progress": {
            "mtd_gmv": progress["total_gmv"],
            "projected": progress["projected"],
            "current_run_rate": progress["current_run_rate"],
            "required_run_rate": progress["required_run_rate"],
            "gap": progress["gap"],
            "active_css": progress["active_css"],
            "pipeline_number": progress["pipeline_number"],
            "pipeline_worth": progress["pipeline_worth"],
            "footfall": data["footfall"],
            "targets": progress["targets"],
        },
        "source": SOURCE.source_name,
        "last_refresh": SOURCE.last_refresh.isoformat()
        if SOURCE.last_refresh else None,
    }


@app.get("/api/wbr")
def api_wbr(week_end: str):
    try:
        end = parse_date(week_end)
    except ValueError as exc:
        raise HTTPException(
            status_code=400,
            detail=str(exc),
        )

    return {
        **wbr_metrics(end),
        "page": "wbr",
        "currency": "USD",
        "source": SOURCE.source_name,
        "last_refresh": SOURCE.last_refresh.isoformat()
        if SOURCE.last_refresh else None,
    }


if __name__ == "__main__":
    uvicorn.run(
        "app:app",
        host="0.0.0.0",
        port=int(os.getenv("PORT", "8000")),
        reload=False,
    )
