"""
Loads the source-adapter portion of app.py in isolation (up to SOURCE = ...)
and exercises the three failure paths the patches target.
"""
import re
import sys
import types
import requests
import pandas as pd
from fastapi import HTTPException

src = open("app.py").read()

# Take everything up to the module-level SOURCE instantiation so we can
# construct SourceManager ourselves with a mocked transport.
cut = src.index("SOURCE = SourceManager()")
head = src[:cut]

mod = types.ModuleType("appcore")
mod.__dict__["__file__"] = "app.py"
exec(compile(head, "app.py", "exec"), mod.__dict__)

AppsScriptSource = mod.AppsScriptSource
SourceManager = mod.SourceManager

# ---- mock Apps Script transport -------------------------------------
class FakeResponse:
    def __init__(self, status, payload=None, text=""):
        self.status_code = status
        self._payload = payload
        self.text = text

    def raise_for_status(self):
        if self.status_code >= 400:
            raise requests.HTTPError(
                f"{self.status_code} Client Error: Not Found for url: "
                "https://script.googleusercontent.com/macros/echo?user_content_key=XXX"
            )

    def json(self):
        if self._payload is None:
            raise ValueError("no json")
        return self._payload


GOOD = {"ok": True, "rows": [["hdr"], ["a"], ["b"]]}

calls = {}


def make_transport(fail_tabs, fail_times):
    """fail_tabs fail `fail_times` times each, then succeed."""
    def _get(url, params=None, timeout=None, allow_redirects=None, headers=None):
        name = params["sheet"]
        calls[name] = calls.get(name, 0) + 1
        if name in fail_tabs and calls[name] <= fail_times:
            return FakeResponse(404)
        return FakeResponse(200, GOOD)
    return _get


TABS = SourceManager.REQUIRED_TABS
results = []


def check(label, cond):
    results.append((label, cond))
    print(("PASS  " if cond else "FAIL  ") + label)


# --- Test 1: transient 404 on 2 tabs, recovers on retry --------------
calls.clear()
mod.requests.get = make_transport({"Target Helper", "DS1Banani CSS"}, fail_times=1)
s = AppsScriptSource("http://fake")
errs = s.warm(TABS)
check("T1 transient 404 recovers via retry -> no errors", errs == [])
check("T1 all 13 tabs cached", s.cached_tabs == len(TABS))

# --- Test 2: permanent failure on 2 tabs, other 11 survive -----------
calls.clear()
mod.requests.get = make_transport({"Target Helper", "DS1Banani CSS"}, fail_times=99)
s2 = AppsScriptSource("http://fake")
errs2 = s2.warm(TABS)
check("T2 two tabs report errors", len(errs2) == 2)
check("T2 the other 11 tabs are kept (old code discarded all)",
      s2.cached_tabs == len(TABS) - 2)
check("T2 error text names the failing tab",
      any("Target Helper" in e for e in errs2))

# --- Test 3: SourceManager keeps partial source ----------------------
calls.clear()
mod.LIVE_REQUIRED = True
mod.APPS_SCRIPT_URL = "http://fake"
mod.requests.get = make_transport({"Target Helper", "DS1Banani CSS"}, fail_times=99)
sm = SourceManager()
check("T3 source is NOT None on partial failure", sm.source is not None)
check("T3 marked degraded", sm.degraded is True)
check("T3 source_name flags PARTIAL", "PARTIAL" in sm.source_name)
check("T3 a good tab still serves data", isinstance(sm.tab("DV"), pd.DataFrame))

# --- Test 4: total outage, then recovery, keeps last good data -------
calls.clear()
mod.requests.get = make_transport(set(TABS), fail_times=99)
sm2 = SourceManager()
check("T4 cold total failure -> source None", sm2.source is None)
check("T4 cold total failure -> UNAVAILABLE", "UNAVAILABLE" in sm2.source_name)

try:
    sm2.tab("DV")
    check("T4 tab() raises HTTPException 503", False)
except HTTPException as e:
    check("T4 tab() raises HTTPException 503 (not AttributeError)",
          e.status_code == 503)
except AttributeError:
    check("T4 tab() raises HTTPException 503 (not AttributeError)", False)

# now recover, then fail again -> must serve stale
calls.clear()
mod.requests.get = make_transport(set(), fail_times=0)
sm2.refresh()
check("T4 recovers to healthy", sm2.source is not None and not sm2.degraded)
good_refresh = sm2.last_refresh

calls.clear()
mod.requests.get = make_transport(set(TABS), fail_times=99)
sm2.refresh()
check("T4 later total failure keeps serving stale data",
      sm2.source is not None)
check("T4 stale flagged in source_name", "STALE" in sm2.source_name)
check("T4 last_refresh preserved so UI can show data age",
      sm2.last_refresh == good_refresh)

# --- Test 5: batch mode ----------------------------------------------
import os
mod.BATCH_FETCH = True

def batch_transport(ok=True, per_tab_errors=None):
    def _get(url, params=None, timeout=None, allow_redirects=None, headers=None):
        if "sheets" in params:
            names = params["sheets"].split(",")
            if not ok:
                return FakeResponse(404)
            sheets = {}
            for n in names:
                if per_tab_errors and n in per_tab_errors:
                    continue
                sheets[n] = [["hdr"], ["a"], ["b"]]
            return FakeResponse(200, {
                "ok": True,
                "sheets": sheets,
                "errors": {n: "Sheet not found" for n in (per_tab_errors or [])},
            })
        return FakeResponse(200, GOOD)
    return _get

calls.clear()
mod.requests.get = batch_transport()
s5 = AppsScriptSource("http://fake")
e5 = s5.warm(TABS)
check("T5 batch mode fetches all tabs in one request", e5 == [] and s5.cached_tabs == len(TABS))

calls.clear()
mod.requests.get = batch_transport(ok=False)
s6 = AppsScriptSource("http://fake")
e6 = s6.warm(TABS)
check("T5 batch failure falls back to per-tab (still gets data)",
      s6.cached_tabs == len(TABS))
check("T5 fallback is recorded in errors",
      any("fell back" in e for e in e6))

mod.BATCH_FETCH = False
print()
failed2 = [l for l, c in results if not c]
print(f"FINAL {len(results) - len(failed2)}/{len(results)} passed")
sys.exit(1 if failed2 else 0)
