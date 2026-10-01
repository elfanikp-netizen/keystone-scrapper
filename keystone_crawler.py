#!/usr/bin/env python3
"""
Keystone (preview.orderkeystone.com/crash) Partslink crawler.

Reads Partslink numbers from an Excel file, logs in once, searches each number
on the site and writes these columns to an Excel file:

    Oldest Year, Newest Year, Brand, Model, Type, Interchange Number, OEM Number,
    Interchange 1, Interchange 2, Interchange 3, Interchange 4, Interchange 5,
    OEM 1, OEM 2, OEM 3, OEM 4, OEM 5

Typical use:
    python keystone_crawler.py --make-template          # creates parts.xlsx
    python keystone_crawler.py --discover <ONE PART NO>  # first test, opens browser
    python keystone_crawler.py --show --limit 5          # small visible test run
    python keystone_crawler.py                           # full run (headless)

Login credentials are read from a .env file (KEYSTONE_USERNAME / KEYSTONE_PASSWORD)
or typed in when the script asks for them.
"""

import argparse
import getpass
import json
import os
import random
import re
import subprocess
import sys
import time
from pathlib import Path

try:
    import pandas as pd
except ImportError:  # handled in main()
    pd = None

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout
except ImportError:  # handled in main()
    sync_playwright = None
    PWTimeout = Exception


def ensure_runtime_dependencies():
    """Validate the environment before the crawler starts."""
    if pd is None:
        raise RuntimeError(
            "Missing Python dependencies. Run:\n"
            "  python -m pip install -r requirements.txt\n"
            "  python -m playwright install chromium"
        )
    try:
        import openpyxl  # noqa: F401
    except ImportError as exc:
        raise RuntimeError(
            "Missing Excel support. Run:\n"
            "  python -m pip install -r requirements.txt"
        ) from exc
    if sync_playwright is None:
        raise RuntimeError(
            "Playwright is not installed. Run:\n"
            "  python -m pip install -r requirements.txt\n"
            "  python -m playwright install chromium"
        )


def ensure_input_file(path):
    """Ensure the parts Excel file exists before parsing it."""
    if not Path(path).exists():
        raise FileNotFoundError(
            f"Input file '{path}' was not found. Create it with:\n"
            "  python keystone_crawler.py --make-template\n"
            "or pass a different --input path."
        )


def ensure_browser_available():
    """Download Chromium if Playwright has not installed the browser yet."""
    try:
        with sync_playwright() as p:
            p.chromium.launch(headless=True)
        return
    except Exception as exc:  # noqa: BLE001
        msg = str(exc)
        if "Executable doesn't exist" not in msg and "Please run the following command" not in msg:
            raise
        print("Playwright Chromium is missing. Downloading it now...")
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=False)
        try:
            with sync_playwright() as p:
                p.chromium.launch(headless=True)
        except Exception as inner:  # noqa: BLE001
            raise RuntimeError(
                "Playwright browser could not be launched. Run:\n"
                "  python -m playwright install chromium"
            ) from inner

# --------------------------------------------------------------------------- #
# CONFIG - the only part you should ever need to edit
# --------------------------------------------------------------------------- #
BASE_URL = "https://preview.orderkeystone.com"
CRASH_URL = f"{BASE_URL}/crash"
STATE_FILE = "auth_state.json"          # saved login session (keep private)

CONFIG = {
    # Leave "" to auto-detect. Otherwise put a CSS selector, e.g. 'input[placeholder="Search"]'
    "search_input_selector": "",
    # Leave "" to just press Enter after typing. Otherwise CSS selector of the search button.
    "search_button_selector": "",
    # If the search shows a list and details are on a second page/panel, put the selector
    # of the first result to click, e.g. 'table tbody tr:first-child a'. Leave "" if not needed.
    "first_result_click_selector": "",
    # Optional: selector that appears when results are loaded, e.g. 'table tbody tr'.
    "result_ready_selector": "",
    # Extra wait (ms) after the page/network looks idle - raise it if the site is slow.
    "settle_ms": 1500,
    # Delay between searches (seconds) - be polite to the server.
    "delay_min": 1.5,
    "delay_max": 3.5,
    # Joins multiple interchange / OEM numbers found for one row.
    "list_join": ", ",
    # Retries per part number when something fails.
    "retries": 2,
}

FIELDS = ["Oldest Year", "Newest Year", "Brand", "Model", "Type",
          "Interchange Number", "OEM Number"]
INTERCHANGE_FIELDS = [f"Interchange {n}" for n in range(1, 6)]
OEM_FIELDS = [f"OEM {n}" for n in range(1, 6)]
NUMBER_VALUES_FIELD = "Number Values"
MULTIPLE_NUMBER_FIELD = "Multiple Number"
PARTSLINK_PATTERN = re.compile(r"(?<![A-Za-z0-9])[A-Za-z]{2}\d{7}(?![A-Za-z0-9])")

# Regexes used to recognise column headers / labels / JSON keys on the site.
FIELD_PATTERNS = {
    "Partslink Number": [r"parts\s*-?\s*link"],
    "Oldest Year": [
        r"(oldest|earliest|start|begin|from|min)\s*(year|yr)",
        r"(year|yr)\s*(start|from|begin|min)",
    ],
    "Newest Year": [
        r"(newest|latest|end|to|max)\s*(year|yr)",
        r"(year|yr)\s*(end|to|max)",
    ],
    "Brand": [r"(^|\s)(brand|make|manufacturer)$"],
    "Model": [r"(^|\s)model$"],
    "Type": [r"(^|\s)type$", r"^category$"],
    "Interchange Number": [r"interchange"],
    "OEM Number": [r"\boem\b"],
}
YEAR_RANGE_KEY = re.compile(r"(^|\s)(years?|year range|fitment|application)$")
YEAR_RANGE_VAL = re.compile(r"((?:19|20)\d{2})\s*(?:-|–|—|to|thru|through)\s*((?:19|20)\d{2})", re.I)

DEBUG_DIR = Path("debug")

# --------------------------------------------------------------------------- #
# Helpers: field matching
# --------------------------------------------------------------------------- #
def norm_key(k):
    k = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", str(k))
    k = re.sub(r"[_\-]+", " ", k)
    return re.sub(r"\s+", " ", k).strip().lower().rstrip(":")


def match_field(key):
    nk = norm_key(key)
    for field, pats in FIELD_PATTERNS.items():
        if any(re.search(p, nk) for p in pats):
            return field
    return None


def clean(v):
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


def extract_partslink_number(value):
    """Find a Partslink identifier containing two letters followed by seven digits."""
    if isinstance(value, dict):
        values = value.values()
    elif isinstance(value, (list, tuple)):
        values = value
    else:
        match = PARTSLINK_PATTERN.search(str(value))
        return match.group(0).upper() if match else ""

    for item in values:
        found = extract_partslink_number(item)
        if found:
            return found
    return ""


def normalize_number_tokens(value):
    """Return unique numbers from a free-form field while keeping a stable order."""
    tokens = []
    seen = set()
    if value is None:
        return tokens
    for piece in re.split(r"[;|,\/\n]+", str(value)):
        for token in re.findall(r"[A-Za-z0-9-]+", piece):
            token = token.strip()
            if not token:
                continue
            if PARTSLINK_PATTERN.fullmatch(token):
                continue
            key = token.upper()
            if key not in seen:
                seen.add(key)
                tokens.append(token)
    return tokens


def join_number_values(*values):
    """Join normalized OEM/interchange values without Partslink identifiers."""
    tokens = normalize_number_tokens(", ".join(str(value) for value in values if value))
    return CONFIG["list_join"].join(tokens)


def split_number_values(value, max_items=5):
    """Split a free-form number field into up to max_items unique values."""
    return normalize_number_tokens(value)[:max_items]


def apply_number_slots(row, interchange_value="", oem_value=""):
    """Populate the canonical and numbered OEM/interchange columns."""
    interchange_tokens = split_number_values(interchange_value)
    oem_tokens = split_number_values(oem_value)

    row["Interchange Number"] = interchange_tokens[0] if interchange_tokens else ""
    row["OEM Number"] = oem_tokens[0] if oem_tokens else ""

    for index in range(1, 6):
        row[f"Interchange {index}"] = interchange_tokens[index - 1] if index <= len(interchange_tokens) else ""
        row[f"OEM {index}"] = oem_tokens[index - 1] if index <= len(oem_tokens) else ""

    return row


def number_values_from_row(row):
    """Combine populated OEM and interchange columns, with OEM values first."""
    oem_values = [row.get("OEM Number", "")]
    oem_values.extend(row.get(f"OEM {index}", "") for index in range(1, 6))
    interchange_values = [row.get("Interchange Number", "")]
    interchange_values.extend(row.get(f"Interchange {index}", "") for index in range(1, 6))
    return join_number_values(*oem_values, *interchange_values)


def is_multiple_number(value):
    """Return True when a field contains more than one distinct number candidate."""
    return len(normalize_number_tokens(value)) > 1


def map_pairs(pairs):
    """[(key, value), ...] -> {field: value} using FIELD_PATTERNS."""
    rec = {}
    for key, val in pairs:
        val = clean(val)
        if not val:
            continue
        field = match_field(key)
        if field and field not in rec:
            rec[field] = val
            continue
        # A single "Years" column such as "2012-2016"
        if YEAR_RANGE_KEY.search(norm_key(key)):
            m = YEAR_RANGE_VAL.search(val)
            if m:
                rec.setdefault("Oldest Year", m.group(1))
                rec.setdefault("Newest Year", m.group(2))
    return rec


# --------------------------------------------------------------------------- #
# Extraction from JSON responses (best source when the site uses an API)
# --------------------------------------------------------------------------- #
def _scalars(d):
    out = []
    for k, v in d.items():
        if isinstance(v, (str, int, float, bool)):
            out.append((k, v))
        elif isinstance(v, list) and v and all(isinstance(x, (str, int, float)) for x in v):
            out.append((k, CONFIG["list_join"].join(str(x) for x in v)))
    return out


def _has_container(d):
    return any(isinstance(v, dict) or (isinstance(v, list) and any(isinstance(x, (dict, list)) for x in v))
               for v in d.values())


def walk_json(obj, inherited, out):
    """Collect leaf records; child records inherit scalar fields from their parents."""
    if isinstance(obj, list):
        for item in obj:
            walk_json(item, inherited, out)
    elif isinstance(obj, dict):
        ctx = dict(inherited)
        ctx.update(dict(_scalars(obj)))
        containers = [v for v in obj.values()
                      if isinstance(v, dict) or (isinstance(v, list) and any(isinstance(x, (dict, list)) for x in v))]
        for c in containers:
            walk_json(c, ctx, out)
        if not containers:
            rec = map_pairs(ctx.items())
            if len(rec) >= 2:
                out.append(rec)


# --------------------------------------------------------------------------- #
# Extraction from the rendered page (tables + "Label: value" text)
# --------------------------------------------------------------------------- #
DOM_JS = """
() => {
  const txt = el => (el.innerText || '').replace(/\\s+/g, ' ').trim();
  const out = {tables: [], dl: [], body: document.body ? document.body.innerText : ''};
  document.querySelectorAll('table').forEach(t => {
    let headers = [...t.querySelectorAll('thead th, thead td')].map(txt);
    let rows = [...t.querySelectorAll('tbody tr')];
    if (!headers.length) {
      const first = t.querySelector('tr');
      if (first) { headers = [...first.querySelectorAll('th,td')].map(txt);
                   rows = [...t.querySelectorAll('tr')].slice(1); }
    }
    const data = rows.map(r => [...r.querySelectorAll('th,td')].map(txt));
    if (headers.length && data.length) out.tables.push({headers, rows: data});
  });
  document.querySelectorAll('[role="table"], [role="grid"]').forEach(g => {
    const headers = [...g.querySelectorAll('[role="columnheader"]')].map(txt);
    const data = [...g.querySelectorAll('[role="row"]')]
      .map(r => [...r.querySelectorAll('[role="gridcell"], [role="cell"]')].map(txt))
      .filter(r => r.length);
    if (headers.length && data.length) out.tables.push({headers, rows: data});
  });
  document.querySelectorAll('dt').forEach(dt => {
    const dd = dt.nextElementSibling;
    if (dd) out.dl.push([txt(dt), txt(dd)]);
  });
  return out;
}
"""


def page_level_pairs(dom):
    pairs = list(dom.get("dl", []))
    lines = [l.strip() for l in dom.get("body", "").splitlines() if l.strip()]
    for i, line in enumerate(lines):
        m = re.match(r"^([A-Za-z][A-Za-z #/]{1,30}):\s*(.+)$", line)
        if m:
            pairs.append((m.group(1), m.group(2)))
        elif match_field(line.rstrip(":")) and line.endswith(":") and i + 1 < len(lines):
            pairs.append((line.rstrip(":"), lines[i + 1]))
    return pairs


def dom_records(dom):
    recs = []
    for t in dom.get("tables", []):
        headers = t["headers"]
        for row in t["rows"]:
            rec = map_pairs(zip(headers, row))
            if len(rec) >= 2:
                recs.append(rec)
    return recs


def score(recs):
    return sum(len(r) for r in recs)


def is_interchange_search(search_column):
    return bool(re.search(r"\binterchange\b", str(search_column), re.I))


def extract(page, captured_json, part=None, search_column="Partslink Number"):
    json_recs = []
    for _, body in captured_json:
        walk_json(body, {}, json_recs)

    dom = page.evaluate(DOM_JS)
    interchange_search = is_interchange_search(search_column)
    parsed_partslink = "" if interchange_search else extract_partslink_number([
        dom.get("body", ""),
        [body for _, body in captured_json],
    ])
    table_recs = dom_records(dom)
    page_rec = map_pairs(page_level_pairs(dom))
    page_partslink = extract_partslink_number(page_rec)

    if not table_recs and len(page_rec) >= 2:
        table_recs = [dict(page_rec)]

    best = json_recs if score(json_recs) >= score(table_recs) else table_recs

    rows, seen = [], set()
    oem_values = extract_oem_tab_values(
        page,
        part,
        exclude_search_value=not interchange_search,
    )
    for rec in best:
        for f in FIELDS:                      # fill gaps from page-level labels
            if f not in rec and f in page_rec:
                rec[f] = page_rec[f]
        row = {f: rec.get(f, "") for f in FIELDS}

        interchange_value = (
            oem_values.get("interchange")
            or row.get("Interchange Number", "")
            or (part if interchange_search else "")
        )
        oem_value = oem_values.get("oem") or row.get("OEM Number", "")
        row = apply_number_slots(row, interchange_value, oem_value)
        row[NUMBER_VALUES_FIELD] = number_values_from_row(row)
        row[MULTIPLE_NUMBER_FIELD] = bool(oem_values.get("multiple")) or (
            is_multiple_number(interchange_value) or is_multiple_number(oem_value)
        )
        row["Partslink Number"] = (
            (oem_values.get("partslink") if interchange_search else "")
            or extract_partslink_number(rec.get("Partslink Number", ""))
            or (page_partslink if interchange_search else "")
            or (extract_partslink_number(rec) if not interchange_search else "")
            or parsed_partslink
        )

        key = tuple(row.values())
        if any(key) and key not in seen:
            seen.add(key)
            rows.append(row)

    if not rows and any(oem_values.values()):
        fallback = {f: "" for f in FIELDS}
        for field in INTERCHANGE_FIELDS + OEM_FIELDS:
            fallback[field] = ""
        interchange_value = oem_values.get("interchange", "") or (
            part if interchange_search else ""
        )
        oem_value = oem_values.get("oem", "")
        fallback = apply_number_slots(fallback, interchange_value, oem_value)
        fallback[NUMBER_VALUES_FIELD] = number_values_from_row(fallback)
        fallback[MULTIPLE_NUMBER_FIELD] = bool(oem_values.get("multiple"))
        fallback["Partslink Number"] = (
            oem_values.get("partslink") or page_partslink or parsed_partslink
        )
        rows.append(fallback)

    if not rows and (page_partslink or parsed_partslink):
        fallback = {field: "" for field in FIELDS + INTERCHANGE_FIELDS + OEM_FIELDS}
        fallback[NUMBER_VALUES_FIELD] = ""
        fallback[MULTIPLE_NUMBER_FIELD] = False
        fallback["Partslink Number"] = (
            oem_values.get("partslink") or page_partslink or parsed_partslink
        )
        rows.append(fallback)

    return rows


def pick_best_interchange(value):
    """Choose the best interchange number: xxx-xxxxx or xxx-xxxxxx; prefer alphabetic prefix and highest suffix."""
    if not value:
        return ""
    matches = []
    for piece in re.split(r"[;|,/\n]+", str(value)):
        for token in re.findall(r"[A-Za-z0-9-]+", piece):
            token = token.strip()
            if not token:
                continue
            m = re.search(r"(?i)([A-Za-z0-9]{2,8})-(\d{5,7})", token)
            if m:
                alpha = m.group(1).upper()
                digits = int(m.group(2))
                matches.append((alpha, digits, token))
    if not matches:
        return ""
    best = sorted(matches, key=lambda x: (x[0], x[1]))[-1]
    return best[2]


def pick_best_oem(value, part=None):
    """Take the best OEM value while preserving a valid same-as-part value when it is the only candidate."""
    if not value:
        return ""

    cleaned, part_match = [], None
    seen = set()
    for piece in re.split(r"[;|,/\n]+", str(value)):
        part_text = re.sub(r"\s+", " ", piece).strip()
        if not part_text:
            continue
        for token in re.findall(r"[A-Za-z0-9-]+", part_text):
            token = token.strip()
            if not token:
                continue
            if part and token.upper() == str(part).upper():
                part_match = token
                continue
            if token not in seen:
                seen.add(token)
                cleaned.append(token)

    if cleaned:
        return ", ".join(cleaned)
    if part_match:
        return part_match
    return ""


def is_oem_candidate(token):
    """Accept common OEM identifiers that contain letters and digits in either order."""
    if not token:
        return False
    token = token.strip()
    if token.upper() in {"OEM", "INTERCHANGE", "DETAILS", "FITS", "NUMBER", "PART"}:
        return False
    if token.isdigit() and len(token) < 5:
        return False
    if re.fullmatch(r"\d{5,20}", token):
        return True
    if not re.fullmatch(r"(?i)[A-Za-z0-9][A-Za-z0-9-_ ]{4,19}", token):
        return False
    return bool(re.search(r"[A-Za-z]", token) and re.search(r"\d", token))


def extract_oem_tab_values(page, part=None, exclude_search_value=True):
    """Read the visible values from the OEM tab wrapper and normalize them."""
    open_oem_tab(page)
    try:
        bodies = page.locator('.mat-tab-body-wrapper .mat-tab-body')
        if bodies.count() == 0:
            return {"interchange": "", "oem": "", "partslink": "", "multiple": False}
        active = page.locator('.mat-tab-body-wrapper .mat-tab-body-active')
        candidates = active if active.count() else bodies
        tab_text = ""
        for index in range(candidates.count()):
            tab_text = candidates.nth(index).inner_text() or ""
            if tab_text.strip():
                break
    except Exception:
        return {"interchange": "", "oem": "", "partslink": "", "multiple": False}

    if not tab_text:
        return {"interchange": "", "oem": "", "partslink": "", "multiple": False}

    print(f"Reading OEM tab values: {tab_text[:250]}")
    partslink = extract_partslink_number(tab_text)
    interchange_candidates = []
    oem_candidates = []
    for token in re.findall(r"[A-Za-z0-9-]+", tab_text):
        token = token.strip()
        if not token:
            continue
        if PARTSLINK_PATTERN.fullmatch(token):
            continue
        if exclude_search_value and part and token.upper() == str(part).upper():
            continue
        if re.search(r"(?i)(?:[A-Za-z0-9]{2,8})-\d{5,7}", token):
            interchange_candidates.append(token)
        elif is_oem_candidate(token):
            oem_candidates.append(token)

    interchange_tokens = normalize_number_tokens(", ".join(interchange_candidates))
    oem_tokens = normalize_number_tokens(", ".join(oem_candidates))
    interchange = ", ".join(interchange_tokens) if len(interchange_tokens) > 1 else (interchange_tokens[0] if interchange_tokens else "")
    oem = ", ".join(oem_tokens) if len(oem_tokens) > 1 else (oem_tokens[0] if oem_tokens else "")
    number_values = join_number_values(interchange, oem)
    multiple = len(interchange_tokens) > 1 or len(oem_tokens) > 1
    print(f"Parsed OEM values -> Interchange: {interchange}, OEM: {oem}, Number Values: {number_values}, Multiple: {multiple}")
    return {
        "interchange": interchange,
        "oem": oem,
        "partslink": partslink,
        "number_values": number_values,
        "multiple": multiple,
    }


def open_oem_tab(page):
    """Switch to an OEM detail tab if the site uses tabs for the part details."""
    selectors = [
        'button:has-text("OEM")', 'a:has-text("OEM")', '[role="tab"]:has-text("OEM")',
        'button:has-text("Interchange")', 'button:has-text("Fits")', 'button:has-text("Details")'
    ]
    for sel in selectors:
        try:
            tab = page.locator(sel).first
            if tab and tab.is_visible():
                tab.click(timeout=5000)
                wait_settled(page, 1000)
                return True
        except Exception:
            pass
    return False


# --------------------------------------------------------------------------- #
# Browser helpers
# --------------------------------------------------------------------------- #
class Recorder:
    """Remembers every JSON (xhr/fetch) response the page receives."""

    def __init__(self, page):
        self.items = []
        page.on("response", self._on)

    def _on(self, resp):
        try:
            ct = resp.headers.get("content-type") or ""
            if resp.request.resource_type in ("xhr", "fetch") and "json" in ct:
                self.items.append((resp.url, resp.json()))
        except Exception:
            pass

    def clear(self):
        self.items = []


def wait_settled(page, extra=None):
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except PWTimeout:
        pass
    page.wait_for_timeout(CONFIG["settle_ms"] if extra is None else extra)


def find_visible(page, selectors, limit=8):
    for sel in selectors:
        loc = page.locator(sel)
        try:
            n = loc.count()
        except Exception:
            continue
        for i in range(min(n, limit)):
            el = loc.nth(i)
            try:
                if el.is_visible():
                    return el
            except Exception:
                pass
    return None


def load_env(path=".env"):
    p = Path(path)
    if not p.exists():
        return
    for line in p.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#") and "=" in line:
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def get_credentials():
    load_env()
    user = os.environ.get("KEYSTONE_USERNAME") or input("Keystone username/email: ").strip()
    pw = os.environ.get("KEYSTONE_PASSWORD") or getpass.getpass("Keystone password: ")
    return user, pw


def needs_login(page):
    url = page.url.lower()
    if "/crash" in url and "orderkeystone" in url:
        try:
            body = (page.locator('body').inner_text() or '').lower()
            if "log in" not in body and "sign in" not in body and "password" not in body:
                return False
        except Exception:
            pass

    # Some authenticated pages still render a generic text box. If the crash page is already
    # loaded and no login form is actually present, there is no reason to re-enter credentials.
    has_user = find_visible(page, [
        'input[type="email"]', 'input[name*="user" i]', 'input[name*="email" i]',
        'input[id*="user" i]', 'input[id*="email" i]', 'input[autocomplete="username"]',
    ]) is not None
    has_password = find_visible(page, ['input[type="password"]']) is not None
    return has_user or has_password


def save_debug(page, name):
    DEBUG_DIR.mkdir(exist_ok=True)
    try:
        (DEBUG_DIR / f"{name}.html").write_text(page.content(), encoding="utf-8")
        page.screenshot(path=str(DEBUG_DIR / f"{name}.png"), full_page=True)
    except Exception:
        pass


def do_login(page):
    user, pw = get_credentials()

    userbox = find_visible(page, [
        'input[type="email"]', 'input[name*="user" i]', 'input[name*="email" i]',
        'input[id*="user" i]', 'input[id*="email" i]', 'input[autocomplete="username"]',
        'input[type="text"]',
    ])

    if userbox:
        print("Login step 1/2: entering username/email...")
        userbox.fill(user)
        next_btn = find_visible(page, [
            'button[type="submit"]', 'button:has-text("Continue")', 'button:has-text("Next")',
            'button:has-text("Submit")', 'button:has-text("Log in")', 'button:has-text("Login")',
            'button:has-text("Sign in")', 'input[type="submit"]',
        ])
        if next_btn:
            next_btn.click()
        else:
            userbox.press("Enter")
        wait_settled(page, 2500)

    pwd = find_visible(page, ['input[type="password"]'])
    if not pwd:
        return False

    print("Login step 2/2: entering password...")
    pwd.fill(pw)
    btn = find_visible(page, [
        'button[type="submit"]', 'button:has-text("Log in")', 'button:has-text("Login")',
        'button:has-text("Sign in")', 'button:has-text("Submit")', 'button:has-text("Continue")',
        'button:has-text("Next")', 'input[type="submit"]',
    ])
    if btn:
        btn.click()
    else:
        pwd.press("Enter")
    wait_settled(page, 2500)
    return True


def ensure_logged_in(page, ctx, headed, force_manual=False):
    page.goto(BASE_URL if force_manual else CRASH_URL)
    wait_settled(page)

    if force_manual:
        input("\nLog in manually in the browser window, open the Crash page, "
              "then press Enter here... ")
    elif not needs_login(page):
        print("Already logged in to the crash page; skipping login flow and starting search.")
    else:
        print("Logging in ...")
        do_login(page)
        page.goto(CRASH_URL)
        wait_settled(page)
        if needs_login(page):
            save_debug(page, "login_failed")
            if headed:
                input("\nAutomatic login did not complete. Log in manually in the browser "
                      "window, then press Enter here... ")
            else:
                sys.exit("Login failed. Check your credentials, or re-run with "
                         "--show --manual-login. (See debug/login_failed.html)")
    ctx.storage_state(path=STATE_FILE)
    page.goto(CRASH_URL)
    wait_settled(page)


def find_search_box(page):
    if CONFIG["search_input_selector"]:
        return page.locator(CONFIG["search_input_selector"]).first
    box = find_visible(page, [
        'input[type="search"]', 'input[placeholder*="part" i]', 'input[placeholder*="search" i]',
        'input[name*="search" i]', 'input[id*="search" i]', 'input[aria-label*="search" i]',
        'input[type="text"]', 'input:not([type])',
    ])
    if box is None:
        save_debug(page, "no_search_box")
        raise RuntimeError("Could not find the search box on the Crash page "
                           "(see debug/no_search_box.png). Set CONFIG['search_input_selector'].")
    return box


def click_result_candidate(page, part):
    """Pick the most likely result entry after a search and open it."""
    card_selectors = [
        '#search-results-container .part-card',
        '#search-results-container app-product-card',
    ]
    for selector in card_selectors:
        try:
            cards = page.locator(selector)
            visible_cards = []
            for index in range(cards.count()):
                card = cards.nth(index)
                if card.is_visible():
                    visible_cards.append((card, card.inner_text() or ""))
            if not visible_cards:
                continue

            selected = next(
                (card for card, text in visible_cards if part.casefold() in text.casefold()),
                visible_cards[0][0],
            )
            targets = [
                selected.locator(".lkq-link a").first,
                selected.locator(".part-card-image").first,
                selected,
            ]
            for target in targets:
                try:
                    if target.count() and target.is_visible():
                        target.click(timeout=8000)
                        wait_settled(page)
                        return True
                except Exception:
                    continue
        except Exception:
            continue

    selectors = [
        'tr', '[role="row"]', 'a', 'button', 'li', 'div[role="button"]',
        'table tbody tr', '[data-testid*="result" i]', '[class*="result" i]'
    ]
    best = None
    for sel in selectors:
        try:
            items = page.locator(sel)
            for idx in range(min(items.count(), 25)):
                el = items.nth(idx)
                try:
                    if not el.is_visible():
                        continue
                except Exception:
                    continue
                text = (el.inner_text() or "")
                if not text:
                    continue
                if part in text:
                    best = el
                    break
                if any(k in text.lower() for k in ["details", "view", "part", "result"]):
                    if best is None:
                        best = el
        except Exception:
            continue
        if best is not None:
            break

    if best is not None:
        try:
            best.click(timeout=8000)
            wait_settled(page)
            return True
        except Exception:
            pass
    return False


def search_part(page, rec, part, search_label="Partslink Number"):
    rec.clear()
    print(f"Searching {search_label}: {part}")
    page.goto(CRASH_URL)
    wait_settled(page, 500)
    box = find_search_box(page)
    box.click()
    box.fill("")
    box.type(part, delay=20)
    if CONFIG["search_button_selector"]:
        page.locator(CONFIG["search_button_selector"]).first.click()
    else:
        box.press("Enter")
    print(f"Submitted search for {part}; waiting for results...")
    wait_settled(page)
    if CONFIG["result_ready_selector"]:
        try:
            page.wait_for_selector(CONFIG["result_ready_selector"], timeout=10000)
        except PWTimeout:
            pass
    if CONFIG["first_result_click_selector"]:
        try:
            page.locator(CONFIG["first_result_click_selector"]).first.click(timeout=8000)
            wait_settled(page)
            print(f"Selected result for {part}; opening OEM tab...")
            open_oem_tab(page)
            return
        except PWTimeout:
            pass

    clicked = click_result_candidate(page, part)
    if clicked:
        print(f"Selected result for {part}; opening OEM tab...")
        open_oem_tab(page)


# --------------------------------------------------------------------------- #
# Excel in / out
# --------------------------------------------------------------------------- #
def read_parts(path, column):
    df = pd.read_excel(path, dtype=str)
    col = None
    if column:
        if column not in df.columns:
            raise ValueError(
                f"Column '{column}' was not found in '{path}'. "
                f"Available columns: {', '.join(str(name) for name in df.columns)}"
            )
        col = column
    else:
        for c in df.columns:
            if re.search(r"part\s*-?\s*link", str(c), re.I):
                col = c
                break
    if col is None:
        col = df.columns[0]
    parts, seen = [], set()
    for v in df[col].dropna():
        v = str(v).strip()
        if v and v not in seen:
            seen.add(v)
            parts.append(v)
    print(f"Read {len(parts)} unique numbers from column '{col}' of {path}")
    return parts


def read_partslink_mapping(path, search_column):
    """Map each interchange query to its Partslink value from the paired input row."""
    if not is_interchange_search(search_column):
        return {}

    df = pd.read_excel(path, dtype=str).fillna("")
    headers = {str(column).strip().casefold(): column for column in df.columns}
    partslink_column = headers.get("partslink number")
    source_column = headers.get(search_column.strip().casefold())
    if partslink_column is None or source_column is None:
        return {}

    mapping = {}
    ambiguous = set()
    for _, row in df[[source_column, partslink_column]].iterrows():
        query = str(row[source_column]).strip()
        partslink = str(row[partslink_column]).strip()
        if not query or not partslink:
            continue
        if query in mapping and mapping[query] != partslink:
            ambiguous.add(query)
        else:
            mapping[query] = partslink

    for query in ambiguous:
        mapping.pop(query, None)
    return mapping


def apply_partslink_mapping(row, query, mapping):
    """Use a paired source Partslink value when the search query has an unambiguous match."""
    mapped_partslink = mapping.get(query, "")
    if mapped_partslink:
        row["Partslink Number"] = mapped_partslink
    return row


def search_result_column(search_column):
    result_fields = FIELDS + INTERCHANGE_FIELDS + OEM_FIELDS + [NUMBER_VALUES_FIELD, MULTIPLE_NUMBER_FIELD, "Status"]
    return f"Search {search_column}" if search_column in result_fields else search_column


def output_columns(search_column="Partslink Number"):
    columns = [search_result_column(search_column)]
    if search_result_column(search_column).casefold() != "partslink number":
        columns.append("Partslink Number")
    return columns + FIELDS + INTERCHANGE_FIELDS + OEM_FIELDS + [NUMBER_VALUES_FIELD, MULTIPLE_NUMBER_FIELD, "Status"]


OUT_COLS = output_columns()


def load_results(path, search_column="Partslink Number"):
    results = {}
    result_column = search_result_column(search_column)
    columns = output_columns(search_column)
    if Path(path).exists():
        df = pd.read_excel(path, dtype=str).fillna("")
        for col in columns:
            if col not in df.columns:
                df[col] = ""
        for _, r in df.iterrows():
            results.setdefault(r[result_column], []).append({c: r.get(c, "") for c in columns})
    return results


def save_results(path, order, results, search_column="Partslink Number"):
    columns = output_columns(search_column)
    rows = []
    for part in order:
        rows.extend(results.get(part, []))
    df = pd.DataFrame(rows, columns=columns)
    for col in columns:
        if col not in df.columns:
            df[col] = ""
    df = df[columns]
    df.to_excel(path, index=False)


# --------------------------------------------------------------------------- #
# Modes
# --------------------------------------------------------------------------- #
def make_context(p, headed):
    browser = p.chromium.launch(headless=not headed)
    kwargs = {"viewport": {"width": 1400, "height": 900}}
    if Path(STATE_FILE).exists():
        kwargs["storage_state"] = STATE_FILE
    ctx = browser.new_context(**kwargs)
    page = ctx.new_page()
    page.set_default_timeout(20000)
    return browser, ctx, page


def discover(args):
    """Run ONE search and dump everything needed to tune the extraction."""
    out = Path("discover")
    out.mkdir(exist_ok=True)
    with sync_playwright() as p:
        browser, ctx, page = make_context(p, headed=True)
        rec = Recorder(page)
        ensure_logged_in(page, ctx, headed=True, force_manual=args.manual_login)
        search_part(page, rec, args.discover)
        (out / "page.html").write_text(page.content(), encoding="utf-8")
        page.screenshot(path=str(out / "page.png"), full_page=True)
        (out / "page_text.txt").write_text(page.inner_text("body"), encoding="utf-8")
        for i, (url, body) in enumerate(rec.items):
            (out / f"json_{i:02d}.json").write_text(
                json.dumps({"url": url, "body": body}, indent=2, ensure_ascii=False), encoding="utf-8")
        rows = extract(page, rec.items, args.discover)
        print("\n=== Extraction preview ===")
        print(json.dumps(rows, indent=2, ensure_ascii=False) if rows else "(nothing extracted yet)")
        print(f"\nSaved page.html / page.png / page_text.txt and {len(rec.items)} JSON response(s) "
              f"to the '{out}' folder.")
        input("\nPress Enter to close the browser... ")
        browser.close()


def run(args):
    all_parts, search_column = read_parts(args.input, args.column), args.column
    if not search_column:
        input_frame = pd.read_excel(args.input, nrows=0)
        search_column = next(
            (str(column) for column in input_frame.columns if re.search(r"part\s*-?\s*link", str(column), re.I)),
            str(input_frame.columns[0]),
        )
    parts = all_parts[: args.limit] if args.limit else all_parts
    results = {} if args.fresh else load_results(args.output, search_column)
    todo = [x for x in parts
            if not any(r["Status"] == "OK" for r in results.get(x, []))]
    print(f"{len(parts) - len(todo)} already done, {len(todo)} to process.")
    if not todo:
        return

    with sync_playwright() as p:
        browser, ctx, page = make_context(p, headed=args.show)
        rec = Recorder(page)
        ensure_logged_in(page, ctx, headed=args.show, force_manual=args.manual_login)
        try:
            for n, part in enumerate(todo, 1):
                rows, status = [], "ERROR"
                for attempt in range(1, CONFIG["retries"] + 2):
                    try:
                        search_part(page, rec, part, search_column)
                        rows = extract(page, rec.items, part, search_column)
                        status = "OK" if rows else "NO DATA"
                        if not rows:
                            save_debug(page, f"nodata_{re.sub(r'[^A-Za-z0-9_-]', '_', part)}")
                        break
                    except Exception as e:  # noqa: BLE001
                        print(f"   attempt {attempt} failed for {part}: {e}")
                        time.sleep(2)
                if not rows:
                    rows = [{f: "" for f in FIELDS}]
                    rows[0][NUMBER_VALUES_FIELD] = ""
                    rows[0][MULTIPLE_NUMBER_FIELD] = False
                result_column = search_result_column(search_column)
                results[part] = []
                for row in rows:
                    result = {result_column: part, **row, "Status": status}
                    if search_column.strip().casefold() == "partslink number":
                        result["Partslink Number"] = part
                    results[part].append(result)
                for row in rows:
                    inter = row.get("Interchange Number", "") or ""
                    oem = row.get("OEM Number", "") or ""
                    print(f"[{n}/{len(todo)}] {part}: Interchange={inter} | OEM={oem} | Status={status}")
                if not rows:
                    print(f"[{n}/{len(todo)}] {part}: Interchange= | OEM= | Status={status}")
                print(f"[{n}/{len(todo)}] {part}: {status} ({len(rows) if status == 'OK' else 0} row(s))")
                if n % 10 == 0:
                    save_results(args.output, all_parts, results, search_column)
                time.sleep(random.uniform(CONFIG["delay_min"], CONFIG["delay_max"]))
        finally:
            save_results(args.output, all_parts, results, search_column)
            browser.close()
    print(f"\nDone. Results saved to {args.output}")


def main():
    ap = argparse.ArgumentParser(description="Keystone Partslink crawler")
    ap.add_argument("--input", default="parts.xlsx", help="Excel file with Partslink numbers")
    ap.add_argument("--column", default="", help="Input column to search; it also identifies results")
    ap.add_argument("--output", default="keystone_results.xlsx")
    ap.add_argument("--show", action="store_true", help="Show the browser window")
    ap.add_argument("--limit", type=int, default=0, help="Only process the first N numbers")
    ap.add_argument("--fresh", action="store_true", help="Ignore an existing results file")
    ap.add_argument("--manual-login", action="store_true", help="Log in by hand in the browser")
    ap.add_argument("--discover", metavar="PARTSLINK", help="Run one search and dump debug files")
    ap.add_argument("--make-template", action="store_true", help="Create an empty parts.xlsx")
    args = ap.parse_args()

    if args.make_template:
        if pd is None:
            raise RuntimeError(
                "Missing Python dependencies. Run:\n"
                "  python -m pip install -r requirements.txt"
            )
        pd.DataFrame({"Partslink Number": ["PASTE_FIRST_NUMBER", "PASTE_SECOND_NUMBER"]}) \
            .to_excel("parts.xlsx", index=False)
        print("Created parts.xlsx - replace the example rows with your numbers.")
        return

    try:
        ensure_runtime_dependencies()
        if not args.discover:
            ensure_input_file(args.input)
        ensure_browser_available()
    except (RuntimeError, FileNotFoundError) as exc:
        sys.exit(str(exc))

    if args.discover:
        discover(args)
    else:
        run(args)


if __name__ == "__main__":
    main()
