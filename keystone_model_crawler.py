#!/usr/bin/env python3
"""Search Keystone Crash by vehicle/model and crawl every matching part."""

import argparse
import re
import sys
import time
from pathlib import Path

try:
    import pandas as pd
except ImportError:  # handled in main()
    pd = None

try:
    from playwright.sync_api import TimeoutError as PWTimeout
except ImportError:  # handled in main()
    PWTimeout = Exception

import keystone_crawler as base


INPUT_COLUMNS = ["Year", "Brand", "Model", "Parts", "Category"]
INTERCHANGE_FIELDS = [f"Interchange Number {number}" for number in range(1, 6)]
OEM_FIELDS = [f"OEM Number {number}" for number in range(1, 6)]
OUTPUT_COLUMNS = [
    "Partslink Number",
    "Oldest Year",
    "Newest Year",
    "Brand",
    "Model",
    "Interchange Number",
    "OEM Number",
    *INTERCHANGE_FIELDS,
    *OEM_FIELDS,
    "Number Values",
    "Description",
]
PARTSLINK_PATTERN = re.compile(r"\b[A-Z]{1,4}[- ]?\d{5,10}[A-Z]?\b", re.I)


def clean(value):
    return re.sub(r"\s+", " ", str(value or "")).strip()


def split_categories(value):
    """Split category names entered as a comma/semicolon/pipe-separated cell."""
    categories = []
    seen = set()
    for category in re.split(r"[,;|\n]+", clean(value)):
        category = category.strip()
        key = category.casefold()
        if category and key not in seen:
            categories.append(category)
            seen.add(key)
    return categories


def parse_year_bounds(text):
    years = [int(year) for year in re.findall(r"\b(?:19|20)\d{2}\b", text or "")]
    return (str(min(years)), str(max(years))) if years else ("", "")


def read_search_rows(path):
    if pd is None:
        raise RuntimeError("Missing pandas/openpyxl. Install requirements.txt first.")
    file_path = Path(path)
    if not file_path.exists():
        raise FileNotFoundError(
            f"Input file '{path}' was not found. Run: "
            "python keystone_model_crawler.py --make-template"
        )
    frame = pd.read_excel(file_path, dtype=str).fillna("")
    missing = [column for column in INPUT_COLUMNS if column not in frame.columns]
    if missing:
        raise ValueError(f"Input file is missing required column(s): {', '.join(missing)}")
    return [
        {column: clean(record[column]) for column in INPUT_COLUMNS}
        for record in frame.to_dict(orient="records")
        if any(clean(record[column]) for column in INPUT_COLUMNS)
    ]


def make_template(path):
    if pd is None:
        raise RuntimeError("Missing pandas/openpyxl. Install requirements.txt first.")
    pd.DataFrame(columns=INPUT_COLUMNS).to_excel(path, index=False)
    print(f"Created {path} with columns: {', '.join(INPUT_COLUMNS)}")


def _visible(locator):
    try:
        return locator.count() > 0 and locator.first.is_visible()
    except Exception:
        return False


def _field_control(page, label):
    """Find a form control from its visible field label, including Angular Material."""
    placeholders = {
        "year": "select year",
        "brand": "select make",
        "model": "select model",
        "model parts": "select category",
    }
    expected_placeholder = placeholders.get(label.casefold())
    if expected_placeholder:
        for index in range(page.locator(".ng-placeholder").count()):
            placeholder = page.locator(".ng-placeholder").nth(index)
            if clean(placeholder.inner_text()).casefold() != expected_placeholder:
                continue
            dropdown = placeholder.locator(
                "xpath=ancestor::*[contains(concat(' ', normalize-space(@class), ' '), ' ng-select ')][1]"
            )
            control = dropdown.locator("[role='combobox']").first
            if _visible(control):
                return control

    exact_text = page.get_by_text(label, exact=True)
    for index in range(min(exact_text.count(), 8)):
        text = exact_text.nth(index)
        for ancestor in (
            "xpath=ancestor::mat-form-field[1]",
            "xpath=ancestor::label[1]",
            "xpath=ancestor::*[self::div or self::section][1]",
        ):
            container = text.locator(ancestor)
            control = container.locator(
                "mat-select, select, input, [role='combobox'], button"
            ).first
            if _visible(control):
                return control
    try:
        control = page.get_by_label(label, exact=True).first
        if _visible(control):
            return control
    except Exception:
        pass
    raise RuntimeError(f"Could not find the '{label}' search field")


def _choose_value(page, label, value):
    control = _field_control(page, label)
    tag = control.evaluate("element => element.tagName.toLowerCase()")
    if tag == "select":
        try:
            control.select_option(label=value)
        except Exception:
            control.select_option(value=value)
    elif tag == "input":
        control.fill(value)
        options = page.locator("mat-option, [role='option']")
        for index in range(min(options.count(), 50)):
            option = options.nth(index)
            if _visible(option) and clean(option.inner_text()).casefold() == value.casefold():
                option.click()
                break
        else:
            control.press("Enter")
    else:
        control.click()
        options = page.locator("mat-option, .ng-option, [role='option']")
        for index in range(min(options.count(), 100)):
            option = options.nth(index)
            try:
                if _visible(option) and clean(option.inner_text()).casefold() == value.casefold():
                    option.click()
                    break
            except Exception:
                continue
        else:
            raise RuntimeError(f"Option '{value}' not found for '{label}'")
    base.wait_settled(page, 300)


def _set_category_checked(page, category):
    exact_text = page.get_by_text(re.compile(rf"^\s*{re.escape(category)}\s*$", re.I))
    for index in range(min(exact_text.count(), 10)):
        text = exact_text.nth(index)
        for ancestor in (
            "xpath=ancestor::mat-checkbox[1]",
            "xpath=ancestor-or-self::label[1]",
            "xpath=ancestor::*[@role='checkbox'][1]",
        ):
            container = text.locator(ancestor)
            if not _visible(container):
                continue
            checkbox = container.locator("input[type='checkbox']")
            if checkbox.count():
                if not checkbox.is_checked():
                    checkbox.check(force=True)
                return
            state = container.get_attribute("aria-checked")
            classes = container.get_attribute("class") or ""
            if state == "true" or "mat-checkbox-checked" in classes:
                return
            container.click()
            return
    raise RuntimeError(f"Could not find category checkbox '{category}'")


def search_model(page, record):
    """Fill the Crash vehicle form in dependency order and submit it."""
    page.goto(base.CRASH_URL)
    base.wait_settled(page)
    _choose_value(page, "Year", record["Year"])
    _choose_value(page, "Brand", record["Brand"])
    _choose_value(page, "Model", record["Model"])
    _choose_value(page, "Model Parts", record["Parts"])
    for category in split_categories(record["Category"]):
        _set_category_checked(page, category)
    search_buttons = page.get_by_role("button", name=re.compile(r"^\s*search\s*$", re.I))
    search_buttons.last.click()
    base.wait_settled(page)


def _result_candidates(page):
    selectors = [
        "section",
        "table tbody tr",
        "[role='row']",
        "mat-row",
        "mat-card",
        "[class*='result' i]",
    ]
    candidates = []
    seen_text = set()
    for selector in selectors:
        items = page.locator(selector)
        for index in range(min(items.count(), 500)):
            item = items.nth(index)
            try:
                if not item.is_visible():
                    continue
                text = clean(item.inner_text())
                if not text or not PARTSLINK_PATTERN.search(text):
                    continue
                if text not in seen_text:
                    seen_text.add(text)
                    candidates.append((item, text))
            except Exception:
                continue
        if candidates:
            break
    return candidates


def _partslink_from_text(text):
    match = PARTSLINK_PATTERN.search(text or "")
    return re.sub(r"[- ]", "", match.group(0)).upper() if match else ""


def _active_dialog(page):
    for selector in ("mat-dialog-container", "[role='dialog']", ".modal-dialog"):
        dialogs = page.locator(selector)
        for index in range(dialogs.count() - 1, -1, -1):
            dialog = dialogs.nth(index)
            if _visible(dialog):
                return dialog
    return None


def _click_tab(dialog, name):
    tabs = dialog.get_by_text(name, exact=True)
    for index in range(min(tabs.count(), 10)):
        tab = tabs.nth(index)
        if _visible(tab):
            try:
                tab.click(timeout=3000)
                return True
            except Exception:
                pass
    return False


def _extract_description(dialog):
    for row in dialog.locator("tr, [role='row']").all():
        try:
            cells = row.locator("th, td, [role='cell']").all_inner_texts()
            if len(cells) >= 2 and re.search(r"description", cells[0], re.I):
                return clean(" ".join(cells[1:]))
        except Exception:
            continue
    text = dialog.inner_text()
    match = re.search(r"(?im)^\s*Description\s*:?\s*(.+?)\s*$", text)
    if match:
        return clean(match.group(1))
    for selector in (".description", ".part-description", ".product-description"):
        nodes = dialog.locator(selector)
        for index in range(nodes.count()):
            node = nodes.nth(index)
            if _visible(node):
                value = clean(node.inner_text())
                if value:
                    return value
    return ""


def _read_fitments(page, dialog):
    _click_tab(dialog, "Fitments")
    base.wait_settled(page, 400)
    dialog = _active_dialog(page) or dialog
    text = dialog.inner_text()
    oldest, newest = parse_year_bounds(text)
    if not oldest:
        pairs = []
        for row in dialog.locator("tr, [role='row']").all():
            try:
                cells = row.locator("th, td, [role='cell']").all_inner_texts()
                if len(cells) >= 2:
                    pairs.append((clean(cells[0]), clean(" ".join(cells[1:]))))
            except Exception:
                continue
        for key, value in pairs:
            if re.search(r"oldest|earliest|start year", key, re.I):
                match = re.search(r"\b(?:19|20)\d{2}\b", value)
                oldest = parse_year_bounds(value)[0] or (match.group(0) if match else "")
            elif re.search(r"newest|latest|end year", key, re.I):
                match = re.search(r"\b(?:19|20)\d{2}\b", value)
                newest = parse_year_bounds(value)[1] or (match.group(0) if match else "")
    return oldest, newest


def _close_dialog(page, dialog):
    for name in ("Close", "Done", "Cancel"):
        button = dialog.get_by_role("button", name=re.compile(name, re.I))
        if _visible(button):
            button.first.click()
            base.wait_settled(page, 250)
            return
    page.keyboard.press("Escape")
    base.wait_settled(page, 250)


def extract_result(page, record, result_text):
    dialog = _active_dialog(page)
    if dialog is None:
        raise RuntimeError("Clicked result did not open a detail modal")
    part = _partslink_from_text(result_text)
    description = _extract_description(dialog)
    oldest, newest = _read_fitments(page, dialog)
    dialog = _active_dialog(page) or dialog
    base.open_oem_tab(page)
    numbers = base.extract_oem_tab_values(page, part)
    current_part_aliases = {part.upper(), re.sub(r"[A-Z]$", "", part.upper())}
    interchange_tokens = [
        token for token in base.normalize_number_tokens(numbers.get("interchange", ""))
        if token.upper() not in current_part_aliases
    ]
    oem_tokens = [
        token for token in base.normalize_number_tokens(numbers.get("oem", ""))
        if token.upper() not in current_part_aliases
    ]
    numbers["interchange"] = ", ".join(interchange_tokens)
    numbers["oem"] = ", ".join(oem_tokens)
    numbers["number_values"] = ", ".join(filter(None, [numbers["interchange"], numbers["oem"]]))
    row = {column: "" for column in OUTPUT_COLUMNS}
    row.update({
        "Partslink Number": part,
        "Oldest Year": oldest,
        "Newest Year": newest,
        "Brand": record["Brand"],
        "Model": record["Model"],
        "Interchange Number": numbers.get("interchange", ""),
        "OEM Number": numbers.get("oem", ""),
        "Number Values": numbers.get("number_values", ""),
        "Description": description,
    })
    for index, value in enumerate(base.split_number_values(numbers.get("interchange", "")), 1):
        row[f"Interchange Number {index}"] = value
    for index, value in enumerate(base.split_number_values(numbers.get("oem", "")), 1):
        row[f"OEM Number {index}"] = value
    _close_dialog(page, dialog)
    return row


def crawl_results(page, record):
    """Open each distinct result while scrolling the result viewport to its end."""
    rows = []
    seen = set()
    quiet_rounds = 0
    previous_height = -1
    while quiet_rounds < 3:
        added = False
        for item, text in _result_candidates(page):
            if text in seen:
                continue
            seen.add(text)
            part = _partslink_from_text(text)
            number_link = item.locator("a").filter(
                has_text=re.compile(rf"^\s*{re.escape(part)}\s*$", re.I)
            )
            if number_link.count():
                number_link.first.click()
            else:
                item.click()
            base.wait_settled(page, 500)
            rows.append(extract_result(page, record, text))
            added = True
        height = page.evaluate(
            "Math.max(document.body.scrollHeight, document.documentElement.scrollHeight)"
        )
        page.evaluate(
            "window.scrollTo(0, document.body.scrollHeight);"
            "document.querySelectorAll('*').forEach(e => {"
            "if (e.scrollHeight > e.clientHeight + 80) e.scrollTop = e.scrollHeight;"
            "});"
        )
        base.wait_settled(page, 700)
        new_height = page.evaluate(
            "Math.max(document.body.scrollHeight, document.documentElement.scrollHeight)"
        )
        if not added and new_height == previous_height == height:
            quiet_rounds += 1
        else:
            quiet_rounds = 0
        previous_height = new_height
    return rows


def save_results(path, rows):
    pd.DataFrame(rows, columns=OUTPUT_COLUMNS).to_excel(path, index=False)


def run(args):
    rows = read_search_rows(args.input)
    if args.limit:
        rows = rows[:args.limit]
    if not rows:
        print("No input rows to process.")
        return
    base.ensure_runtime_dependencies()
    base.ensure_browser_available()
    output_rows = []
    if not args.fresh and Path(args.output).exists():
        existing = pd.read_excel(args.output, dtype=str).fillna("")
        output_rows = existing.to_dict(orient="records")

    with base.sync_playwright() as playwright:
        browser, context, page = base.make_context(playwright, headed=args.show)
        base.ensure_logged_in(page, context, headed=args.show, force_manual=args.manual_login)
        try:
            for index, record in enumerate(rows, 1):
                print(f"[{index}/{len(rows)}] Search: {record['Year']} {record['Brand']} "
                      f"{record['Model']} / {record['Parts']} / {record['Category']}")
                search_model(page, record)
                results = crawl_results(page, record)
                output_rows.extend(results)
                save_results(args.output, output_rows)
                print(f"[{index}/{len(rows)}] Saved {len(results)} result(s) to {args.output}")
                time.sleep(1)
        finally:
            save_results(args.output, output_rows)
            browser.close()


def main():
    parser = argparse.ArgumentParser(description="Keystone Crash model/parts crawler")
    parser.add_argument("--input", default="find_model.xlsx", help="Input workbook")
    parser.add_argument("--output", default="keystone_model_results.xlsx", help="Output workbook")
    parser.add_argument("--limit", type=int, default=0, help="Only process the first N input rows")
    parser.add_argument("--show", action="store_true", help="Show the browser window")
    parser.add_argument("--fresh", action="store_true", help="Ignore an existing output workbook")
    parser.add_argument("--manual-login", action="store_true", help="Log in manually in the browser")
    parser.add_argument("--make-template", action="store_true", help="Create find_model.xlsx")
    args = parser.parse_args()
    try:
        if args.make_template:
            make_template(args.input)
            return
        run(args)
    except (RuntimeError, FileNotFoundError, ValueError) as exc:
        sys.exit(str(exc))


if __name__ == "__main__":
    main()