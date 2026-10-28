"""Behaviour of the built page (../index.html) in a real browser, by clicks and real file inputs.

    python check-page.py

Listens to every request from page load to the last file choice: the page must make none.
Checks that parameters recalculate the screen, that the demo, a small subset of it, a file without
a category column and broken files behave as intended, and that the numbers on screen
match the engine run in Node on the same input. Exit 1 on any failure.
"""
import json
import pathlib
import re
import subprocess
import sys
import tempfile

from playwright.sync_api import sync_playwright

HERE = pathlib.Path(__file__).resolve().parent
TMP = pathlib.Path(tempfile.mkdtemp())  # the subset, broken and edited copies of it
PAGE = (HERE.parent / "index.html").as_uri()
DEMO = HERE.parent / "sample-employees.csv"
RANGES = HERE.parent / "sample-salary-ranges.csv"
fails = []

# a small file of the reader's size: every eighth employee of the demo, under its own name
SUBSET = TMP / "subset-employees.csv"
_rows = DEMO.read_text(encoding="utf-8").splitlines()
SUBSET.write_text("\n".join([_rows[0]] + _rows[1::8]) + "\n", encoding="utf-8")
# and only the ranges it uses: the page refuses an employee file on the demo's own ranges
_col = _rows[0].split(",")
_used = {(r.split(",")[_col.index("category")], r.split(",")[_col.index("grade")]) for r in _rows[1::8]}
_ranges = RANGES.read_text(encoding="utf-8").splitlines()
SUBSET_RANGES = TMP / "subset-salary-ranges.csv"
SUBSET_RANGES.write_text("\n".join([_ranges[0]] + [r for r in _ranges[1:] if tuple(r.split(",")[:2]) in _used]) + "\n",
                         encoding="utf-8")


def check(name, ok, detail=""):
    print(("OK   " if ok else "FAIL ") + name + (f" — {detail}" if detail and not ok else ""))
    if not ok:
        fails.append(name)


def engine(employees, bands, settings):
    """The engine in Node on the same files: what the screen must show."""
    js = ("import {run} from './src/engine.js';import {readFileSync as r} from 'node:fs';"
          f"const d=run(r({json.dumps(str(employees))},'utf8'),r({json.dumps(str(bands))},'utf8'),{json.dumps(settings)});"
          "console.log(JSON.stringify({total:d.budget.change_total_in_year,annual:d.budget.change_total_annual,"
          "pct:d.pool.merit_pct,n:d.headcount,below:[d.cr_all.below_low_before,d.cr_all.below_low_after]}))")
    out = subprocess.run(["node", "--input-type=module", "-e", js], cwd=HERE, capture_output=True, text=True, check=True)
    return json.loads(out.stdout)


def m(x):
    return f"{x / 1e6:.2f} млн ₽".replace(".", ",")


def kpis(p):
    return p.locator(".kpi-value").all_inner_texts()


with sync_playwright() as pw:
    b = pw.chromium.launch()
    p = b.new_page(viewport={"width": 1280, "height": 800})
    requests, errors = [], []
    p.on("request", lambda r: requests.append(r.url) if not r.url.startswith(("file:", "data:", "blob:")) else None)
    p.on("pageerror", lambda e: errors.append(str(e)))
    p.goto(PAGE)
    p.wait_for_selector(".kpi-value")

    demo = engine(DEMO, RANGES, {})
    v = kpis(p)
    check("demo: 2027 cost = engine", v[0] == m(demo["total"]), f"{v[0]} vs {m(demo['total'])}")
    check("demo: merit % = engine", v[2] == f"{demo['pct']:.2f}%".replace(".", ","), f"{v[2]}")
    check("demo: headcount label", f"сотрудников: {demo['n']}" in p.inner_text(".source-label"))

    # parameters recalculate; January: both cards equal
    p.get_by_label("Дата пересмотра", exact=True).select_option("янв 2027")
    v = kpis(p)
    check("January: 2027 card = 12-month card", v[0] == v[1], f"{v[0]} vs {v[1]}")
    check("January: labels follow the date", "12 месяцев 2028" in p.inner_text(".kpi-groups"))
    p.get_by_label("Дата пересмотра", exact=True).select_option("дек 2028")
    check("Dec 2028: years follow the date", "2028 · с даты пересмотра" in p.inner_text(".kpi-groups") and "12 месяцев 2029" in p.inner_text(".kpi-groups"))
    p.get_by_label("Дата пересмотра", exact=True).select_option("апр 2027")
    p.get_by_label("Целевой мерит", exact=True).fill("3,5")
    j = engine(DEMO, RANGES, {"pool_pct": 3.5})
    j35 = engine(DEMO, RANGES, {"pool_pct": 3.5})
    check("target merit increase with a comma recalculates", f"{j35['pct'] - 3.5:+.2f}".replace(".", ",").replace("-", "−") + " п.п. к цели" in p.inner_text(".kpi-groups"), p.inner_text(".kpi-groups")[:300])
    check("target merit increase: cost unchanged by target", kpis(p)[0] == m(j["total"]))
    p.get_by_label("Целевой мерит", exact=True).fill("4.0")
    p.get_by_label("Потолок повышения", exact=True).select_option("Без ограничения")
    j = engine(DEMO, RANGES, {"max_increase": None})
    check("No cap recalculates = engine", kpis(p)[0] == m(j["total"]), f"{kpis(p)[0]} vs {m(j['total'])}")
    p.get_by_label("Потолок повышения", exact=True).select_option("30")
    p.get_by_label("Выдающийся: целевой CR", exact=True).fill("")
    j = engine(DEMO, RANGES,
               {"ratings": {"1": [0, 0], "2": [0.95, 0], "3": [1, 0], "4": [1.05, 0.01], "5": [0, 0.02]}})
    check("empty rating target = no increase", kpis(p)[0] == m(j["total"]), f"{kpis(p)[0]} vs {m(j['total'])}")
    p.get_by_label("Выдающийся: целевой CR", exact=True).fill("110")
    check("back to defaults", kpis(p)[0] == m(demo["total"]))

    # data section: real file input
    p.get_by_role("button", name="Загрузить свои данные").click()
    p.wait_for_selector("#data-content")
    inp = p.locator("input[type=file]")
    inp.set_input_files([str(SUBSET), str(SUBSET_RANGES)])
    subset = engine(SUBSET, SUBSET_RANGES, {})
    p.wait_for_function(f"document.querySelector('.source-label').innerText.includes('сотрудников: {subset['n']}')", timeout=5000)
    check("subset loaded: 2027 cost = engine", kpis(p)[0] == m(subset["total"]), f"{kpis(p)[0]} vs {m(subset['total'])}")
    check("subset loaded: file names in the header", "subset-employees.csv" in p.inner_text(".source-label"))
    below = f"{subset['below'][0]} → {subset['below'][1]}\nниже 90% середины вилки"
    check("subset loaded: below-90% count = engine", below in p.inner_text(".people-strip") and subset["below"] != demo["below"], below)

    # broken file: red line, data kept
    tmp = TMP / "broken.csv"
    tmp.write_text("id,grade\nX,1\n", encoding="utf-8")
    inp.set_input_files(str(tmp))
    p.wait_for_selector(".import-result.error")
    check("broken file: error shown, numbers kept", kpis(p)[0] == m(subset["total"]))

    # rating outside 1-5: red line, data kept
    badr = TMP / "bad-rating.csv"
    rows = SUBSET.read_text(encoding="utf-8").splitlines()
    rcol = rows[0].split(",").index("rating")
    first = rows[1].split(","); first[rcol] = "7"; rows[1] = ",".join(first)
    badr.write_text("\n".join(rows) + "\n", encoding="utf-8")
    inp.set_input_files([str(badr), str(SUBSET_RANGES)])
    p.wait_for_function("document.querySelector('.import-result.error')?.innerText.includes('от 1 до 5')", timeout=5000)
    check("rating outside 1-5: error, numbers kept", kpis(p)[0] == m(subset["total"]))

    # no category column: ranges matched by grade only, one total row
    txt = SUBSET.read_text(encoding="utf-8").splitlines()
    ci = txt[0].split(",").index("category")
    noc = TMP / "no-category.csv"
    noc.write_text("\n".join(",".join(c for i, c in enumerate(r.split(",")) if i != ci) for r in txt) + "\n", encoding="utf-8")
    by_grade = {}
    for r in SUBSET_RANGES.read_text(encoding="utf-8").splitlines()[1:]:
        by_grade.setdefault(r.split(",")[1], ",".join(r.split(",")[1:]))
    noc_r = TMP / "no-category-ranges.csv"
    noc_r.write_text("\n".join(["grade,min,mid,max"] + list(by_grade.values())) + "\n", encoding="utf-8")
    inp.set_input_files([str(noc), str(noc_r)])
    j = engine(noc, noc_r, {})
    p.wait_for_function("document.querySelector('.source-label').innerText.includes('no-category.csv')", timeout=5000)
    check("no category: cost = engine", kpis(p)[0] == m(j["total"]), f"{kpis(p)[0]} vs {m(j['total'])}")
    check("no category: one total row", "Все сотрудники" in p.inner_text(".category-table"))

    p.get_by_role("button", name="Вернуть демо-данные").click()
    inp.set_input_files(str(SUBSET))
    p.wait_for_function("document.querySelector('.import-result.error')?.innerText.includes('Загрузите вилки вместе')", timeout=5000)
    check("employee file alone on demo ranges: refused", kpis(p)[0] == m(demo["total"]))
    check("demo restored", kpis(p)[0] == m(demo["total"]))

    # own employees with the unchanged sample ranges, both chosen at once: accepted
    inp.set_input_files([str(SUBSET), str(RANGES)])
    try:
        p.wait_for_function(f"document.querySelector('.source-label').innerText.includes('сотрудников: {subset['n']}')", timeout=5000)
        ok = not p.query_selector(".import-result.error")
    except Exception:
        ok = False
    check("employee file with the unchanged sample ranges: accepted", ok)
    p.get_by_role("button", name="Вернуть демо-данные").click()

    # 390 px after loading files: long file names must not widen the page
    ph = b.new_page(viewport={"width": 390, "height": 844})
    ph.on("request", lambda r: requests.append(r.url) if not r.url.startswith(("file:", "data:", "blob:")) else None)
    ph.goto(PAGE)
    ph.get_by_role("button", name="Загрузить свои данные").click()
    ph.wait_for_selector("#data-content")
    ph.locator("input[type=file]").set_input_files([str(SUBSET), str(SUBSET_RANGES)])
    ph.wait_for_function(f"document.querySelector('.source-label').innerText.includes('сотрудников: {subset['n']}')", timeout=5000)
    w = ph.evaluate("document.documentElement.scrollWidth")
    check("390 after file load: no horizontal scroll", w <= 390, f"scrollWidth {w}")
    ph.close()

    check("zero network requests, file choices included", not requests, "; ".join(requests[:5]))
    check("no JS errors", not errors, "; ".join(errors[:3]))
    b.close()

print(f"FAILURES: {len(fails)}")
sys.exit(1 if fails else 0)
