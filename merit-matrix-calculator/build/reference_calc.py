# -*- coding: utf-8 -*-
"""
Independent reference for the merit workbook.

The workbook and the page are checked against these numbers, not by eye.
Target-based merit: each rating has a target compa-ratio, the rise closes the
gap to it within the caps. Amounts are in roubles; employer contributions follow
the Russian 2026 rates.
"""
import csv
import os
import sys
from datetime import date

# the demo files and the workbook sit in the artefact folder, one level up
ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..")
SAMPLE_EMPLOYEES = os.path.join(ROOT, "sample-employees.csv")
SAMPLE_RANGES = os.path.join(ROOT, "sample-salary-ranges.csv")
MODEL = os.path.join(ROOT, "merit-model.xlsx")

DEFAULTS = {
    "effective_date": date(2027, 4, 1),
    "pool_pct": 4.0,
    "min_months_hire": 6,
    "min_months_change": 3,
    "clip_at_max": True,
    "company_target": 1.00,
    # one cap for everyone, % of base; the rise itself is the gap to target.
    # None or 0 = no cap
    "max_increase": 0.30,
    # rating: (target compa-ratio, minimum increase)
    "ratings": {1: (0.00, 0.00), 2: (0.95, 0.00), 3: (1.00, 0.00),
                4: (1.05, 0.01), 5: (1.10, 0.02)},
    "zones": [0.0, 0.90, 1.00, 1.10],
    # employer insurance contributions, Russia 2026 (Government Resolution No. 1705 of
    # 31.10.2025): 30% up to the annual base ceiling, 15.1% above it; injury insurance not included
    "rate_below": 30.0,
    "rate_above": 15.1,
    "ceiling": 2979000.0,
}
EPS = 1e-9


def parse_date(s):
    y, m, d = (int(x) for x in s.split("-"))
    return date(y, m, d)


def load_bands(path):
    """Salary ranges keyed by (category, grade): annual full-time base, min / mid / max."""
    return {(r["category"].strip(), int(r["grade"])): (float(r["min"]), float(r["mid"]), float(r["max"]))
            for r in csv.DictReader(open(path, encoding="utf-8-sig"))}


def months_between(d1, d2):
    """Complete months from d1 to d2, as Excel DATEDIF(d1, d2, "m")."""
    return (d2.year - d1.year) * 12 + (d2.month - d1.month) - (1 if d2.day < d1.day else 0)


def load(path):
    out = []
    for r in csv.DictReader(open(path, encoding="utf-8-sig")):
        out.append({
            "id": r["id"], "category": r["category"].strip(),
            "grade": int(r["grade"]), "rating": int(r["rating"]),
            "hire": parse_date(r["hire_date"]), "change": parse_date(r["last_change_date"]),
            "base": float(r["base_salary"]), "bonus": float(r["bonus_pct"]) / 100,
            "fte": float(r["fte"]),
        })
    return out


def contributions(pay, S):
    """Employer contributions on annual pay: one rate up to the ceiling, another above it."""
    c = S["ceiling"]
    return S["rate_below"] / 100 * min(pay, c) + S["rate_above"] / 100 * max(0.0, pay - c)


def contributions_on(before, uplift, S):
    if uplift <= 0:
        return 0.0
    return contributions(before + uplift, S) - contributions(before, S)


def analyse(emp, bands, S=None):
    S = dict(DEFAULTS, **(S or {}))
    eff = S["effective_date"]
    for e in emp:
        e["min"], e["mid"], e["max"] = bands[(e["category"], e["grade"])]
        e["tenure"] = (eff - e["hire"]).days / 365.25
        e["total"] = e["base"] * (1 + e["bonus"])

    # Target-based merit
    for e in emp:
        e["eligible"] = (months_between(e["hire"], eff) >= S["min_months_hire"]
                         and months_between(e["change"], eff) >= S["min_months_change"])
        e["cr"] = e["base"] / e["mid"]
        tgt, mn = S["ratings"][e["rating"]]
        cap = S["max_increase"]
        target_salary = e["mid"] * S["company_target"] * tgt
        gap = max(0.0, target_salary - e["base"])
        raw_inc = 0.0 if tgt == 0 else max(gap, mn * e["base"])
        if cap:
            raw_inc = min(raw_inc, cap * e["base"])
        if e["base"] >= e["max"]:
            raw_inc = 0.0
        if S["clip_at_max"]:
            raw_inc = min(raw_inc, max(0.0, e["max"] - e["base"]))
        e["increase"] = raw_inc if e["eligible"] else 0.0
        e["new_base"] = e["base"] + e["increase"]
        # target: the same pool % to every eligible person, as a full-cost benchmark
        e["target_inc"] = S["pool_pct"] / 100 * e["base"] if e["eligible"] else 0.0
        e["zone"] = max(i for i, z in enumerate(S["zones"]) if e["cr"] >= z) + 1

    # Costs
    factor = (13 - eff.month) / 12
    res = {"merit": {}, "target": {}}
    for tag, inc_key in (("merit", "increase"), ("target", "target_inc")):
        base_inc = sum(e[inc_key] * e["fte"] for e in emp)
        bonus_inc = sum(e[inc_key] * e["bonus"] * e["fte"] for e in emp)
        contrib = sum(contributions_on(e["base"] * (1 + e["bonus"]) * e["fte"],
                                       e[inc_key] * (1 + e["bonus"]) * e["fte"], S) for e in emp)
        annual = base_inc + bonus_inc + contrib
        res[tag] = {"base": base_inc, "bonus": bonus_inc, "contrib": contrib,
                    "annual": annual, "in_year": annual * factor}

    pay_before = sum(e["total"] * e["fte"] for e in emp)
    contrib_before = sum(contributions(e["total"] * e["fte"], S) for e in emp)
    payroll = pay_before + contrib_before
    eligible_base = sum(e["base"] * e["fte"] for e in emp if e["eligible"])
    merit_pct = res["merit"]["base"] / eligible_base * 100
    cost_annual = res["merit"]["annual"]
    res.update({
        "pay_before": pay_before, "contrib_before": contrib_before,
        "payroll": payroll, "eligible_base": eligible_base,
        "merit_pct": merit_pct, "dev_pp": merit_pct - S["pool_pct"],
        "dev_rub": res["merit"]["annual"] - res["target"]["annual"],
        "dev_rub_in_year": (res["merit"]["annual"] - res["target"]["annual"]) * factor,
        "cost_annual": cost_annual, "cost_in_year": cost_annual * factor,
        "load_full_pct": cost_annual / payroll * 100,
        "load_in_year_pct": cost_annual * factor / payroll * 100,
        "below_min_after": sum(1 for e in emp if e["new_base"] < e["min"]),
        "red_circled": sum(1 for e in emp if e["base"] >= e["max"]),
        "emp": emp,
    })
    return res


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else SAMPLE_EMPLOYEES
    bnd = sys.argv[2] if len(sys.argv) > 2 else SAMPLE_RANGES
    r = analyse(load(src), load_bands(bnd))
    for e in r["emp"]:
        print("%s CR=%.4f elig=%d inc=%12.2f zone=%d" % (
            e["id"], e["cr"], e["eligible"], e["increase"], e["zone"]))
    for t in ("merit", "target"):
        print(t, {k: round(v, 2) for k, v in r[t].items()})
    for k in ("pay_before", "contrib_before", "payroll", "eligible_base", "merit_pct", "dev_pp", "dev_rub", "dev_rub_in_year",
              "cost_annual", "cost_in_year", "load_full_pct", "load_in_year_pct", "below_min_after", "red_circled"):
        print("%-18s %.4f" % (k, r[k]))
