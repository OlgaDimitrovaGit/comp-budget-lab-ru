#!/usr/bin/env python3
"""
The plan-vs-actual payroll budget dataset behind the dashboard.

Built on top of the pay-gap calculator's demo dataset (212 employees,
7 departments), so both tools describe the same modelled company.

Three figures per department and month:
  plan     — payroll budget: base plus variable pay, WITHOUT employer
             contributions (in Russian practice "ФОТ" excludes them)
  fact     — actuals Jan-Aug; empty from September
  forecast — Sep-Dec projection: plan carried on the department's own
             run-rate deviation

Everything is deterministic: the seed is a constant and randomness flows only
through it. Running this twice must produce a byte-for-byte identical file.
"""

import csv
import os
import random
import sys

HERE = os.path.dirname(os.path.abspath(__file__))


def _find_paygap():
    """Locate the pay-gap calculator's demo data, which this dataset is built on.

    In the repository each tool has a folder of its own, so the data sits at
    <repo>/pay-gap-calculator/build/excel.
    """
    path = os.path.normpath(os.path.join(
        HERE, "..", "..", "pay-gap-calculator", "build", "excel"))
    if os.path.exists(os.path.join(path, "demo-data.csv")):
        return path
    raise SystemExit("demo-data.csv not found in " + path)


PAYGAP = _find_paygap()

# The seed is a date rather than an arbitrary number, so that it means something.
SEED = 20261028

# Russian short month names: the CSV is what the reader opens and fills in.
MONTHS = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг",
          "сен", "окт", "ноя", "дек"]
FACT_MONTHS = 8          # actuals are known through August inclusive

# --- Target properties of the actuals ---------------------------------------
MONTHLY_DEV_MIN = 1.0      # |deviation| per department per month, % — lower bound
MONTHLY_DEV_MAX = 3.0      # the same, upper bound

# The company-level result is NOT set as a number; it accumulates from the
# departments' deviations.
#
# The dashboard has to show a problem: a signalling tool that never signals is
# useless. The "problem" threshold is 3% for the year. With no pay decision
# landing in the fourth quarter, the forecast carries each department's own
# run rate, so the year lands close to the eight-month figure. The eight months
# therefore have to clear the threshold on their own, with a margin: a year at
# exactly 3.00% drops back to yellow on any edit to the numbers.
EXPECTED_OVERSPEND_MIN = 3.2  # % overspend over 8M, lower bound
EXPECTED_OVERSPEND_MAX = 4.5  # % upper bound

# The dashboard's "problem" threshold: 3% and above. It is configurable on the
# page; here it is the reference against which the assembled dataset is checked
# for producing a signal at all.
SIGNAL_THRESHOLD = 3.0     # %

# One department deliberately breaks out of the corridor.
#
# A payroll budget does not break evenly, it breaks in one place. Sales, where
# variable pay follows the sales plan, runs at +9-10% while the rest stay inside
# ±1-3%: the threshold is cleared by one visible culprit, and the other
# departments stay plausible.
BLOWOUT_DEPARTMENT = "Продажи"
BLOWOUT_MIN = 9.0          # %
BLOWOUT_MAX = 10.0         # %

# --- Seasonality of the plan ------------------------------------------------
#
# Monthly payroll is uneven, and the reason is not bonuses: those are smoothed by
# accruals. The reason is holiday reserves. The start and end of the year run
# higher, with a gentle dip over the summer: headcount is thin and employees take
# the leave they have accrued.
#
# The coefficients are normalised to sum to exactly 12, so annual payroll is
# unchanged and seasonality only redistributes money inside the year. The swing is
# ±3-4% around the monthly average.
SEASONALITY = {
    1: 1.038,   # Jan  full headcount after the holidays, leave reserve building
    2: 1.035,   # Feb
    3: 1.022,   # Mar
    4: 1.005,   # Apr
    5: 0.986,   # May  start of the holiday season
    6: 0.968,   # Jun
    7: 0.962,   # Jul  the floor: thin headcount, peak use of accrued leave
    8: 0.964,   # Aug
    9: 0.985,   # Sep  back from holiday
    10: 1.006,  # Oct
    11: 1.014,  # Nov
    12: 1.015,  # Dec  hiring up towards year end
}

DEPARTMENT_SIGN = {
    # Most departments overspend and one underspends. The mix of directions is
    # deliberate: if every department overspent, the table would be uniform and
    # there would be no way to see that the sorting works.
    "Разработка":         +1,
    "Продажи":            +1,
    "Операции":           +1,
    "Производство":       +1,
    "Маркетинг":          +1,
    "Поддержка клиентов": +1,
    "Юристы":             -1,   # underspends: an open vacancy
}


def load_source():
    """Reads the source dataset. Never modifies it."""
    path = os.path.join(PAYGAP, "demo-data.csv")
    with open(path, encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


def annual_payroll_by_department(rows):
    """Annual payroll by department: base plus variable pay, as actually paid.

    No FTE normalisation is applied here: a budget is executed in real money,
    not in full-time equivalents.
    """
    agg = {}
    for r in rows:
        dep = r["category"]
        pay = float(r["base_salary"]) + float(r["variable_pay"])
        agg[dep] = agg.get(dep, 0.0) + pay
    return agg


def build():
    rng = random.Random(SEED)

    rows = load_source()
    payroll = annual_payroll_by_department(rows)

    departments = sorted(payroll, key=lambda d: -payroll[d])

    # --- Plan: annual payroll spread across months ---
    # Plan figures are never assigned: they are computed from the source dataset.
    # Seasonality redistributes the total inside the year without changing the
    # annual result (the coefficients are normalised to 12).
    plan_annual = {}
    plan_month = {}
    for dep in departments:
        annual = payroll[dep]
        plan_annual[dep] = annual
        plan_month[dep] = {m: annual / 12.0 * SEASONALITY[m] for m in range(1, 13)}

    # --- Actuals: stable sign per department, |deviation| 1-3% per month ---
    # Each department's deviation is drawn from the corridor and is NOT scaled to
    # hit a target total: the corridor is given, the total accumulates.
    # Overspending departments draw from the upper half of the corridor, the
    # underspending one from the lower half: both stay inside ±1-3%.
    dev = {}
    for dep in departments:
        if dep == BLOWOUT_DEPARTMENT:
            magnitude = rng.uniform(BLOWOUT_MIN, BLOWOUT_MAX)
        elif DEPARTMENT_SIGN[dep] > 0:
            magnitude = rng.uniform(2.0, MONTHLY_DEV_MAX)
        else:
            magnitude = rng.uniform(MONTHLY_DEV_MIN, 1.5)
        dev[dep] = DEPARTMENT_SIGN[dep] * magnitude

    # --- Assembling the rows ---
    out = []
    for dep in departments:
        for idx, month_name in enumerate(MONTHS, start=1):
            # The plan is rounded BEFORE actuals and forecast are computed, not
            # after. Otherwise the derived figures come from an unrounded plan
            # while the rounded one goes into the file, and the Excel counterpart
            # reading the plan from a cell can no longer reproduce them.
            plan = round(plan_month[dep][idx], 2)

            if idx <= FACT_MONTHS:
                # Month-to-month noise around a stable sign: a department's
                # deviation need not repeat to the kopeck every month, but it
                # does not change direction.
                jitter = rng.uniform(-0.15, 0.15)
                actual_dev = dev[dep] + jitter
                if (actual_dev > 0) != (dev[dep] > 0):
                    actual_dev = dev[dep]  # the sign matters more than the noise
                fact = plan * (1 + actual_dev / 100)
                forecast = ""
            else:
                fact = ""
                # The forecast carries the department's own run rate: a
                # department that overspent for eight months running keeps
                # overspending. A plain "forecast = plan" would silently assume
                # a return to plan exactly from September.
                forecast = plan * (1 + dev[dep] / 100)

            out.append({
                "department": dep,
                "month": month_name,
                "month_no": idx,
                "plan": round(plan, 2),
                "fact": round(fact, 2) if fact != "" else "",
                "forecast": round(forecast, 2) if forecast != "" else "",
            })

    return out, plan_month, plan_annual, dev, departments


def verify(out, plan_month, plan_annual, dev, departments):
    """A check has to be able to fail, otherwise it checks nothing.

    CHECKS counts the checks below. It is kept next to them rather than printed
    as a constant at the end: a hard-coded "6 of 6" once survived the addition of
    two more checks and quietly lied.
    """
    CHECKS = 8
    failures = []

    # 1. The eight-month result is an overspend, of a plausible size
    plan_8m = sum(r["plan"] for r in out if r["month_no"] <= FACT_MONTHS)
    fact_8m = sum(r["fact"] for r in out if r["month_no"] <= FACT_MONTHS and r["fact"] != "")
    overspend_pct = (fact_8m - plan_8m) / plan_8m * 100
    if overspend_pct <= 0:
        failures.append(
            f"8M is not an overspend but an underspend of {-overspend_pct:.2f}%")
    elif not (EXPECTED_OVERSPEND_MIN <= overspend_pct <= EXPECTED_OVERSPEND_MAX):
        failures.append(
            f"8M overspend {overspend_pct:.2f}% outside the expected "
            f"{EXPECTED_OVERSPEND_MIN}-{EXPECTED_OVERSPEND_MAX}%")

    # 2. The sign of the deviation is stable within a department
    for dep in departments:
        signs = set()
        for r in out:
            if r["department"] == dep and r["month_no"] <= FACT_MONTHS:
                signs.add(r["fact"] > r["plan"])
        if len(signs) != 1:
            failures.append(f"{dep}: the deviation changes sign from month to month")

    # 3. Monthly deviation within 1-3% for everyone except the outlier department.
    # The exception is named explicitly and checked separately (check 8): quietly
    # relaxing the bound "for somebody" would let a second such department through.
    for r in out:
        if r["month_no"] <= FACT_MONTHS and r["fact"] != "":
            if r["department"] == BLOWOUT_DEPARTMENT:
                continue
            d = abs(r["fact"] - r["plan"]) / r["plan"] * 100
            if not (MONTHLY_DEV_MIN * 0.8 <= d <= MONTHLY_DEV_MAX * 1.2):
                failures.append(
                    f"{r['department']} {r['month']}: deviation {d:.2f}% out of bounds")

    # 4. The forecast is exactly the plan carried on the department's run rate
    for r in out:
        if r["forecast"] != "":
            expected = round(r["plan"] * (1 + dev[r["department"]] / 100), 2)
            if abs(r["forecast"] - expected) > 0.005:
                failures.append(
                    f"{r['department']} {r['month']}: forecast {r['forecast']:,.2f} "
                    f"!= run rate {expected:,.2f}")

    # 5. Seasonality does not change the annual plan: the months sum to annual payroll
    for dep in departments:
        annual_from_months = sum(
            r["plan"] for r in out if r["department"] == dep)
        expected = plan_annual[dep]
        if abs(annual_from_months - expected) > 1.0:
            failures.append(
                f"{dep}: plan summed over months {annual_from_months:,.2f} "
                f"!= annual {expected:,.2f} — seasonality moved the total")

    # 6. Shape of the seasonality: summer below winter, floor in July
    jan = sum(r["plan"] for r in out if r["month_no"] == 1)
    jul = sum(r["plan"] for r in out if r["month_no"] == 7)
    dec = sum(r["plan"] for r in out if r["month_no"] == 12)
    if not (jul < jan and jul < dec):
        failures.append("seasonality: July is not below January and December")
    swing = (max(SEASONALITY.values()) - min(SEASONALITY.values())) / 2 * 100
    if not (2.5 <= swing <= 4.5):
        failures.append(f"seasonal swing {swing:.1f}% outside the intended 3-4%")

    # 7. The year including the forecast is above the "problem" threshold.
    # If this comes out green, the dashboard has nothing to show the HR director.
    plan_year_chk = sum(r["plan"] for r in out)
    fc_chk = sum(r["forecast"] for r in out if r["forecast"] != "")
    year_dev = (fact_8m + fc_chk - plan_year_chk) / plan_year_chk * 100
    if year_dev < SIGNAL_THRESHOLD:
        failures.append(
            f"year {year_dev:+.2f}% below the {SIGNAL_THRESHOLD}% threshold — "
            f"the signal block would show 'on track' and there is nothing to show")

    # 8. There is exactly one outlier department, and it really does stand out.
    # Otherwise the badly overspending department dissolves among the others and
    # the signal in the table stops pointing at the culprit.
    out_of_corridor = set()
    for r in out:
        if r["month_no"] <= FACT_MONTHS and r["fact"] != "":
            d = abs(r["fact"] - r["plan"]) / r["plan"] * 100
            if d > MONTHLY_DEV_MAX * 1.2:
                out_of_corridor.add(r["department"])
    if out_of_corridor != {BLOWOUT_DEPARTMENT}:
        failures.append(
            f"exactly {BLOWOUT_DEPARTMENT} should be outside the corridor, "
            f"but these are: {sorted(out_of_corridor) or 'nobody'}")

    return failures, overspend_pct, plan_8m, fact_8m, CHECKS


def main():
    out, plan_month, plan_annual, dev, departments = build()
    failures, overspend_pct, plan_8m, fact_8m, checks = verify(
        out, plan_month, plan_annual, dev, departments)

    path = os.path.join(HERE, "budget-plan-fact.csv")
    with open(path, "w", encoding="utf-8", newline="") as fh:
        writer = csv.DictWriter(
            fh, fieldnames=["department", "month", "month_no", "plan", "fact", "forecast"])
        writer.writeheader()
        writer.writerows(out)

    # The deviations used by the forecast are published as a separate file. The
    # Excel counterpart reads them from here rather than deriving them back from
    # the actuals: the actuals carry month-to-month noise, and deriving backwards
    # gives an average instead of the original figure.
    dev_path = os.path.join(HERE, "department-deviation.csv")
    with open(dev_path, "w", encoding="utf-8", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(["department", "run_rate_deviation_pct"])
        for dep in departments:
            w.writerow([dep, f"{dev[dep]:.10f}"])

    print(f"Written: {path}")
    print(f"Written: {dev_path}")
    print(f"Rows: {len(out)} ({len(departments)} departments x 12 months)")
    print()
    print("Deviation of actuals from plan over 8 months, by department:")
    for dep in departments:
        p = sum(r["plan"] for r in out if r["department"] == dep and r["month_no"] <= FACT_MONTHS)
        f = sum(r["fact"] for r in out if r["department"] == dep and r["month_no"] <= FACT_MONTHS and r["fact"] != "")
        mark = "underspend" if f < p else "overspend "
        print(f"  {dep:20s} {mark} {abs(f - p):14,.2f}  ({(f - p) / p * 100:+.2f}%)")
    print()
    print(f"  {'TOTAL plan 8M':20s} {plan_8m:16,.2f}")
    print(f"  {'TOTAL actual 8M':20s} {fact_8m:16,.2f}")
    print(f"  {'overspend':20s} {fact_8m - plan_8m:16,.2f}  ({overspend_pct:+.2f}%)")

    # The annual result: the dashboard's headline number
    plan_year = sum(r["plan"] for r in out)
    fc_total = sum(r["forecast"] for r in out if r["forecast"] != "")
    expected_year = fact_8m + fc_total
    delta = expected_year - plan_year
    plan_4m = sum(r["plan"] for r in out if r["month_no"] > FACT_MONTHS)
    print()
    print(f"Sep-Dec: plan {plan_4m:,.2f} | forecast {fc_total:,.2f} | "
          f"{(fc_total - plan_4m) / plan_4m * 100:+.2f}%")
    print(f"YEAR: plan {plan_year:,.2f} | expected {expected_year:,.2f} | "
          f"{delta:+,.2f} ({delta / plan_year * 100:+.2f}%)")
    print(f"'Problem' threshold {SIGNAL_THRESHOLD}%: "
          f"{'REACHED' if delta / plan_year * 100 >= SIGNAL_THRESHOLD else 'NOT REACHED'}")

    print()
    if failures:
        print(f"CHECKS: {len(failures)} violations out of {checks} checks")
        for f in failures:
            print(f"  x {f}")
        sys.exit(1)
    print(f"CHECKS: all {checks} passed")


if __name__ == "__main__":
    main()
