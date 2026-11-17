# -*- coding: utf-8 -*-
"""
Compares the recalculated workbook with reference_calc.py, figure by figure.

    python check_model.py [positions.csv] [hiring-history.csv] [workbook.xlsx]

Run recalc.ps1 first: the values are read as Excel computed them.
Exit 1 on any difference over one kopeck.
"""
import sys

from openpyxl import load_workbook

from reference_calc import (GRADES, MODEL, SAMPLE_HISTORY, SAMPLE_POSITIONS, analyse,
                            load_history, load_positions)

SRC = sys.argv[1] if len(sys.argv) > 1 else SAMPLE_POSITIONS
HIST = sys.argv[2] if len(sys.argv) > 2 else SAMPLE_HISTORY
BOOK = sys.argv[3] if len(sys.argv) > 3 else MODEL
TOL = 0.01

R = analyse(load_positions(SRC), load_history(HIST))
wb = load_workbook(BOOK, data_only=True)
fails = []
checked = [0]


def same(label, want, got, tol=TOL):
    checked[0] += 1
    if got is None or isinstance(got, str) or abs(float(got) - float(want)) > tol:
        fails.append(f"{label}: workbook {got!r}, reference {want!r}")


# Formula errors anywhere, and an unrecalculated file.
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("#") and c.value.endswith(("!", "?", "A")):
                fails.append(f"{ws.title}!{c.coordinate} = {c.value}")


def named(n):
    dest = list(wb.defined_names[n].destinations)[0]
    return wb[dest[0]][dest[1].replace("$", "")].value


if named("FullYear") is None:
    sys.exit("The workbook has no computed values: run recalc.ps1 first.")

# Timing table
ws = wb["Настройки"]
gt = next(r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value == "Грейд")
for k, g in enumerate(GRADES):
    t = R["timing"][g]
    same(f"timing {g} n", t["n"], ws.cell(gt + 1 + k, 4).value)
    same(f"timing {g} lag", t["lag"], ws.cell(gt + 1 + k, 7).value)
    same(f"timing {g} ttf", t["ttf"], ws.cell(gt + 1 + k, 8).value)

# Totals
for n, want in [("FullYear", R["full_year"]), ("CostRequested", R["cost"][0]), ("CostRealistic", R["cost"][4]),
                ("CostExpected", R["expected"]), ("LeaversEffect", -R["leavers"]),
                ("ExistingNet", R["existing"] - R["leavers"]),
                ("CostAfterMoves", R["after_moves"]), ("GapToClose", R["gap"]), ("GapAfter", R["gap_after"]),
                ("Capacity", R["capacity"]), ("NextYearCost", R["next_after"])]:
    same(n, want, named(n))

# Summary rows by label
sm = wb["Итоги"]
lab = {}
for r in range(1, sm.max_row + 1):
    v = sm.cell(r, 1).value
    if isinstance(v, str):
        lab.setdefault(v.strip(), r)
for label, want in [
    ("Позиций в плане", R["roles"]),
    ("согласование и публикация вакансии", R["steps"][1]), ("срок до выхода по грейду", R["steps"][2]),
    ("кандидат уже в воронке", R["steps"][3]), ("привязка к запуску проекта", R["steps"][4]),
    ("позиции, которые не закроются в бюджетном году", R["unfilled"]),
    ("перенос выхода позиций, которые могут подождать", R["moves"]),
    ("Общий ФОТ, заявлено бизнесом", R["total_requested"]), ("Общий ФОТ, реалистичный прогноз", R["total_realistic"]),
    ("Общий ФОТ после переноса выхода", R["total_after"]),
    ("в том числе переходящий эффект (аннуализация)", R["carry"] + R["new_next"]),
    ("в том числе уже в бюджете года, новые даты выхода", R["after_moves"]),
]:
    same(label, want, sm.cell(lab[label], 2).value)
same("2028 bridge adds up", R["next_after"], R["after_moves"] + R["carry"] + R["new_next"])

# Curve
r0 = lab["Месяц"]
for k in range(12):
    for j, key in enumerate(["requested", "realistic", "after_moves"]):
        same(f"curve {key} month {k + 1}", R["curve"][key][k], sm.cell(r0 + 1 + k, 2 + j).value)

# Departments
d0 = lab["Подразделение"]
for k, d in enumerate(R["depts"]):
    r = d0 + 1 + k
    if sm.cell(r, 1).value != d["department"]:
        fails.append(f"department row {r}: {sm.cell(r, 1).value!r}, expected {d['department']!r}")
        continue
    for j, key in enumerate(["roles", "requested", "realistic", "after_moves", "moved", "full_year"]):
        same(f"{d['department']} {key}", d[key], sm.cell(r, 2 + j).value)

# Row by row: cost month, move and cost
wp = wb["План"]
head = {wp.cell(4, c).value: c for c in range(1, wp.max_column + 1)}
step_cols = {}
for c in range(1, wp.max_column + 1):
    h = wp.cell(4, c).value
    step_cols.setdefault(h, []).append(c)
c_month = step_cols["Реалистичные даты выхода"][0]   # step 4: realistic cost month
c_cost = step_cols["Реалистичные даты выхода"][1]    # step 5: realistic cost
c_w = head["Закрываемость"]
c_shift = head["Перенос выхода, мес."]
c_cost5 = head["Стоимость в бюджетном году после переноса выхода"]
c_ann = head["Полная годовая стоимость"]
for i, row in enumerate(R["rows"]):
    r = 5 + i
    same(f"row {row['id']} annual", row["annual"], wp.cell(r, c_ann).value)
    same(f"row {row['id']} realistic month", row["m"][4], wp.cell(r, c_month).value)
    same(f"row {row['id']} realistic cost", row["cost"][4], wp.cell(r, c_cost).value)
    same(f"row {row['id']} share filled", row["w"], wp.cell(r, c_w).value)
    same(f"row {row['id']} shift", row["shift"], wp.cell(r, c_shift).value)
    same(f"row {row['id']} cost after moves", row["cost5"], wp.cell(r, c_cost5).value)

if fails:
    print(f"MISMATCHES: {len(fails)} of {checked[0]} figures")
    for f in fails[:40]:
        print("  " + f)
    sys.exit(1)
print(f"OK: {checked[0]} figures agree with the reference calculation to the kopeck "
      f"({R['roles']} roles; gap {R['gap']:,.2f}, closed by {R['moved_roles']} moves; gap after {R['gap_after']:,.2f}).")
