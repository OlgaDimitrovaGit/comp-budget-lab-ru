# -*- coding: utf-8 -*-
"""
Builds vacancy-calendar-model.xlsx (Russian edition) from a CSV of open roles and a hiring history.

Every figure is a live formula. The Plan sheet runs left to right in numbered
steps, one simple operation per column; rows 1-4 carry the step, a note, a spare
row and the column header. Yellow cells are inputs. Labels are in Russian.

    python build_model.py [positions.csv] [hiring-history.csv] [out.xlsx]
"""
import sys

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.workbook.defined_name import DefinedName

from reference_calc import (DEFAULTS, GRADES, MODEL, SAMPLE_HISTORY, SAMPLE_POSITIONS,
                            load_history, load_positions)

SRC = sys.argv[1] if len(sys.argv) > 1 else SAMPLE_POSITIONS
HIST = sys.argv[2] if len(sys.argv) > 2 else SAMPLE_HISTORY
OUT = sys.argv[3] if len(sys.argv) > 3 else MODEL

INPUT = PatternFill("solid", fgColor="FFF4CC")
STEP = PatternFill("solid", fgColor="DCE6F0")
TOTAL = PatternFill("solid", fgColor="EAF1E4")
BOLD = Font(bold=True)
GREY = Font(italic=True, color="7F7F7F")
WRAP = Alignment(wrap_text=True, vertical="top")
RUB = '#,##0.00" ₽"'
RUB0 = '#,##0" ₽"'
DATE = "dd.mm.yyyy"
MONTH = "mmmm yyyy"
MONTHS = '0 "мес."'

# Sheet names; quoted in formulas because they are Cyrillic.
SH_IN, SH_HIST, SH_PLAN, SH_SUM, SH_METHOD = "Настройки", "История", "План", "Итоги", "Метод"


def q(sheet):
    return f"'{sheet}'!"


# Values the reader types on the Plan sheet; the calculation keeps English stage codes inside.
STAGE_RU = {"interview": "собеседование", "offer": "оффер", "accepted": "принят"}
YES, NO = "да", "нет"

positions = load_positions(SRC)
history = load_history(HIST)
S = DEFAULTS
CAP = max(1000, len(positions))
CAPH = max(300, len(history))
wb = Workbook()


def name(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


# ---------------------------------------------------------------- Inputs
ws = wb.active
ws.title = SH_IN
ws["A1"] = "Настройки: меняйте жёлтые ячейки"
ws["A1"].font = Font(bold=True, size=13)
ws.append([])
ws.append(["Параметр", "Значение", "Пояснение"])
for c in ws[3]:
    c.font = BOLD

settings = [
    ("BudgetYear", "Бюджетный год", S["budget_year"], "0", True,
     "Год, на который планируется бюджет найма."),
    ("YearStart", "Первый день бюджетного года", "=DATE(BudgetYear,1,1)", DATE, False,
     "Раньше этой даты в плане никто не выходит."),
    ("AsOf", "Дата плана", S["as_of"], DATE, True,
     "Дата составления плана. От неё считаем выход кандидатов, которые уже на собеседованиях или с оффером."),
    ("DaysInterview", "Дней до выхода: кандидат на собеседованиях", S["stage_days"]["interview"], "0", True,
     "От даты плана до первого рабочего дня, если кандидат на собеседованиях."),
    ("DaysOffer", "Дней до выхода: оффер отправлен", S["stage_days"]["offer"], "0", True,
     "От даты плана до первого рабочего дня, если оффер отправлен. У принятого оффера своя дата выхода."),
    ("MinHistory", "Минимум закрытых вакансий по грейду в истории", S["min_history"], "0", True,
     "Если в листе «История» по грейду меньше закрытых вакансий, берутся сроки по умолчанию из таблицы ниже."),
    ("MaxShift", "Предел сдвига выхода", S["max_shift"], MONTHS, True,
     "Не сдвигаем: критичные позиции, позиции с кандидатом в воронке, позиции под запуск проекта."),
    ("ExistingPayroll", "ФОТ текущего состава за бюджетный год", S["existing_payroll"], RUB, True,
     "Полная стоимость уже работающих за весь бюджетный год, ФОТ со страховыми взносами, из прогноза ФОТ. "
     "В демо число смоделировано."),
    ("Target", "Лимит ФОТ на бюджетный год", S["target"], RUB, True,
     "Утверждённый лимит ФОТ со страховыми взносами на текущий состав и новых сотрудников вместе. "
     "В демо число смоделировано."),
    ("Attrition", "Выбытие текущего состава за год, %", S["attrition"] / 100, "0.0%", True,
     "Доля текущих сотрудников, которые уйдут за год. Условное значение: подставьте свою текучесть."),
    ("VacantMonths", "Месяцев место ушедшего пустует", S["vacant_months"], '0.0 "мес."', True,
     "В среднем, до выхода замены. Пустые месяцы экономят ФОТ текущего состава. Условное значение."),
    ("FillRate", "Доля вакансий, закрытых в бюджетном году", S["fill_rate"] / 100, "0%", True,
     "Доля позиций без кандидата, которые реально выйдут в бюджетном году. Позиции с кандидатом в воронке считаем "
     "полностью. Незакрытую в этом году позицию на следующий год согласуют заново. Условное значение: подставьте "
     "свою историю."),
    ("RateBelow", "Ставка страховых взносов до предельной базы", S["rate_below"] / 100, "0.0%", True,
     "Россия: единый тариф 30% до предельной базы (п. 3 ст. 425 НК РФ). "
     "Взносы на травматизм не включены: их ставка зависит от класса риска, добавьте её сюда при необходимости. "
     "Ставки и база на 2027 год ещё не утверждены: взяты значения 2026 года."),
    ("RateAbove", "Ставка страховых взносов сверх предельной базы", S["rate_above"] / 100, "0.0%", True,
     "Россия: 15,1% с выплат сверх предельной базы (п. 3 ст. 425 НК РФ)."),
    ("Ceiling", "Предельная база в год на человека (пусто: нет)", S["ceiling"], RUB, True,
     "Россия, 2026: 2 979 000 ₽ (Постановление Правительства РФ от 31.10.2025 № 1705). "
     "База у каждого работодателя своя и копится с первого месяца работы человека, нарастающим итогом "
     "(ст. 421 НК РФ): доход на прошлом месте её не уменьшает. Кто выходит поздно, может не дойти до базы и "
     "платит 30% со всей оплаты, модель так и считает. Для доли ставки база сравнивается с оплатой одного "
     "человека, а результат умножается на долю."),
]
for i, (n, label, val, fmt, is_input, note) in enumerate(settings):
    r = 4 + i
    ws.cell(r, 1, label)
    c = ws.cell(r, 2, val)
    c.number_format = fmt
    if is_input:
        c.fill = INPUT
    ws.cell(r, 3, note).alignment = WRAP
    name(n, f"{q(SH_IN)}$B${r}")

# Timing by grade: medians from the History sheet, defaults where the history is thin.
GT = 4 + len(settings) + 2
ws.cell(GT - 1, 1, "Сроки выхода по грейдам").font = BOLD
heads = ["Грейд", "По умолчанию: заявка → публикация, дней", "По умолчанию: публикация → выход, дней",
         "Закрытых вакансий в истории", "Медиана истории: заявка → публикация", "Медиана истории: публикация → выход",
         "В расчёте: заявка → публикация, дней", "В расчёте: публикация → выход, дней"]
for j, h in enumerate(heads):
    c = ws.cell(GT, 1 + j, h)
    c.font = BOLD
    c.alignment = WRAP
HG, HL, HT = (f"{q(SH_HIST)}${L}$2:${L}${CAPH + 1}" for L in "BGH")
for k, g in enumerate(GRADES):
    r = GT + 1 + k
    ws.cell(r, 1, g).fill = INPUT  # rename to your own grades if your scale differs
    for j, v in ((2, S["default_lag"][g]), (3, S["default_ttf"][g])):
        ws.cell(r, j, v).fill = INPUT
    ws.cell(r, 4, f'=COUNTIFS({HG},A{r},{HT},"<>")')
    # _xlfn. marks functions newer than Excel 2007 in the file format; without it Excel shows #NAME?.
    ws.cell(r, 5, f'=IF(COUNTIFS({HG},A{r},{HL},"<>")=0,"",_xlfn.AGGREGATE(17,6,{HL}/(({HG}=A{r})*({HL}<>"")),2))')
    ws.cell(r, 6, f'=IF(D{r}=0,"",_xlfn.AGGREGATE(17,6,{HT}/(({HG}=A{r})*({HT}<>"")),2))')
    ws.cell(r, 7, f'=IF(AND(D{r}>=MinHistory,E{r}<>""),ROUND(E{r},0),B{r})')
    ws.cell(r, 8, f'=IF(D{r}>=MinHistory,ROUND(F{r},0),C{r})')
    for j in (5, 6):
        ws.cell(r, j).number_format = "0.0"
name("GradeList", f"{q(SH_IN)}$A${GT + 1}:$A${GT + len(GRADES)}")
name("UsedLag", f"{q(SH_IN)}$G${GT + 1}:$G${GT + len(GRADES)}")
name("UsedTTF", f"{q(SH_IN)}$H${GT + 1}:$H${GT + len(GRADES)}")
ws.cell(GT + len(GRADES) + 1, 1,
        "Медианы берём из листа «История» (выгрузка закрытых вакансий в формате sample-hiring-history.csv). "
        "Срок до выхода считаем до первого рабочего дня, а не до принятого оффера: оплата начинается с первого дня. "
        "В таблице вся шкала из 25 грейдов. Сроки по умолчанию условные: заданы для демо-грейдов G10–G18; у нечётных "
        "среднее между соседями, ниже G10 и выше G18 как у ближайшего демо-грейда. Названия грейдов жёлтые: если шкала "
        "другая, переименуйте. Если грейда нет в таблице, на листе «План» будет #Н/Д.").alignment = WRAP
ws.merge_cells(start_row=GT + len(GRADES) + 1, start_column=1, end_row=GT + len(GRADES) + 1, end_column=3)
ws.row_dimensions[GT + len(GRADES) + 1].height = 75
ws.column_dimensions["A"].width = 46
ws.column_dimensions["B"].width = 20
ws.column_dimensions["C"].width = 80
for col in "DEFGH":
    ws.column_dimensions[col].width = 16

# ---------------------------------------------------------------- History
wh = wb.create_sheet(SH_HIST)
hcols = ["requisition", "grade", "department", "requested", "opened", "started",
         "request_to_open_days", "open_to_start_days"]
wh.append(["Заявка", "Грейд", "Подразделение", "Дата заявки", "Дата публикации", "Дата выхода",
           "Заявка → публикация, дней", "Публикация → выход, дней"])
for c in wh[1]:
    c.font = BOLD
    c.alignment = WRAP
for r in history:
    vals = []
    for k in hcols:
        v = (r.get(k) or "").strip()
        if k.endswith("_days") or k == "requisition":
            v = float(v) if v else None
        vals.append(v)
    wh.append(vals)
for col in "ABCDEFGH":
    wh.column_dimensions[col].width = 16
wh.row_dimensions[1].height = 30
wh.freeze_panes = "A2"

# ---------------------------------------------------------------- Plan
# Contributions on the annual pay of the role: rate below the ceiling, a lower rate above it.
CONTRIB = 'IF(Ceiling="",RateBelow*{x},RateBelow*MIN({x},Ceiling)+RateAbove*MAX(0,{x}-Ceiling))'
CM = "(YEAR({d})-BudgetYear)*12+MONTH({d})+(DAY({d})>15)"
IN1 = "MAX(0,MIN(12,13-{m}))"
IN2 = "MAX(0,MIN(12,25-{m}))"
# No ceiling = a ceiling no pay reaches; keeps the cost formula free of IF so it also works on arrays.
name("CeilEff", 'IF(Ceiling="",1E+15,Ceiling)')


def YC(n, pay="{pay}", fte="{fte}"):
    """Cost of a row paid for n months of one year: FTE x (pay for n months plus contributions on it).
    The base runs from the first paid month, so a late starter may stay under the ceiling.
    Written without MIN/MAX so it works element by element on arrays."""
    x = f"({pay}*({n})/12)"
    return f"{fte}*({x}*(1+RateBelow)-(RateBelow-RateAbove)*({x}>CeilEff)*({x}-CeilEff))"


SHIFTS = "{{1,2,3,4,5,6,7,8,9,10,11}}"  # candidate moves; a move is at most 11 months (12 leaves the year)
ST_I, ST_O, ST_A = (STAGE_RU[k] for k in ("interview", "offer", "accepted"))

INPUT_COLS = [
    ("id", "Код позиции", None, "Код позиции (position_id в CSV)."),
    ("dept", "Подразделение", None, "Подразделение."),
    ("grade", "Грейд", None, "Грейд компании, как в таблице сроков на листе «Настройки». В демо шкала из 25 грейдов, "
     "заняты только чётные: от G10 (младший специалист) до G18 (руководитель направления)."),
    ("base", "Годовой оклад", RUB0, "Годовой оклад до вычета НДФЛ."),
    ("bonus", "Целевая премия, %", "0.0", "Целевая премия, % от оклада."),
    ("req", "Месяц по запросу", MONTH, "Месяц, который попросил руководитель (1-е число)."),
    ("stage", "Этап кандидата", None,
     f"Пусто или {ST_I} / {ST_O} / {ST_A}, если кандидат уже в воронке."),
    ("exp", "Дата выхода по офферу", DATE, "Для принятого оффера: согласованная дата выхода."),
    ("proj", "Запуск проекта", DATE, "Для позиции под проект: дата запуска. Раньше неё позиция не выходит."),
    ("crit", "Критичная", None, f"{YES}: позиция не сдвигается."),
    ("fte_in", "Доля ставки (FTE)", "0.0#", "Сколько ставок в строке: 1 — одна, 0,5 — полставки, 10 — десять "
     "одинаковых позиций. Пусто = 1. Стоимость и взносы умножаются на долю."),
]

VARIANTS = ["По запросу", "+ согласование", "+ срок до выхода", "+ воронка", "Реалистично"]

CALC = [
    ("Шаг 1. Полная годовая стоимость", [
        ("fte", "Доля ставки в расчёте", "0.0#", '=IF({fte_in}="",1,{fte_in})', "Пусто во вводе = 1."),
        ("pay", "Годовая оплата одной ставки (оклад + премия)", RUB, "={base}*(1+{bonus}/100)",
         "Оклад плюс целевая премия, на одного человека."),
        ("con", "Страховые взносы за полный год", RUB, "={fte}*" + CONTRIB.format(x="{pay}"),
         "Ставка до предельной базы и ставка сверх неё (лист «Настройки»), на одного человека × доля ставки."),
        ("ann", "Полная годовая стоимость", RUB, "={fte}*{pay}+{con}",
         "Стоимость позиции за полный год: доля ставки × оплата плюс взносы."),
        ("mon", "Стоимость в месяц", RUB, "={ann}/12", "Полная годовая ÷ 12."),
    ]),
    ("Шаг 2. Сроки по грейду", [
        ("lag", "Заявка → публикация, дней", "0", '=IF({stage}<>"",0,INDEX(UsedLag,MATCH({grade},GradeList,0)))',
         "Из таблицы сроков на листе «Настройки». 0, если кандидат уже в воронке."),
        ("ttf", "Публикация → выход, дней", "0", '=IF({stage}<>"",0,INDEX(UsedTTF,MATCH({grade},GradeList,0)))',
         "Та же таблица."),
    ]),
    ("Шаг 3. Дата выхода, по одному эффекту", [
        ("s0", "По запросу", DATE, "={req}", "Месяц, когда руководитель хочет видеть человека."),
        ("s1", "+ согласование", DATE, "=MAX({s0},YearStart+{lag})",
         "Найм можно начать заранее, но не раньше открытия года: не раньше 1 января плюс дни на согласование."),
        ("s2", "+ срок до выхода", DATE, "=MAX({s0},YearStart+{lag}+{ttf})",
         "Не раньше 1 января плюс дни на согласование и до первого рабочего дня."),
        ("s3", "+ воронка", DATE,
         f'=IF({{stage}}="",{{s2}},IF({{stage}}="{ST_A}",MAX(YearStart,{{exp}}),'
         f'MAX(YearStart,AsOf+IF({{stage}}="{ST_I}",DaysInterview,DaysOffer))))',
         "Кандидат в воронке: дата плана плюс дни по этапу или согласованная дата."),
        ("s4", "+ запуск проекта = реалистичный выход", DATE, '=IF({proj}="",{s3},MAX({proj},{s3}))',
         "Не раньше запуска проекта."),
    ]),
    ("Шаг 4. Месяц, с которого платим", [
        (f"m{k}", h, "0", "=" + CM.format(d="{s%d}" % k),
         "Месяц 1 = январь бюджетного года. Выход после 15-го считается со следующего месяца." if k == 0 else "")
        for k, h in enumerate(VARIANTS)
    ]),
    ("Шаг 5. Стоимость в бюджетном году", [
        (f"c{k}", h, RUB, "=" + YC(IN1.format(m="{m%d}" % k)),
         "Доля ставки × (оплата за оплачиваемые месяцы + взносы на неё). База взносов считается с месяца выхода "
         "нарастающим итогом: кто выходит поздно, может не дойти до предельной базы." if k == 0 else "")
        for k, h in enumerate(VARIANTS)
    ] + [
        ("w", "Доля закрытия", "0%", '=IF({stage}<>"",1,FillRate)',
         "1, если кандидат в воронке; иначе доля закрытия с листа «Настройки»."),
        ("e4", "Реалистичный прогноз", RUB, "={c4}*{w}",
         "Реалистичная стоимость × доля закрытия. Дальше все суммы взвешены этой долей."),
        ("aw", "Полная годовая × доля закрытия", RUB, "={ann}*{w}", "Полная годовая, взвешенная."),
    ]),
    ("Шаг 6. Сдвиги, чтобы уложиться в лимит", [
        ("elig", "Можно сдвинуть (1/0)", "0",
         f'=IF(AND(NOT(OR(LOWER({{crit}}&"")={{{{"{YES}","yes","y","1","true"}}}})),{{stage}}="",{{proj}}=""),1,0)',
         "Не критичная, без кандидата в воронке, не под проект."),
        ("maxe", "Можно сдвинуть на, мес.", "0", "=IF({elig}=1,MIN(MaxShift," + IN1.format(m="{m4}") + "),0)",
         "Предел с листа «Настройки», но не больше месяцев, оставшихся в бюджетном году: позиция, сдвинутая "
         "в следующий год, больше не экономит."),
        ("savx", "Наибольшая экономия", RUB,
         "={w}*({c4}-" + YC(IN1.format(m="{m4}") + "-{maxe}") + ")",
         "Реалистичная стоимость минус стоимость при сдвиге до предела, × доля закрытия. Взносы пересчитаны на "
         "укороченный год."),
        ("key", "Ключ очереди", RUB, "=IF({elig}=1,{mon},-1)",
         "Месячная стоимость; -1 у позиций, которые двигать нельзя."),
        ("rank", "Очередь", "0", "=_xlfn.RANK.EQ({key},{R_key})+COUNTIF({A_key},{key})",
         "Сначала самые дорогие: так сдвигаем меньше людей ради той же экономии. При равенстве по порядку в файле."),
        ("cum", "Экономия позиций раньше в очереди", RUB, '=SUMIFS({R_savx},{R_rank},"<"&{rank})',
         "Сколько экономят позиции раньше в очереди, если сдвинуть каждую до её предела."),
        ("shift", "Сдвиг, мес.", "0",
         "=IF(AND({elig}=1,{maxe}>0,{cum}<GapToClose,{mon}*{w}>0),MIN({maxe},1+SUMPRODUCT((" + SHIFTS
         + "<{maxe})*({w}*({c4}-" + YC(IN1.format(m="{m4}") + "-" + SHIFTS) + ")<GapToClose-{cum}))),0)",
         "Наименьшее целое число месяцев, которое закрывает остаток разрыва, но не больше предела позиции. "
         "Формула перебирает сдвиги от 1 месяца и считает, сколько из них ещё не закрывают остаток."),
        ("m5", "Месяц, с которого платим, после сдвига", "0", "={m4}+{shift}", ""),
        ("c5", "Стоимость в бюджетном году после сдвигов", RUB, "={w}*" + YC(IN1.format(m="{m5}")),
         "Взвешено долей закрытия."),
        ("sav", "Эффект сдвига (минус — экономия)", RUB, "={c5}-{e4}",
         "Стоимость после сдвига минус реалистичный прогноз: минус — экономия."),
    ]),
    ("Шаг 7. Следующий год", [
        ("n4", "Следующий год, реалистично", RUB, "={w}*" + YC(IN2.format(m="{m4}")),
         "Стоимость в следующем году × доля закрытия. Позиция, вышедшая в бюджетном году, оплачивается все 12 "
         "месяцев; база взносов в новом году считается заново с января."),
        ("n5", "Следующий год, после сдвигов", RUB, "={w}*" + YC(IN2.format(m="{m5}")), ""),
        ("p4", "Платим с, реалистично", MONTH, "=DATE(BudgetYear,{m4},1)", "Первый оплачиваемый месяц."),
        ("p5", "Платим с, после сдвигов", MONTH, "=DATE(BudgetYear,{m5},1)", ""),
    ]),
]

wp = wb.create_sheet(SH_PLAN, 1)
DATA0 = 5
col_of = {}
for j, (key, *_rest) in enumerate(INPUT_COLS):
    col_of[key] = get_column_letter(1 + j)
ci = len(INPUT_COLS) + 1
for _step, items in CALC:
    for key, *_rest in items:
        col_of[key] = get_column_letter(ci)
        ci += 1
LAST = DATA0 + CAP - 1

wp.cell(1, 1, "Ваши позиции: вставьте поверх жёлтых столбцов").font = BOLD
for j, (key, header, fmt, note) in enumerate(INPUT_COLS):
    c = 1 + j
    wp.cell(1, c).fill = INPUT
    wp.cell(2, c, note).alignment = WRAP
    wp.cell(4, c, header).font = BOLD
    wp.cell(4, c).alignment = WRAP
    for i in range(CAP):
        cell = wp.cell(DATA0 + i, c)
        cell.fill = INPUT
        if fmt:
            cell.number_format = fmt
    wp.column_dimensions[get_column_letter(c)].width = 14
for i, p in enumerate(positions):
    r = DATA0 + i
    vals = [p["id"], p["department"], p["grade"], p["base"], p["bonus_pct"], p["requested"],
            STAGE_RU.get(p["stage"]) if p["stage"] else None,
            p["expected_start"], p["project_start"], YES if p["critical"] else NO, p["fte"]]
    for j, v in enumerate(vals):
        if v is not None:
            wp.cell(r, 1 + j, v)

# Dropdowns for stage and critical, so a typo cannot slip in by hand.
for key, choices in (("stage", f"{ST_I},{ST_O},{ST_A}"), ("crit", f"{YES},{NO}")):
    dv = DataValidation(type="list", formula1=f'"{choices}"', allow_blank=True, showErrorMessage=True,
                        errorTitle="Выберите из списка", error=f"Допустимо: {choices.replace(',', ', ')}.")
    dv.add(f"{col_of[key]}{DATA0}:{col_of[key]}{LAST}")
    wp.add_data_validation(dv)

ci = len(INPUT_COLS) + 1
for step, items in CALC:
    first = True
    for key, header, fmt, tmpl, note in items:
        if first:
            wp.cell(1, ci, step).font = BOLD
            first = False
        wp.cell(1, ci).fill = STEP
        wp.cell(2, ci, note).alignment = WRAP
        wp.cell(4, ci, header).font = BOLD
        wp.cell(4, ci).alignment = WRAP
        tmpl = '=IF({id}="","",' + tmpl[1:] + ")"
        for i in range(CAP):
            r = DATA0 + i
            refs = {k: f"{v}{r}" for k, v in col_of.items()}
            refs.update({"R_" + k: f"${v}${DATA0}:${v}${LAST}" for k, v in col_of.items()})
            refs.update({"A_" + k: f"${v}${DATA0 - 1}:{v}{r - 1}" for k, v in col_of.items()})
            c = wp.cell(r, ci, tmpl.format(**refs))
            if fmt:
                c.number_format = fmt
        wp.column_dimensions[get_column_letter(ci)].width = 15
        ci += 1
wp.row_dimensions[2].height = 100
wp.row_dimensions[3].height = 6
wp.row_dimensions[4].height = 45
wp.freeze_panes = "B5"


def P(key):
    return f"{q(SH_PLAN)}${col_of[key]}${DATA0}:${col_of[key]}${LAST}"


# ---------------------------------------------------------------- Summary
wsum = wb.create_sheet(SH_SUM, 0)
wsum["A1"] = "Календарь вакансий: сколько открытые позиции стоят в бюджетном году"
wsum["A1"].font = Font(bold=True, size=13)
wsum["A2"] = "Данные смоделированы. Каждая цифра считается живой формулой по листу «План»."
srow = [4]


def section(title):
    r = srow[0]
    wsum.cell(r, 1, title).font = BOLD
    wsum.cell(r, 1).fill = STEP
    wsum.cell(r, 2).fill = STEP
    srow[0] += 1


def line(label, formula, fmt=RUB, note="", nm=None, total=False):
    r = srow[0]
    wsum.cell(r, 1, label)
    c = wsum.cell(r, 2, formula)
    c.number_format = fmt
    if total:
        wsum.cell(r, 1).font = BOLD
        c.font = BOLD
        c.fill = TOTAL
    if note:
        wsum.cell(r, 3, note).alignment = WRAP
    if nm:
        name(nm, f"{q(SH_SUM)}$B${r}")
    srow[0] += 1
    return r


section("БЮДЖЕТНЫЙ ГОД: СТОИМОСТЬ НАЙМА, ПО ОДНОМУ ЭФФЕКТУ")
line("Позиций в плане", f'=COUNTIF({P("id")},"?*")', "0")
r = line("Полная годовая стоимость всех позиций (если бы все работали 12 месяцев)", f"=SUM({P('aw')})", RUB,
         "Справочно, в расчёт ниже не входит. С учётом доли закрытия.", nm="FullYear")
wsum.cell(r, 1).font = GREY  # reference figure, not a step of the bridge: grey italic
wsum.cell(r, 2).font = GREY
line("Бюджет на найм в пределах лимита", "=Target-ExistingPayroll", RUB,
     "Лимит ФОТ минус ФОТ текущего состава (лист «Настройки»).", total=True)
line("Найм, даты выхода по запросу руководителей", f"=SUM({P('c0')})", RUB, nm="CostRequested", total=True)
SIGNED = []  # deviation rows: plus = cost goes up / over target (red), minus = saving (green, bold)
SIGNED.append(line("   согласование и публикация вакансии", f"=SUM({P('c1')})-SUM({P('c0')})", RUB,
                   "В цветных строках отклонений: минус — экономия, плюс — рост."))
SIGNED.append(line("   срок до выхода по грейду", f"=SUM({P('c2')})-SUM({P('c1')})", RUB))
SIGNED.append(line("   кандидат уже в воронке", f"=SUM({P('c3')})-SUM({P('c2')})", RUB))
SIGNED.append(line("   привязка к запуску проекта", f"=SUM({P('c4')})-SUM({P('c3')})", RUB))
line("Найм, реалистичные даты выхода", f"=SUM({P('c4')})", RUB,
     "С учётом сроков подбора, воронки и запуска проектов.", nm="CostRealistic", total=True)
SIGNED.append(line("   позиции, которые не закроются в бюджетном году", f"=SUM({P('e4')})-SUM({P('c4')})", RUB,
                   "Доля закрытия на листе «Настройки». На следующий год их согласуют заново."))
line("Найм, реалистичный прогноз", f"=SUM({P('e4')})", RUB,
     "Реалистичные даты с учётом доли закрытия.", nm="CostExpected", total=True)
SIGNED.append(line("   сдвиги позиций, которые могут подождать", f"=SUM({P('c5')})-SUM({P('e4')})", RUB,
                   "Чтобы уложиться в лимит."))
line("Найм, новые даты выхода после сдвигов", f"=SUM({P('c5')})", RUB, nm="CostAfterMoves", total=True)
srow[0] += 1

section("БЮДЖЕТНЫЙ ГОД: ОБЩИЙ ФОТ И ЛИМИТ")
line("ФОТ текущего состава", "=ExistingPayroll", RUB, "Со страховыми взносами (лист «Настройки»).")
SIGNED.append(line("   пустые места ушедших", "=-ExistingPayroll*Attrition*VacantMonths/12", RUB,
                   "ФОТ текущего состава × выбытие × месяцы без замены ÷ 12.", nm="LeaversEffect"))
line("Текущий состав за вычетом ушедших", "=ExistingPayroll+LeaversEffect", RUB, nm="ExistingNet")
line("Общий ФОТ, даты выхода по запросу", "=ExistingPayroll+CostRequested", RUB, "Без учёта выбытия.")
line("Общий ФОТ, реалистичный прогноз", "=ExistingNet+CostExpected", RUB,
     "Сроки найма, доля закрытия, выбытие.")
line("Лимит ФОТ", "=Target", RUB, "Лист «Настройки».")
SIGNED.append(line("Разрыв к лимиту (плюс — перерасход)", "=ExistingNet+CostExpected-Target", RUB,
                   "Его закрываем сдвигами.", nm="GapToClose", total=True))
line("Наибольшая экономия от более позднего выхода", f"=SUM({P('savx')})", RUB,
     "Показана положительной суммой.", nm="Capacity")
line("Общий ФОТ, новые даты выхода после сдвигов", "=ExistingNet+CostAfterMoves", RUB)
SIGNED.append(line("Разрыв после сдвигов (плюс — всё ещё перерасход)", "=ExistingNet+CostAfterMoves-Target", RUB,
                   "Минус — запас до лимита.", nm="GapAfter", total=True))
# FIXED, not TEXT: TEXT format codes depend on the Excel locale, FIXED groups thousands in any locale.
line("Вывод",
     '=IF(GapToClose<=0,"Реалистичные даты выхода уже укладываются в лимит: сдвигать нечего.",'
     'IF(GapAfter>0.005,"Сдвигов в пределах допустимого не хватает: "&FIXED(GapAfter,0)&" ₽ сверх лимита. '
     'Следующий шаг — договориться, какие позиции снять.",'
     '"Лимит выдержан: сдвинуто позиций — "&COUNTIF(' + P("shift") + ',">0")&"."))', "@")
srow[0] += 1

section("ВЛИЯНИЕ НА ФОТ СЛЕДУЮЩЕГО ГОДА")
line("Найм бюджетного года в ФОТ следующего года", f"=SUM({P('n5')})", RUB,
     "Полный год всех, кто вышел.", nm="NextYearCost", total=True)
line("   в том числе уже в бюджете года, новые даты выхода", "=CostAfterMoves", RUB)
line("   в том числе прирост: переходящий эффект (аннуализация)", "=NextYearCost-CostAfterMoves", RUB,
     "Месяцы, не оплаченные в бюджетном году, потому что люди вышли не с января. Сюда же попадут позиции, "
     "которые выходят только в следующем году, если такие есть.")
wsum.cell(srow[0], 1, "Допущения: в следующем году новых позиций нет, повышений нет, ставки взносов те же.").alignment = WRAP
srow[0] += 2

section("МЕСЯЧНЫЙ ФОТ НОВЫХ СОТРУДНИКОВ, БЮДЖЕТНЫЙ ГОД")
r = srow[0]
for j, h in enumerate(["Месяц", "По запросу", "Реалистично", "После сдвигов"]):
    wsum.cell(r, 1 + j, h).font = BOLD

# Month k of a role = its cost for months paid up to k minus up to k-1, on the running contribution
# base: the month a person crosses the ceiling is cheaper. A start before January counts from January.
# Array formulas (entered as with Ctrl+Shift+Enter): IF skips the empty rows of the Plan sheet.
def curve_cell(k, key, weighted):
    first = f'IF({P(key)}<1,1,{P(key)})'
    paid = lambda kk: f"IF({kk}-{first}+1>0,{kk}-{first}+1,0)"
    one = lambda kk: YC(paid(kk), pay=P("pay"), fte=P("fte"))
    w = f"{P('w')}*" if weighted else ""
    return f'=SUM(IF({P("id")}<>"",{w}({one(k)}-{one(k - 1)}),0))'


for k in range(1, 13):
    rr = r + k
    wsum.cell(rr, 1, f"=DATE(BudgetYear,{k},1)").number_format = MONTH
    for j, (key, weighted) in enumerate([("m0", False), ("m4", True), ("m5", True)]):
        ref = f"{get_column_letter(2 + j)}{rr}"
        wsum[ref] = ArrayFormula(ref, curve_cell(k, key, weighted))
        wsum[ref].number_format = RUB
srow[0] = r + 14

section("ПО ПОДРАЗДЕЛЕНИЯМ, БЮДЖЕТНЫЙ ГОД")
r = srow[0]
for j, h in enumerate(["Подразделение", "Позиций", "По запросу", "Реалистично", "После сдвигов",
                       "Сдвинуто позиций", "Полная годовая"]):
    wsum.cell(r, 1 + j, h).font = BOLD
DEPT0 = r + 1
depts = sorted({p["department"] for p in positions})
for k, d in enumerate(depts):
    rr = DEPT0 + k
    wsum.cell(rr, 1, d).fill = INPUT
    wsum.cell(rr, 2, f'=COUNTIF({P("dept")},A{rr})')
    for j, key in enumerate(["c0", "e4", "c5"]):
        wsum.cell(rr, 3 + j, f'=SUMIFS({P(key)},{P("dept")},A{rr})').number_format = RUB
    wsum.cell(rr, 6, f'=COUNTIFS({P("dept")},A{rr},{P("shift")},">0")')
    wsum.cell(rr, 7, f'=SUMIFS({P("aw")},{P("dept")},A{rr})').number_format = RUB
rr = DEPT0 + len(depts)
wsum.cell(rr, 1, "Итого").font = BOLD
for j in range(2, 8):
    L = get_column_letter(j)
    c = wsum.cell(rr, j, f"=SUM({L}{DEPT0}:{L}{rr - 1})")
    c.font = BOLD
    if j not in (2, 6):
        c.number_format = RUB
wsum.cell(rr + 1, 1, "Названия подразделений вписываете вы: добавьте строку на каждое своё. "
                     "Итог должен совпасть с цифрами выше.").alignment = WRAP
wsum.column_dimensions["A"].width = 62
wsum.column_dimensions["B"].width = 22
wsum.column_dimensions["C"].width = 60
for col in "DEFG":
    wsum.column_dimensions[col].width = 20

# ---------------------------------------------------------------- Method
wm = wb.create_sheet(SH_METHOD)
method = [
    "Метод",
    "Здесь краткая выжимка. Подробное описание метода и обучающая инструкция — по ссылке «Как считаем» на странице "
    "инструмента.",
    "1. Полная стоимость позиции: оклад × (1 + целевая премия) плюс страховые взносы: 30% с выплат до "
    "предельной базы и 15,1% сверх неё. Ставки и базу на листе «Настройки» можно заменить. Доля ставки (FTE) "
    "умножает всё линейно: 0,5 стоит ровно половину ставки, 10 — десять одинаковых позиций в одной строке.",
    "2. Реалистичный выход: месяц руководителя плюс два срока, «заявка → публикация» и «публикация → первый день». "
    "Оба срока берём медианой грейда по вашей истории найма. Кандидат на собеседованиях или с оффером выходит через "
    "заданное число дней от даты плана, принятый оффер — в согласованную дату. Позиция под проект выходит не раньше "
    "запуска.",
    "3. Оплата считается с месяца выхода; выход после 15-го — со следующего месяца. Стоимость в году = оплата за "
    "оплачиваемые месяцы плюс взносы на неё. База взносов копится с месяца выхода, поэтому кто выходит поздно, "
    "может не дойти до предельной базы и платит 30% со всей оплаты. Сдвиг выхода экономит оплату и взносы, "
    "пересчитанные на укороченный год.",
    "4. Эффекты добавляем по одному (лист «План», шаги 3–5), поэтому видно, сколько стоит каждый.",
    "4а. Реалистичный прогноз: не каждая вакансия закрывается в году. Позиции без кандидата взвешиваем долей закрытия "
    "с листа «Настройки», с кандидатом считаем полностью. Ушедшие сотрудники экономят ФОТ, пока место пустует: ФОТ "
    "текущего состава × выбытие × пустые месяцы ÷ 12. Оба значения по умолчанию условные: подставьте свою историю. "
    "Незакрытую в этом году позицию на следующий год согласуют заново.",
    "5. Сдвиги: если реалистичный итог всё ещё выше лимита, сдвигаем позже позиции, которые могут подождать. Начинаем "
    "с самых дорогих, каждую двигаем на наименьшее число месяцев, закрывающее остаток, но не больше предела с листа "
    "«Настройки» и не больше месяцев, оставшихся в бюджетном году. Критичные, с кандидатом и проектные не двигаем. "
    "Если этого мало, вывод так и говорит: следующий шаг — договориться, какие позиции снять.",
    "6. Следующий год: каждая позиция, вышедшая в бюджетном году, оплачивается 12 месяцев. Поэтому сдвиги экономят "
    "только в бюджетном году; финансам нужны оба года.",
    "Упрощения: оклад и премия начисляются равными долями по месяцам; на практике взносы с премии начисляют на "
    "резерв, обычно при 100% выполнения, и итог года тот же. Взносы на травматизм не включены.",
    "Ставки: Россия, 2026 (п. 3 ст. 425 НК РФ; Постановление Правительства РФ от 31.10.2025 № 1705), применены к 2027. "
    "Сроки по умолчанию условные, не рыночные. Все данные в файле смоделированы.",
]
for i, text in enumerate(method):
    c = wm.cell(1 + i, 1, text)
    c.alignment = WRAP
    if i == 0:
        c.font = Font(bold=True, size=13)
wm.column_dimensions["A"].width = 110

# Deviations: plus (cost up, over target) in red, minus (a saving) in bold green.
OVER = Font(color="C00000")
SAVE = Font(color="00804B", bold=True)


def signed(sheet, ref):
    sheet.conditional_formatting.add(ref, CellIsRule(operator="greaterThan", formula=["0.005"], font=OVER))
    sheet.conditional_formatting.add(ref, CellIsRule(operator="lessThan", formula=["-0.005"], font=SAVE))


for r in SIGNED:
    signed(wsum, f"B{r}")
signed(wp, f"{col_of['sav']}{DATA0}:{col_of['sav']}{LAST}")

wb.save(OUT)
print(f"{len(positions)} roles, {len(history)} history rows -> {OUT}")
