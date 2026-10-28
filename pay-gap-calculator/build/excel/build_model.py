# -*- coding: utf-8 -*-
"""
Builds the Excel model of the pay-gap calculator.

Reads a CSV of the same shape as the calculator and produces an .xlsx in which
the whole money chain is live: normalisation -> gaps -> regression -> selection
of recipients -> distribution of the uplift -> contributions -> totals.

Python supplies only the source data. Not one total is written in as a number:
anything that can be recomputed is recomputed by an Excel formula.

Correspondence with calc.js:
  normalise()        -> Data sheet, columns J..N
  regress()          -> Categories sheet, explicit sums + guard conditions
  computeGaps()      -> Categories sheet
  runScenario()      -> Data sheet (recipients) + Categories sheet (totals)
  contributionsOn()  -> Data sheet, contribution columns
  totalsBlock()      -> Summary sheet
"""
import csv
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName

# --- Defaults: DEFAULT_SETTINGS in calc.js ----------------------------------
DEFAULTS = {
    "threshold_pct": 5,
    "implementation_month": 7,
    "rate_below": 30,
    "rate_above": 15.1,
    "ceiling": 2979000,
}

# --- Styling ----------------------------------------------------------------
H_FILL = PatternFill("solid", fgColor="1F3B57")
H_FONT = Font(color="FFFFFF", bold=True, size=10)
IN_FILL = PatternFill("solid", fgColor="FFF4D6")    # source data
CALC_FILL = PatternFill("solid", fgColor="EAF3FA")  # computed by formula
SET_FILL = PatternFill("solid", fgColor="E6F4EA")   # a setting, meant to be changed
TITLE = Font(bold=True, size=13)
BOLD = Font(bold=True)
MUTED = Font(color="6B7280", size=9)
THIN = Side(style="thin", color="D0D7DE")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

EUR = '#,##0.00" ₽"'
PCT = '0.00"%"'
NUM = '#,##0.0000'


def read_csv(path):
    with open(path, newline="", encoding="utf-8-sig") as fh:
        rows = list(csv.DictReader(fh))
    out = []
    for i, r in enumerate(rows):
        out.append({
            "id": r.get("id") or "E%03d" % (i + 1),
            "category": r["category"].strip(),
            "gender": "F" if r["gender"].strip().upper() == "F" else "M",
            "base_salary": float(r["base_salary"]),
            "variable_pay": float(r["variable_pay"]),
            "tenure_years": float(r["tenure_years"]),
            "grade": float(r["grade"]),
            "fte": float(r["fte"]) if r.get("fte") else 1.0,
            "months_worked": float(r["months_worked"]) if r.get("months_worked") else 12.0,
        })
    return out


def header(ws, row, labels, widths=None):
    for j, text in enumerate(labels, start=1):
        c = ws.cell(row=row, column=j, value=text)
        c.fill = H_FILL
        c.font = H_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BOX
    if widths:
        for j, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(j)].width = w
    ws.row_dimensions[row].height = 30


def build(csv_path, out_path, title_note):
    emp = read_csv(csv_path)
    n = len(emp)

    # category order is order of first appearance, as in analyse()
    order, seen = [], set()
    for e in emp:
        if e["category"] not in seen:
            seen.add(e["category"])
            order.append(e["category"])

    wb = Workbook()

    # ======================================================================
    # SHEET Settings
    # ======================================================================
    st = wb.active
    st.title = "Настройки"
    st["A1"] = "Настройки модели"
    st["A1"].font = TITLE
    st["A2"] = title_note
    st["A2"].font = MUTED
    st["A4"] = "Менять можно только зелёные ячейки. Всё остальное пересчитывается само."
    st["A4"].font = MUTED

    rows_set = [
        ("threshold_pct", DEFAULTS["threshold_pct"], "Допустимый остаток, %",
         "Целевой необъяснённый разрыв для минимального сценария. Внутренний ориентир "
         "компании: закон РФ такого порога не задаёт (ст. 22 ТК РФ требует равной "
         "оплаты за труд равной ценности, ст. 132 запрещает дискриминацию в оплате)."),
        ("implementation_month", DEFAULTS["implementation_month"], "Месяц выравнивания",
         "1 = январь. Определяет затраты в текущем году: (13 - m) / 12."),
        ("rate_below", DEFAULTS["rate_below"], "Тариф взносов до предельной базы, %",
         "РФ, п. 3 ст. 425 НК РФ: единый тариф страховых взносов 30%. Применяется к части прибавки, "
         "которая укладывается в предельную базу. Взносы на травматизм (0,2–8,5%) не "
         "включены — добавьте свой тариф к обеим ставкам. У МСП с правом на льготу и IT-компаний тарифы ниже."),
        ("rate_above", DEFAULTS["rate_above"], "Тариф взносов сверх предельной базы, %",
         "РФ, п. 3 ст. 425 НК РФ: 15,1%. Применяется к части прибавки, которая выводит оклад за "
         "предельную базу."),
        ("ceiling", DEFAULTS["ceiling"], "Единая предельная база, ₽",
         "РФ 2026: 2 979 000 ₽ в год (Постановление Правительства РФ от 31.10.2025 "
         "№ 1705). Сравнивается с фактическим окладом сотрудника."),
    ]
    header(st, 6, ["Имя", "Значение", "Что это", "Пояснение"], [24, 14, 34, 76])
    for i, (name, val, label, note) in enumerate(rows_set):
        r = 7 + i
        st.cell(row=r, column=1, value=name).border = BOX
        c = st.cell(row=r, column=2, value=val)
        c.fill = SET_FILL
        c.border = BOX
        c.font = BOLD
        st.cell(row=r, column=3, value=label).border = BOX
        d = st.cell(row=r, column=4, value=note)
        d.border = BOX
        d.alignment = Alignment(wrap_text=True, vertical="top")
        wb.defined_names.add(DefinedName(name, attr_text="'Настройки'!$B$%d" % r))
        st.row_dimensions[r].height = 42

    st["A14"] = "Исходные данные"
    st["A14"].font = BOLD
    st["A15"] = ("На листе «Данные» — зафиксированная выгрузка. Генератор демо-данных "
                 "использует генератор псевдослучайных чисел (mulberry32 + Box-Muller), "
                 "который нельзя побайтно повторить формулами Excel, поэтому массив выгружен "
                 "числами. Всё, что посчитано поверх него, — живые формулы.")
    st["A15"].font = MUTED
    st["A15"].alignment = Alignment(wrap_text=True, vertical="top")
    st.merge_cells("A15:D17")

    # ======================================================================
    # SHEET Data
    # ======================================================================
    ws = wb.create_sheet("Данные")
    ws["A1"] = "Данные и расчёт по сотрудникам"
    ws["A1"].font = TITLE
    ws["A2"] = ("Жёлтое — исходные данные из CSV. Голубое — считается формулой. "
                "Формулы ссылаются на лист «Категории», где считаются разрывы по категориям.")
    ws["A2"].font = MUTED

    cols = [
        "id", "category", "gender",
        "base_salary", "variable_pay",
        "tenure_years", "grade", "fte", "months_worked",
        "factor", "n_base", "n_var", "n_total",
        "actual_total",
        "median_cat", "recip_min", "deficit_min", "share_min",
        "uplift_norm_min", "uplift_act_min", "contrib_min", "cost_min",
        "recip_full", "deficit_full", "share_full",
        "uplift_norm_full", "uplift_act_full", "contrib_full", "cost_full",
        "_rank_cat", "_n_cat", "_rank_sex", "_n_sex",
    ]
    HDR = 4
    header(ws, HDR, cols, [9, 18, 8, 13, 13, 11, 8, 7, 11,
                           9, 12, 12, 12, 13,
                           12, 11, 12, 10, 13, 13, 12, 12,
                           11, 12, 10, 13, 13, 12, 12,
                           10, 8, 10, 8])
    ws.freeze_panes = "A%d" % (HDR + 1)

    # units go in comments on the header cells rather than in the names
    # themselves: long labels in the header clutter the table
    notes = {
        1: "Табельный номер или любой уникальный код.",
        2: "Категория работников: группа должностей одинаковой ценности.",
        3: "Пол: F — женщина, M — мужчина.",
        4: "Оклад, ₽, фактически выплачено за период. Не приведено к году.",
        5: "Переменная часть, ₽, фактически выплачено за период — так же, как оклад, "
           "а не годовая сумма.",
        6: "Стаж в компании, лет.",
        7: "Место в лестнице грейдов: больше число — выше грейд. Для расчёта важен "
           "только порядок, а не само значение.",
        8: "Доля ставки: 1,0 — полная, 0,5 — половина.",
        9: "Отработано месяцев за период. 12 — полный год.",
        10: "factor = fte x (months_worked / 12). Приводит фактическую оплату к полной "
            "ставке и полному году.",
        11: "Оклад, ₽ в год на полную ставку: base_salary / factor.",
        12: "Переменная часть, ₽ в год на полную ставку: variable_pay / factor.",
        13: "Вся оплата, ₽ в год на полную ставку: n_base + n_var. По этой цифре "
            "сравниваются сотрудники.",
        14: "Фактически выплачено, ₽: base_salary + variable_pay. ФОТ — сумма этого столбца.",
        15: "Медиана n_total по всей категории, оба пола. ₽ в год на полную ставку.",
        16: "Минимальный сценарий: 1 — сотрудница получает прибавку.",
        17: "Минимальный сценарий: отставание от медианы категории, ₽ в год на полную ставку.",
        18: "Минимальный сценарий: доля сотрудницы в бюджете категории.",
        19: "Минимальный сценарий: прибавка, ₽ в год на полную ставку.",
        20: "Минимальный сценарий: прибавка, ₽ с учётом ставки и отработанных месяцев.",
        21: "Минимальный сценарий: страховые взносы с прибавки, ₽.",
        22: "Минимальный сценарий: прибавка + взносы, ₽.",
        23: "Полное выравнивание: 1 — сотрудница получает прибавку.",
        24: "Полное выравнивание: отставание от медианы категории, ₽.",
        25: "Полное выравнивание: доля сотрудницы в бюджете категории.",
        26: "Полное выравнивание: прибавка, ₽ в год на полную ставку.",
        27: "Полное выравнивание: прибавка, ₽ с учётом ставки и отработанных месяцев.",
        28: "Полное выравнивание: страховые взносы с прибавки, ₽.",
        29: "Полное выравнивание: прибавка + взносы, ₽.",
        30: "Служебный: ранг внутри категории по n_total — для медианы без формул массива.",
        31: "Служебный: численность категории.",
        32: "Служебный: ранг внутри пары «категория + пол».",
        33: "Служебный: численность пары «категория + пол».",
    }
    for col, text in notes.items():
        ws.cell(row=HDR, column=col).comment = Comment(text, "Модель")

    first, last = HDR + 1, HDR + n

    # category rows on the Categories sheet (its header is on row CHDR_CAT)
    CHDR_CAT = 4
    CFIRST = CHDR_CAT + 1
    CLAST = CFIRST + len(order) - 1

    def cat_lookup(col, row):
        """Value from column col of the Categories sheet for the category on row.

        Looked up by category name rather than by row number: SUMIFS returns a
        single match because category names are unique.
        """
        return ("SUMIFS('Категории'!${c}${f}:${c}${l},'Категории'!$A${f}:$A${l},B{r})"
                .format(c=col, f=CFIRST, l=CLAST, r=row))

    for i, e in enumerate(emp):
        r = first + i

        vals = [e["id"], e["category"], e["gender"], e["base_salary"],
                e["variable_pay"], e["tenure_years"], e["grade"],
                e["fte"], e["months_worked"]]
        for j, v in enumerate(vals, start=1):
            c = ws.cell(row=r, column=j, value=v)
            c.fill = IN_FILL
            c.border = BOX
            if j in (4, 5):
                c.number_format = EUR

        def put(col, formula, fmt=None):
            c = ws.cell(row=r, column=col, value=formula)
            c.fill = CALC_FILL
            c.border = BOX
            if fmt:
                c.number_format = fmt
            return c

        # --- normalisation: normalise() ---
        put(10, "=H{r}*(I{r}/12)".format(r=r), NUM)      # factor
        put(11, "=D{r}/J{r}".format(r=r), EUR)           # n_base
        put(12, "=E{r}/J{r}".format(r=r), EUR)           # n_var
        put(13, "=K{r}+L{r}".format(r=r), EUR)           # n_total
        put(14, "=D{r}+E{r}".format(r=r), EUR)           # actual_total

        # median of the category (both genders), without array formulas.
        # SUMPRODUCT computes the row's rank within its category by n_total, and
        # the median is assembled from one (odd n) or two (even n) middle values
        # via SUMIFS on that rank. Ordinary formulas: no Ctrl+Shift+Enter is
        # needed in any version of Excel.
        rank = ("SUMPRODUCT(($B${f}:$B${l}=B{r})*($M${f}:$M${l}<M{r}))+1"
                .format(f=first, l=last, r=r))
        put(15, ("=(SUMIFS($M${f}:$M${l},$B${f}:$B${l},B{r},$AD${f}:$AD${l},"
                 "INT((AE{r}+1)/2))+SUMIFS($M${f}:$M${l},$B${f}:$B${l},B{r},"
                 "$AD${f}:$AD${l},INT(AE{r}/2)+1))/2").format(
            f=first, l=last, r=r), EUR)

        # --- minimum scenario ---
        # Values from the Categories sheet are looked up by category name
        # (SUMIFS on column A) rather than by a reference to a specific row: the
        # category's row can move, its name cannot.
        put(16, '=IF(AND(C{r}="F",M{r}<O{r},{corr}=1),1,0)'.format(
            r=r, corr=cat_lookup("Z", r)))
        put(17, "=IF(P{r}=1,O{r}-M{r},0)".format(r=r), EUR)
        put(18, "=IF(AND(P{r}=1,{td}>0),Q{r}/{td},0)".format(
            r=r, td=cat_lookup("AA", r)), NUM)
        put(19, "=R{r}*{need}".format(r=r, need=cat_lookup("AB", r)), EUR)
        put(20, "=S{r}*J{r}".format(r=r), EUR)
        put(21, ("=IF(T{r}<=0,0,"
                 "(rate_below/100)*MIN(T{r},MAX(0,ceiling-D{r}))"
                 "+(rate_above/100)*(T{r}-MIN(T{r},MAX(0,ceiling-D{r}))))").format(r=r), EUR)
        put(22, "=T{r}+U{r}".format(r=r), EUR)

        # --- full scenario ---
        put(23, '=IF(AND(C{r}="F",M{r}<O{r},{corr}=1),1,0)'.format(
            r=r, corr=cat_lookup("AH", r)))
        put(24, "=IF(W{r}=1,O{r}-M{r},0)".format(r=r), EUR)
        put(25, "=IF(AND(W{r}=1,{td}>0),X{r}/{td},0)".format(
            r=r, td=cat_lookup("AI", r)), NUM)
        put(26, "=Y{r}*{need}".format(r=r, need=cat_lookup("AJ", r)), EUR)
        put(27, "=Z{r}*J{r}".format(r=r), EUR)
        put(28, ("=IF(AA{r}<=0,0,"
                 "(rate_below/100)*MIN(AA{r},MAX(0,ceiling-D{r}))"
                 "+(rate_above/100)*(AA{r}-MIN(AA{r},MAX(0,ceiling-D{r}))))").format(r=r), EUR)
        put(29, "=AA{r}+AB{r}".format(r=r), EUR)

        # --- helper columns for the non-array median ---
        # AD: the row's rank by n_total within its own category (1 = lowest).
        # The second term breaks ties by row number: with equal n_total the ranks
        # still differ, otherwise SUMIFS on rank would return the sum of two rows
        # instead of one and the median would come out twice too large.
        put(30, ("=SUMPRODUCT(($B${f}:$B${l}=B{r})*($M${f}:$M${l}<M{r}))"
                 "+SUMPRODUCT(($B${f}:$B${l}=B{r})*($M${f}:$M${l}=M{r})"
                 "*(ROW($M${f}:$M${l})<ROW(M{r})))+1"
                 ).format(f=first, l=last, r=r))
        # AE: size of the category
        put(31, "=COUNTIF($B${f}:$B${l},B{r})".format(f=first, l=last, r=r))
        # AF: rank within the "category + gender" pair — for medians by gender
        put(32, ("=SUMPRODUCT(($B${f}:$B${l}=B{r})*($C${f}:$C${l}=C{r})"
                 "*($M${f}:$M${l}<M{r}))"
                 "+SUMPRODUCT(($B${f}:$B${l}=B{r})*($C${f}:$C${l}=C{r})"
                 "*($M${f}:$M${l}=M{r})*(ROW($M${f}:$M${l})<ROW(M{r})))+1"
                 ).format(f=first, l=last, r=r))
        # AG: size of the "category + gender" group
        put(33, "=COUNTIFS($B${f}:$B${l},B{r},$C${f}:$C${l},C{r})".format(
            f=first, l=last, r=r))

    # ======================================================================
    # SHEET Categories
    # ======================================================================
    cs = wb.create_sheet("Категории")
    cs["A1"] = "Расчёт по категориям"
    cs["A1"].font = TITLE
    cs["A2"] = ("Регрессия: МНК n_total на грейд и стаж внутри каждой категории, решена через "
                "явные суммы (без ЛИНЕЙН и формул массива). Условия те же, что в калькуляторе: "
                "меньше 4 наблюдений или постоянный грейд либо стаж — регрессия не строится, "
                "весь разрыв считается необъяснённым.")
    cs["A2"].font = MUTED

    ccols = [
        "category", "headcount", "n_F", "n_M",                          # A-D
        "mean_M", "mean_F", "median_all",                               # E-G
        "raw_gap_mean_%", "median_M", "median_F", "raw_gap_median_%",   # H-K
        "var_M", "var_F", "variable_gap_%",                             # L-N
        "base_M", "base_F", "base_gap_%",                               # O-Q
        "var_grade", "var_tenure", "reg_ok",                            # R-T
        "coef_grade", "coef_tenure", "r2",                              # U-W
        "explained_%", "unexplained_%",                                 # X-Y
        "correct_min", "totalDef_min", "need_min",                      # Z-AB
        "adjust_min", "contrib_min", "cost_min", "recip_min",           # AC-AF
        "unexpl_after_min",                                             # AG
        "correct_full", "totalDef_full", "need_full",                   # AH-AJ
        "adjust_full", "contrib_full", "cost_full", "recip_full",       # AK-AN
        "unexpl_after_full",                                            # AO
        "unreliable", "reverse_gap",                                    # AP-AQ
    ] + [""] * 9 + [                                                    # AR-AZ empty
        "n", "Sum_g", "Sum_t", "Sum_y", "Sum_gg", "Sum_tt",             # BA-BF
        "Sum_gt", "Sum_gy", "Sum_ty", "Sum_yy", "det",                  # BG-BK
    ]
    CHDR = CHDR_CAT   # the same value: Data references this layout
    header(cs, CHDR, ccols, [18, 10, 7, 7] + [12] * (len(ccols) - 15) + [20] * 11)
    cs.freeze_panes = "B%d" % (CHDR + 1)

    D = "'Данные'!"

    def rng(col):
        return "%s$%s$%d:$%s$%d" % (D, col, first, col, last)

    B, G_ = rng("B"), rng("C")          # category, gender
    M_, K_, L_ = rng("M"), rng("K"), rng("L")   # n_total, n_base, n_var
    GR, TE = rng("G"), rng("F")         # grade, tenure

    for i, name in enumerate(order):
        r = CHDR + 1 + i
        cf = "$A%d" % r

        def put(col, formula, fmt=None, fill=CALC_FILL):
            c = cs.cell(row=r, column=col, value=formula)
            c.fill = fill
            c.border = BOX
            if fmt:
                c.number_format = fmt
            return c

        cat = cs.cell(row=r, column=1, value=name)
        cat.fill = IN_FILL
        cat.border = BOX

        put(2, "=COUNTIF({b},{c})".format(b=B, c=cf))
        put(3, '=COUNTIFS({b},{c},{g},"F")'.format(b=B, c=cf, g=G_))
        put(4, '=COUNTIFS({b},{c},{g},"M")'.format(b=B, c=cf, g=G_))

        put(5, '=IF(D{r}=0,0,AVERAGEIFS({m},{b},{c},{g},"M"))'.format(
            r=r, m=M_, b=B, c=cf, g=G_), EUR)
        put(6, '=IF(C{r}=0,0,AVERAGEIFS({m},{b},{c},{g},"F"))'.format(
            r=r, m=M_, b=B, c=cf, g=G_), EUR)
        # median of the whole category — taken ready-made from Data (already
        # computed there by a non-array formula, and identical on every row of
        # the category)
        put(7, "=IF(B{r}=0,0,SUMIFS({o},{b},{c},{ad},1))".format(
            r=r, o=rng("O"), b=B, c=cf, ad=rng("AD")), EUR)
        put(8, "=IF(OR(E{r}<=0,C{r}=0,D{r}=0),0,(E{r}-F{r})/E{r}*100)".format(r=r), PCT)

        # medians by gender — via the rank within the "category + gender" group
        # (Data!AF). With an even group size two middle values are taken; with an
        # odd one both halves of the formula point at the same row.
        def med_by_sex(sex, cnt_cell):
            return ('=IF({cnt}=0,0,(SUMIFS({m},{b},{c},{g},"{s}",{af},'
                    'INT(({cnt}+1)/2))+SUMIFS({m},{b},{c},{g},"{s}",{af},'
                    'INT({cnt}/2)+1))/2)').format(
                m=M_, b=B, c=cf, g=G_, s=sex, af=rng("AF"), cnt=cnt_cell)

        put(9, med_by_sex("M", "D%d" % r), EUR)
        put(10, med_by_sex("F", "C%d" % r), EUR)
        put(11, "=IF(OR(I{r}<=0,C{r}=0,D{r}=0),0,(I{r}-J{r})/I{r}*100)".format(r=r), PCT)

        put(12, '=IF(D{r}=0,0,AVERAGEIFS({l},{b},{c},{g},"M"))'.format(
            r=r, l=L_, b=B, c=cf, g=G_), EUR)
        put(13, '=IF(C{r}=0,0,AVERAGEIFS({l},{b},{c},{g},"F"))'.format(
            r=r, l=L_, b=B, c=cf, g=G_), EUR)
        put(14, "=IF(OR(L{r}<=0,C{r}=0,D{r}=0),0,(L{r}-M{r})/L{r}*100)".format(r=r), PCT)

        put(15, '=IF(D{r}=0,0,AVERAGEIFS({k},{b},{c},{g},"M"))'.format(
            r=r, k=K_, b=B, c=cf, g=G_), EUR)
        put(16, '=IF(C{r}=0,0,AVERAGEIFS({k},{b},{c},{g},"F"))'.format(
            r=r, k=K_, b=B, c=cf, g=G_), EUR)
        put(17, "=IF(OR(O{r}<=0,C{r}=0,D{r}=0),0,(O{r}-P{r})/O{r}*100)".format(r=r), PCT)

        # --- regression: the same guard as in regress() ---
        # predictor variance via Var = E[x^2] - (E[x])^2, without array formulas
        put(18, ("=IF(B{r}<2,0,SUMPRODUCT(({b}={c})*{g}*{g})/B{r}"
                 "-(SUMIFS({g},{b},{c})/B{r})^2)").format(
            r=r, b=B, c=cf, g=GR), NUM)
        put(19, ("=IF(B{r}<2,0,SUMPRODUCT(({b}={c})*{t}*{t})/B{r}"
                 "-(SUMIFS({t},{b},{c})/B{r})^2)").format(
            r=r, b=B, c=cf, t=TE), NUM)
        put(20, "=IF(AND(B{r}>=4,R{r}>0.000000001,S{r}>0.000000001),1,0)".format(r=r))

        # --- OLS without LINEST: 3x3 normal equations solved by Cramer's rule ---
        # First the sums (columns BA..BI below), then the centred moments:
        #   Sgg = Σg² - (Σg)²/n     Sgt = Σgt - Σg·Σt/n
        #   Stt = Σt² - (Σt)²/n     Sgy = Σgy - Σg·Σy/n     Sty = Σty - Σt·Σy/n
        # Determinant of the 2x2 system (after eliminating the intercept):
        #   det = Sgg·Stt - Sgt²
        #   coef_grade  = (Sgy·Stt - Sty·Sgt) / det
        #   coef_tenure = (Sty·Sgg - Sgy·Sgt) / det
        # This is exactly what solveGauss() computes by Gaussian elimination, but
        # written out explicitly: at 3x3 Cramer's rule is shorter and fits
        # visibly inside a single cell.
        # helper sums, strictly in columns BA(53)..BJ(62)
        sums = [
            ("BA", "=B{r}".format(r=r)),                                        # n
            ("BB", "=SUMIFS({g},{b},{c})".format(g=GR, b=B, c=cf)),             # Σg
            ("BC", "=SUMIFS({t},{b},{c})".format(t=TE, b=B, c=cf)),             # Σt
            ("BD", "=SUMIFS({m},{b},{c})".format(m=M_, b=B, c=cf)),             # Σy
            ("BE", "=SUMPRODUCT(({b}={c})*{g}*{g})".format(b=B, c=cf, g=GR)),   # Σgg
            ("BF", "=SUMPRODUCT(({b}={c})*{t}*{t})".format(b=B, c=cf, t=TE)),   # Σtt
            ("BG", "=SUMPRODUCT(({b}={c})*{g}*{t})".format(b=B, c=cf, g=GR, t=TE)),
            ("BH", "=SUMPRODUCT(({b}={c})*{g}*{m})".format(b=B, c=cf, g=GR, m=M_)),
            ("BI", "=SUMPRODUCT(({b}={c})*{t}*{m})".format(b=B, c=cf, t=TE, m=M_)),
            ("BJ", "=SUMPRODUCT(({b}={c})*{m}*{m})".format(b=B, c=cf, m=M_)),   # Σyy
        ]
        for k, (_col, formula) in enumerate(sums):
            put(53 + k, formula, NUM)
        Sgg = "(BE{r}-BB{r}^2/B{r})".format(r=r)
        Stt = "(BF{r}-BC{r}^2/B{r})".format(r=r)
        Sgt = "(BG{r}-BB{r}*BC{r}/B{r})".format(r=r)
        Sgy = "(BH{r}-BB{r}*BD{r}/B{r})".format(r=r)
        Sty = "(BI{r}-BC{r}*BD{r}/B{r})".format(r=r)
        Syy = "(BJ{r}-BD{r}^2/B{r})".format(r=r)
        det = "({gg}*{tt}-{gt}^2)".format(gg=Sgg, tt=Stt, gt=Sgt)

        # BK: the determinant, kept separate so the guard can check it
        put(63, "=IF(T{r}=0,0,{det})".format(r=r, det=det), NUM)
        put(21, "=IF(OR(T{r}=0,ABS(BK{r})<0.000000001),0,({gy}*{tt}-{ty}*{gt})/BK{r})".format(
            r=r, gy=Sgy, tt=Stt, ty=Sty, gt=Sgt), EUR)
        put(22, "=IF(OR(T{r}=0,ABS(BK{r})<0.000000001),0,({ty}*{gg}-{gy}*{gt})/BK{r})".format(
            r=r, ty=Sty, gg=Sgg, gy=Sgy, gt=Sgt), EUR)
        # R2 = (coef_grade*Sgy + coef_tenure*Sty) / Syy — share of variance explained
        put(23, ('=IF(OR(T{r}=0,ABS(BK{r})<0.000000001,{yy}<0.000000001),"",'
                 '(U{r}*{gy}+V{r}*{ty})/{yy})').format(
            r=r, yy=Syy, gy=Sgy, ty=Sty), NUM)

        # explained: coef_grade * dGrade + coef_tenure * dTenure, as % of mean_M
        dg = ('(IF(D{r}=0,0,AVERAGEIFS({g},{b},{c},{gd},"M"))'
              '-IF(C{r}=0,0,AVERAGEIFS({g},{b},{c},{gd},"F")))').format(
            r=r, g=GR, b=B, c=cf, gd=G_)
        dt = ('(IF(D{r}=0,0,AVERAGEIFS({t},{b},{c},{gd},"M"))'
              '-IF(C{r}=0,0,AVERAGEIFS({t},{b},{c},{gd},"F")))').format(
            r=r, t=TE, b=B, c=cf, gd=G_)
        raw_expl = "IF(OR(T{r}=0,E{r}<=0),0,(U{r}*{dg}+V{r}*{dt})/E{r}*100)".format(
            r=r, dg=dg, dt=dt)
        put(24, "=IF(H{r}>=0,MAX(0,MIN({e},H{r})),MIN(0,MAX({e},H{r})))".format(
            r=r, e=raw_expl), PCT)
        put(25, "=H{r}-X{r}".format(r=r), PCT)

        # --- minimum scenario ---
        put(26, "=IF(AND(H{r}>=0,Y{r}>threshold_pct),1,0)".format(r=r))
        put(27, "=IF(Z{r}=0,0,SUMIFS({q},{b},{c}))".format(
            r=r, q="%s$Q$%d:$Q$%d" % (D, first, last), b=B, c=cf), EUR)
        put(28, "=IF(Z{r}=0,0,(Y{r}-threshold_pct)/100*E{r}*C{r})".format(r=r), EUR)
        put(29, "=SUMIFS({t},{b},{c})".format(
            r=r, t="%s$T$%d:$T$%d" % (D, first, last), b=B, c=cf), EUR)
        put(30, "=SUMIFS({u},{b},{c})".format(
            r=r, u="%s$U$%d:$U$%d" % (D, first, last), b=B, c=cf), EUR)
        put(31, "=AC{r}+AD{r}".format(r=r), EUR)
        put(32, "=SUMIFS({p},{b},{c})".format(
            r=r, p="%s$P$%d:$P$%d" % (D, first, last), b=B, c=cf))
        new_f = ('(IF(C{r}=0,0,AVERAGEIFS({m},{b},{c},{g},"F"))'
                 '+IF(C{r}=0,0,SUMIFS({s},{b},{c},{g},"F")/C{r}))').format(
            r=r, m=M_, b=B, c=cf, g=G_, s="%s$S$%d:$S$%d" % (D, first, last))
        put(33, "=IF(E{r}<=0,0,(E{r}-{nf})/E{r}*100-X{r})".format(r=r, nf=new_f), PCT)

        # --- full scenario: target 0 ---
        put(34, "=IF(AND(H{r}>=0,Y{r}>0),1,0)".format(r=r))
        put(35, "=IF(AH{r}=0,0,SUMIFS({x},{b},{c}))".format(
            r=r, x="%s$X$%d:$X$%d" % (D, first, last), b=B, c=cf), EUR)
        put(36, "=IF(AH{r}=0,0,Y{r}/100*E{r}*C{r})".format(r=r), EUR)
        put(37, "=SUMIFS({aa},{b},{c})".format(
            r=r, aa="%s$AA$%d:$AA$%d" % (D, first, last), b=B, c=cf), EUR)
        put(38, "=SUMIFS({ab},{b},{c})".format(
            r=r, ab="%s$AB$%d:$AB$%d" % (D, first, last), b=B, c=cf), EUR)
        put(39, "=AK{r}+AL{r}".format(r=r), EUR)
        put(40, "=SUMIFS({w},{b},{c})".format(
            r=r, w="%s$W$%d:$W$%d" % (D, first, last), b=B, c=cf))
        new_f2 = ('(IF(C{r}=0,0,AVERAGEIFS({m},{b},{c},{g},"F"))'
                  '+IF(C{r}=0,0,SUMIFS({z},{b},{c},{g},"F")/C{r}))').format(
            r=r, m=M_, b=B, c=cf, g=G_, z="%s$Z$%d:$Z$%d" % (D, first, last))
        put(41, "=IF(E{r}<=0,0,(E{r}-{nf})/E{r}*100-X{r})".format(r=r, nf=new_f2), PCT)

        put(42, "=IF(OR(B{r}<10,C{r}<3,D{r}<3),1,0)".format(r=r))
        put(43, "=IF(H{r}<0,1,0)".format(r=r))

    clast = CHDR + len(order)

    # ======================================================================
    # SHEET Summary
    # ======================================================================
    sm = wb.create_sheet("Итоги")
    sm["A1"] = "Итоги"
    sm["A1"].font = TITLE
    sm["A2"] = ("Каждая цифра здесь — формула со ссылкой на листы «Данные» и «Категории». "
                "Ни одно число на этом листе не введено вручную.")
    sm["A2"].font = MUTED
    sm.column_dimensions["A"].width = 44
    for col in "BCD":
        sm.column_dimensions[col].width = 22

    sm["A4"] = "ФОТ до выравнивания (оклад и переменная часть, без взносов)"
    sm["A4"].font = BOLD
    sm["B4"] = "=SUM('Данные'!$N${f}:$N${l})".format(f=first, l=last)
    sm["B4"].number_format = EUR
    sm["B4"].fill = CALC_FILL

    sm["A5"] = "Численность"
    sm["B5"] = "=COUNTA('Данные'!$A${f}:$A${l})".format(f=first, l=last)
    sm["B5"].fill = CALC_FILL

    SHDR = 7
    header(sm, SHDR, ["Показатель", "Минимальный сценарий", "Полное выравнивание", "Разница"])
    # row numbers are derived from SHDR rather than written in as constants:
    # they once drifted from the actual layout and "Total" ended up referring to
    # itself, producing a circular reference and an empty cell
    R_ADJ = SHDR + 1        # payroll adjustment
    R_CON = SHDR + 2        # contributions
    R_TOT = SHDR + 3        # total
    TOTAL_LABEL = "Всего затрат"
    lines = [
        ("Прибавка к окладам", "=SUM('Категории'!$AC${f}:$AC${l})",
         "=SUM('Категории'!$AK${f}:$AK${l})", EUR),
        ("Страховые взносы", "=SUM('Категории'!$AD${f}:$AD${l})",
         "=SUM('Категории'!$AL${f}:$AL${l})", EUR),
        (TOTAL_LABEL, "=B{a}+B{c}".format(a=R_ADJ, c=R_CON),
         "=C{a}+C{c}".format(a=R_ADJ, c=R_CON), EUR),
        ("Затраты к ФОТ, %", "=IF($B$4=0,0,B{t}/$B$4*100)".format(t=R_TOT),
         "=IF($B$4=0,0,C{t}/$B$4*100)".format(t=R_TOT), PCT),
        ("Получают прибавку, чел.", "=SUM('Категории'!$AF${f}:$AF${l})",
         "=SUM('Категории'!$AN${f}:$AN${l})", '#,##0'),
        ("Затраты в этом году",
         "=B{t}*(13-implementation_month)/12".format(t=R_TOT),
         "=C{t}*(13-implementation_month)/12".format(t=R_TOT), EUR),
        ("Затраты в следующем году, полный год", "=B{t}".format(t=R_TOT),
         "=C{t}".format(t=R_TOT), EUR),
    ]
    for i, (label, fmin, ffull, fmt) in enumerate(lines):
        r = SHDR + 1 + i
        c = sm.cell(row=r, column=1, value=label)
        c.border = BOX
        if label == TOTAL_LABEL:
            c.font = BOLD
        for col, f in ((2, fmin), (3, ffull)):
            cc = sm.cell(row=r, column=col, value=f.format(f=CHDR + 1, l=clast))
            cc.number_format = fmt
            cc.fill = CALC_FILL
            cc.border = BOX
            if label == TOTAL_LABEL:
                cc.font = BOLD
        d = sm.cell(row=r, column=4, value="=C{r}-B{r}".format(r=r))
        d.number_format = fmt
        d.fill = CALC_FILL
        d.border = BOX
        if label == TOTAL_LABEL:
            d.font = BOLD

    # rows below the table are placed from its actual end, not by constants
    r_note = SHDR + len(lines) + 2
    sm.cell(row=r_note, column=1, value="Сверка с калькулятором").font = BOLD
    note_cell = sm.cell(row=r_note + 1, column=1, value=(
        "Итоги калькулятора на демо-данных совпадают с B{t} / C{t} до копейки. "
        "Если загрузить в калькулятор свой CSV и вставить те же строки на лист «Данные», "
        "итоги тоже должны совпасть.".format(t=R_TOT)))
    note_cell.font = MUTED
    note_cell.alignment = Alignment(wrap_text=True, vertical="top")
    sm.merge_cells(start_row=r_note + 1, start_column=1,
                   end_row=r_note + 2, end_column=4)

    # ======================================================================
    # SHEET Legend
    # ======================================================================
    lg = wb.create_sheet("Пояснения")
    lg["A1"] = "Пояснения и методика"
    lg["A1"].font = TITLE
    lg.column_dimensions["A"].width = 30
    lg.column_dimensions["B"].width = 110

    blocks = [
        ("Цвета", ""),
        ("Жёлтый", "Исходные данные. Берутся из CSV, не считаются."),
        ("Голубой", "Считается формулой. Править нечего — пересчитывается само."),
        ("Зелёный", "Настройка на листе «Настройки». Эти ячейки можно менять."),
        ("", ""),
        ("Единицы", ""),
        ("Где смотреть",
         "Единицы каждого столбца — в примечании к его заголовку на листе «Данные» "
         "(красный уголок — наведите курсор). Коротко то же самое — ниже."),
        ("base_salary, variable_pay",
         "Рубли, фактически выплаченные за период, без приведения к году. У сотрудника на "
         "0,5 ставки или отработавшего 7 месяцев стоит то, что он действительно получил. "
         "Переменная часть записана так же, как оклад."),
        ("n_base, n_var, n_total",
         "Рубли в год на полную ставку: фактическая оплата, делённая на factor. Сравнивать "
         "сотрудников можно только по этим цифрам, иначе половина разрыва окажется разницей "
         "в отработанном времени, а не в оплате."),
        ("grade",
         "Место в лестнице грейдов: больше число — выше грейд. Для расчёта важен только "
         "порядок, само значение смысла не несёт."),
        ("", ""),
        ("Порядок расчёта", ""),
        ("1. Приведение",
         "factor = fte x (months_worked / 12). Оклад и переменная часть делятся на factor и "
         "приводятся к полной ставке и полному году. Сравнимы между сотрудниками только "
         "приведённые цифры."),
        ("2. Разрывы",
         "raw_gap = (mean_M - mean_F) / mean_M x 100 — по n_total (приведённые оклад и "
         "переменная часть), отдельно в каждой категории."),
        ("3. Регрессия",
         "МНК n_total на грейд и стаж внутри каждой категории, решена через явные суммы "
         "методом Крамера. Отвечает на один вопрос: какую часть разрыва объясняют грейд и "
         "стаж. Условие то же, что в калькуляторе: меньше 4 наблюдений или постоянный грейд "
         "либо стаж — регрессия не строится, explained = 0, весь разрыв считается "
         "необъяснённым (осторожная сторона)."),
        ("4. Объяснённая доля",
         "explained = coef_grade x (средний грейд мужчин - средний грейд женщин) + "
         "coef_tenure x (средний стаж мужчин - средний стаж женщин), в % от mean_M. Не больше "
         "самого разрыва: модель не может объяснить больше, чем есть."),
        ("5. Кто получает прибавку",
         "Женщины, у которых приведённый n_total ниже медианы своей категории (медиана по "
         "всей категории, оба пола). Категории, где разрыв в пользу женщин, не выравниваются "
         "ни в одном сценарии — выравнивание только поднимает оплату."),
        ("6. Распределение",
         "Пропорционально отставанию от медианы: доля = (медиана - n_total) / сумма всех "
         "отставаний. Чем ниже середины сотрудница, тем большую часть бюджета она получает."),
        ("Две разные величины",
         "Средний разрыв категории и отставание сотрудницы решают разные задачи. Средний "
         "разрыв определяет, СКОЛЬКО ДЕНЕГ нужно категории (шаг 5, формула потребности). "
         "Отставание от медианы определяет, КАК ЭТИ ДЕНЬГИ ДЕЛЯТСЯ между людьми (шаг 6). "
         "В индивидуальную прибавку средний разрыв не входит: отставание меряется от медианы "
         "категории, а не от средней оплаты мужчин."),
        ("Следствие",
         "Сотрудница с оплатой выше медианы ничего не получает, даже если разрыв в её "
         "категории большой. Метод поднимает нижнюю половину, а не всех женщин подряд."),
        ("7. Страховые взносы",
         "Начисляются на прибавку к окладу, а не на всю зарплату, и всегда по фактическому "
         "окладу. Часть прибавки, которая укладывается в единую предельную базу, облагается "
         "по полному тарифу (РФ 2026: 30%), часть сверх неё — по пониженному (15,1%). "
         "Взносы на травматизм не включены: добавьте свой тариф к обеим ставкам на листе "
         "«Настройки»."),
        ("8. Итоги",
         "Затраты в этом году = полная сумма x (13 - месяц выравнивания) / 12. Январь даёт "
         "12/12, декабрь — 1/12."),
        ("", ""),
        ("Два сценария", ""),
        ("Минимальный сценарий",
         "Довести необъяснённый разрыв до допустимого остатка (по умолчанию 5%). Категории, "
         "где разрыв уже ниже, не трогаются. Допустимый остаток — внутренний ориентир "
         "компании: закон РФ такого порога не задаёт."),
        ("Полное выравнивание",
         "Довести необъяснённый разрыв до нуля в каждой категории, где он положительный."),
        ("", ""),
        ("Оговорки", ""),
        ("Что значит «объяснённый»",
         "Слово техническое, а не оправдательное. Если женщин системно не продвигают по "
         "грейдам, регрессия спишет разрыв на грейд и назовёт его объяснённым. Это "
         "ограничение метода, а не вывод о справедливости."),
        ("Без формул массива",
         "Вся модель — на обычных формулах: СУММЕСЛИМН, СУММПРОИЗВ, СЧЁТЕСЛИМН. "
         "Ctrl+Shift+Enter не нужен ни в одной версии Excel. Медианы считаются через ранг "
         "(служебные столбцы _rank_cat / _rank_sex на листе «Данные»), регрессия — через "
         "явные суммы, а не ЛИНЕЙН."),
        ("Генератор данных",
         "Демо-данные созданы генератором mulberry32 + Box-Muller с фиксированным зерном. "
         "Формулами Excel его побайтно не повторить, поэтому массив выгружен числами. Всё, "
         "что посчитано поверх него, — живые формулы. Данные смоделированы, реальных "
         "сотрудников за ними нет."),
    ]
    r = 3
    for label, text in blocks:
        a = lg.cell(row=r, column=1, value=label)
        b = lg.cell(row=r, column=2, value=text)
        if text == "":
            a.font = BOLD
        else:
            a.font = Font(bold=True, size=10)
            b.alignment = Alignment(wrap_text=True, vertical="top")
            lg.row_dimensions[r].height = max(15, 13 * (len(text) // 105 + 1))
        r += 1

    wb.save(out_path)
    return n, len(order)


if __name__ == "__main__":
    src = sys.argv[1] if len(sys.argv) > 1 else "build/excel/demo-data.csv"
    dst = sys.argv[2] if len(sys.argv) > 2 else "pay-gap-model.xlsx"
    note = sys.argv[3] if len(sys.argv) > 3 else "Демо-данные, 212 сотрудников (смоделированы)"
    rows, cats = build(src, dst, note)
    print("OK: %s  (%d rows, %d categories)" % (dst, rows, cats))
