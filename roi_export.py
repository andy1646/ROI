"""Write the model out as a live Excel workbook.

Every derived cell is a formula, not a number, so the exported file is a
working model rather than a snapshot: change a yellow input and Excel
recalculates the whole sheet, ROI included. Each scenario gets its own sheet,
and a Comparison sheet cross-references them so it updates too.

Layout of a scenario sheet (column D is Year 0, J is Year 6):

    row  4  Gross Revenue      D = input, E..J = prior year * (1 + growth)
    row  5  Revenue Growth     input, years 1..6
    row  7  Commission         = gross revenue * commission %
    row  8  Commission %       input in E, E..J locked to it
    row 11  Net Revenue        = gross revenue - commission
    row 13  COGS               = SUM of the eight category lines below
    row 14-21  COGS by category = gross revenue * sales mix % * that
                                 category's COGS rate (rows 70-77)
    row 22  COGS % of net rev   = COGS / net revenue, shown for comparison
    row 25  Payroll            Y1 = net revenue * %, then grown each year
    row 26  Payroll % / growth  input: Y1 is a share, Y2..Y6 are growth rates
    row 29  Expenses           = net revenue * expenses %
    row 30  Expenses %         input, per year
    row 32  NOI                = net revenue - COGS - payroll - expenses
    row 33  NOI Margin         = NOI / net revenue
    row 36  Capex Investment   Y0 = gross revenue Y0 * %, then revenue * %
    row 37  Capex %            input
    row 40  Net Cash           = NOI - capex
    row 41  Cumulative Net Cash
    row 43  ROI                = RATE(term, , NetCash_Y0, SUM(NetCash_Y0:Y6))
    row 70  COGS Rates         input, one per income category
    row 107 Contract Term      input, and it may be fractional
    row 108 Run Rate           what a year bills in full; growth compounds here
    row 109 Year Fraction      = MEDIAN(0, term - (n-1), 1)

A term of 6.25 years runs six whole years and a three-month stub. The stub
gets a column of its own and is prorated, not counted as a year: row 90
carries the annualised run rate that growth compounds on, row 91 the share
of each year that falls inside the term, and row 4 multiplies the two. Every
line below row 4 keys off row 4, so they all prorate with it. The roster
(rows 80+) bills months directly, so it reads row 91 into its month span,
and the bonus write-off divides by the fractional term rather than a year
count - which keeps it amortising to exactly the year-0 bonus.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from roi_core import (CATEGORIES, COGS_CATEGORIES, DEFAULT_TERM, MAX_YEARS,
                      STAFF_TYPES, as_years, cogs_rates_of, role_cost,
                      year_count)

# Column D holds Year 0; a sheet uses as many columns as its term needs.
COL = [get_column_letter(4 + t) for t in range(MAX_YEARS + 1)]

R = dict(title=2, year=3, rev=4, growth=5, comm=7, comm_pct=8, net=11,
         cogs=13, cogs0=14, cogs_pct=22, pay=25, pay_pct=26,
         exp=29, exp_pct=30,
         amort=31, noi=32, margin=33, capex=36, capex_pct=37,
         cash=40, cum=41,
         roi=43, total=44, payback=45, avgmargin=46,
         cat_head=48, rate0=49, mix_head=58, mix0=59,
         cogs_head=69, cogs_rate0=70,
         staff_head=79, staff0=80, merit=95, incentive=96,
         profit=97, pay0=102, exp0=103, exp0_pct=104, bonus=105,
         invest=106,
         term_row=107, rate=108, frac=109,
         legend=112)

MONEY = '_(* #,##0_);_(* (#,##0);_(* "-"??_);_(@_)'
PCT = "0.00%"

YELLOW = PatternFill("solid", start_color="FFFFFF00")
GREEN = PatternFill("solid", start_color="FFC6EFCE")
BLUE = PatternFill("solid", start_color="FFDEEBF7")   # the COGS block
HEAD_FILL = PatternFill("solid", start_color="FFF2F5F8")

INK = "FF1D2430"
MUTED = "FF6B7684"
GREEN_INK = "FF006100"

THIN = Side(style="thin", color="FFD8DEE6")
TOP_RULE = Border(top=THIN)


def _note_cell(ws, ref, text):
    """Write the user's note as text.

    A note beginning =, +, - or @ would otherwise be stored as a formula and
    break the sheet, so those are forced back to a plain string.
    """
    c = ws[ref]
    c.value = text
    if text[:1] in "=+-@":
        c.data_type = "s"
    c.font = Font(name="Calibri", size=11, italic=True, color=INK)
    c.alignment = Alignment(horizontal="left")
    return c


def _label(ws, row, text, bold=False, indent=0):
    c = ws.cell(row=row, column=3, value=text)
    c.font = Font(name="Calibri", size=11, bold=bold, color=INK)
    c.alignment = Alignment(horizontal="left", indent=indent)
    return c


def _put(ws, row, col_letter, value, fmt=MONEY, fill=None, bold=False,
         colour=INK):
    c = ws[f"{col_letter}{row}"]
    c.value = value
    c.number_format = fmt
    c.font = Font(name="Calibri", size=11, bold=bold, color=colour)
    if fill:
        c.fill = fill
    return c


def _sheet(wb, name, p, note=""):
    """One scenario, written entirely as formulas over its input cells."""
    term = float(p.get("term", DEFAULT_TERM))
    # A fractional term still needs a whole column for its stub year.
    n = year_count(term)
    cols = COL[:n + 1]
    last = cols[-1]
    growth = as_years(p["rev_growth"], n)
    cogs_pct = as_years(p.get("cogs", 0.0), n)
    cogs_rates = cogs_rates_of(p)
    exp_pct = as_years(p["expenses"], n)
    capex_pct = as_years(p["capex"], n)
    n_staff = len(STAFF_TYPES)
    titles = p.get("staff_titles") or list(STAFF_TYPES)
    starts = p.get("staff_starts") or [0] * n_staff
    bonus = p.get("signing_bonus", 0.0)
    if not isinstance(bonus, (list, tuple)):
        bonus = [float(bonus)] + [0.0] * n
    bonus = (list(bonus) + [0.0] * (n + 1))[:n + 1]
    first, last_s = R["staff0"], R["staff0"] + n_staff - 1

    def ycol(t):
        """Roster year columns start at H, one per operating year."""
        return get_column_letter(7 + t)
    ws = wb.create_sheet(name)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 2
    ws.column_dimensions["C"].width = 26
    for c in cols:
        ws.column_dimensions[c].width = 15

    t = ws.cell(row=R["title"], column=3, value=f"{name} - Hotel ROI Modeling Tool")
    t.font = Font(name="Calibri", size=14, bold=True, color=INK)
    if note:
        _note_cell(ws, f"{COL[0]}{R['title']}", note)

    # ---- year header ----
    for t_i, c in enumerate(cols):
        cell = _put(ws, R["year"], c, t_i, "0", HEAD_FILL, bold=True,
                    colour=MUTED)
        cell.alignment = Alignment(horizontal="right")
    hdr = ws.cell(row=R["year"], column=3, value="Year")
    hdr.font = Font(name="Calibri", size=11, bold=True, color=MUTED)
    hdr.fill = HEAD_FILL

    # ---- gross revenue and its growth ----
    # Booked revenue is the run rate (row 90) times the share of the year
    # that falls inside the term (row 91), so a stub year bills only its
    # months while the escalator still steps a whole notch each anniversary.
    _label(ws, R["rev"], "Gross Revenue", bold=True)
    _put(ws, R["rev"], cols[0], p["gross_revenue"], MONEY, YELLOW)
    for t_i in range(1, n + 1):
        c = cols[t_i]
        _put(ws, R["rev"], c, f"={c}{R['rate']}*{c}{R['frac']}")

    _label(ws, R["growth"], "Revenue Growth", indent=1)
    for t_i in range(1, n + 1):
        _put(ws, R["growth"], cols[t_i], growth[t_i - 1] / 100.0, PCT, YELLOW)

    # ---- commission: one input, the later years locked to it ----
    _label(ws, R["comm"], "Commission")
    _label(ws, R["comm_pct"], "Commission %", indent=1)
    # Blended from the category block further down, as the workbook does.
    rate_span = f"$D${R['rate0']}:$D${R['rate0'] + len(CATEGORIES) - 1}"
    mix_span = f"$D${R['mix0']}:$D${R['mix0'] + len(CATEGORIES) - 1}"
    _put(ws, R["comm_pct"], cols[1], f"=SUMPRODUCT({rate_span},{mix_span})", PCT)
    for t_i in range(1, n + 1):
        _put(ws, R["comm"], cols[t_i],
             f"={cols[t_i]}{R['rev']}*{cols[t_i]}{R['comm_pct']}")
        if t_i > 1:
            _put(ws, R["comm_pct"], cols[t_i], f"=${cols[1]}${R['comm_pct']}", PCT)

    # ---- net revenue ----
    _label(ws, R["net"], "Net Revenue", bold=True)
    for t_i in range(1, n + 1):
        c = cols[t_i]
        _put(ws, R["net"], c, f"={c}{R['rev']}-{c}{R['comm']}", bold=True)
        ws[f"{c}{R['net']}"].border = TOP_RULE
    ws[f"C{R['net']}"].border = TOP_RULE

    # ---- COGS, costed category by category ----
    # Each line is that category's own revenue - gross revenue times its
    # share of sales - at its own rate, and COGS is the sum of them. The
    # rates live in the input block further down, one per category.
    _label(ws, R["cogs"], "COGS")
    if cogs_rates:
        first_line = R["cogs0"]
        last_line = R["cogs0"] + len(COGS_CATEGORIES) - 1
        for i, cat in enumerate(COGS_CATEGORIES):
            row = R["cogs0"] + i
            _label(ws, row, cat, indent=1).fill = BLUE
            for t_i in range(1, n + 1):
                c = cols[t_i]
                _put(ws, row, c,
                     f"={c}{R['rev']}*$D${R['mix0'] + i}"
                     f"*$D${R['cogs_rate0'] + i}", fill=BLUE)
        for t_i in range(1, n + 1):
            c = cols[t_i]
            _put(ws, R["cogs"], c, f"=SUM({c}{first_line}:{c}{last_line})")
        # What the category rates come to against net revenue, so the sheet
        # still says what COGS is costing in the terms the old one used.
        _label(ws, R["cogs_pct"], "COGS % of net revenue", indent=1)
        for t_i in range(1, n + 1):
            c = cols[t_i]
            _put(ws, R["cogs_pct"], c, f"={c}{R['cogs']}/{c}{R['net']}", PCT,
                 colour=MUTED)
    else:
        # No category rates: one input rate per year, against net revenue.
        _label(ws, R["cogs_pct"], "COGS %", indent=1)
        for t_i in range(1, n + 1):
            c = cols[t_i]
            _put(ws, R["cogs"], c, f"={c}{R['net']}*{c}{R['cogs_pct']}")
            _put(ws, R["cogs_pct"], c, cogs_pct[t_i - 1] / 100.0, PCT, YELLOW)

    # ---- expenses, a flat share of net revenue ----
    _label(ws, R["exp"], "Expenses")
    _label(ws, R["exp_pct"], "Expenses %", indent=1)
    for t_i in range(1, n + 1):
        c = cols[t_i]
        _put(ws, R["exp"], c, f"={c}{R['net']}*{c}{R['exp_pct']}")
        _put(ws, R["exp_pct"], c, exp_pct[t_i - 1] / 100.0, PCT, YELLOW)

    # ---- payroll: a share of year-1 net revenue, then grown ----
    _label(ws, R["pay"], "Payroll")
    _label(ws, R["pay_pct"], "Payroll % of net revenue", indent=1)
    # Roster escalated by merit, plus the revenue-linked overlays. The
    # overlays key off booked revenue, so they prorate with it.
    for t_i in range(1, n + 1):
        c = cols[t_i]
        yc = ycol(t_i)
        _put(ws, R["pay"], c,
             f"=SUM({yc}{first}:{yc}{last_s})"
             f"+{c}{R['rev']}*$D${R['incentive']}"
             f"+{c}{R['net']}*$D${R['profit']}")
        _put(ws, R["pay_pct"], c, f"={c}{R['pay']}/{c}{R['net']}", PCT,
             colour=MUTED)

    # ---- the year-0 bonus written off over the term, plus this year's own;
    # it lowers NOI but is added back to cash, so it is not a cash cost ----
    _label(ws, R["amort"], "Bonus Amortization")
    for t_i in range(1, n + 1):
        c = cols[t_i]
        # Divided by the fractional term and weighted by the year, so the
        # write-off still totals exactly the year-0 bonus. The recurring
        # bonus is an annual cost, so a stub year pays its share of it.
        _put(ws, R["amort"], c,
             f"=$D${R['bonus']}/$D${R['term_row']}*{c}{R['frac']}"
             f"+{c}{R['bonus']}*{c}{R['frac']}", colour=MUTED)

    # ---- NOI ----
    _label(ws, R["noi"], "NOI", bold=True)
    _label(ws, R["margin"], "NOI Margin", indent=1)
    for t_i in range(1, n + 1):
        c = cols[t_i]
        _put(ws, R["noi"], c,
             f"={c}{R['net']}-{c}{R['cogs']}-{c}{R['pay']}-{c}{R['exp']}"
             f"-{c}{R['amort']}", bold=True)
        ws[f"{c}{R['noi']}"].border = TOP_RULE
        _put(ws, R["margin"], c, f"={c}{R['noi']}/{c}{R['net']}", PCT,
             colour=MUTED)
    ws[f"C{R['noi']}"].border = TOP_RULE

    # ---- capex ----
    _label(ws, R["capex"], "Capex Investment")
    _label(ws, R["capex_pct"], "Capex %", indent=1)
    # Year 0 keys off year-1 gross revenue, as the workbook does (=+E4*D52).
    # That is a full year of billing, so it reads the run rate.
    _put(ws, R["capex"], cols[0], f"={cols[1]}{R['rate']}*{cols[0]}{R['capex_pct']}")
    _put(ws, R["capex_pct"], cols[0], p["capex_initial"] / 100.0, PCT, YELLOW)
    for t_i in range(1, n + 1):
        c = cols[t_i]
        _put(ws, R["capex"], c, f"={c}{R['rev']}*{c}{R['capex_pct']}")
        _put(ws, R["capex_pct"], c, capex_pct[t_i - 1] / 100.0, PCT, YELLOW)

    # ---- net cash and its running total ----
    _label(ws, R["cash"], "Net Cash", bold=True)
    _label(ws, R["cum"], "Cumulative Net Cash", indent=1)
    _put(ws, R["cash"], cols[0], f"=-{cols[0]}{R['invest']}", bold=True)
    _put(ws, R["cum"], cols[0], f"={cols[0]}{R['cash']}", colour=MUTED)
    for t_i in range(1, n + 1):
        c, prev = cols[t_i], cols[t_i - 1]
        _put(ws, R["cash"], c,
             f"={c}{R['noi']}-{c}{R['capex']}-{c}{R['bonus']}*{c}{R['frac']}"
             f"+{c}{R['amort']}", bold=True)
        ws[f"{c}{R['cash']}"].border = TOP_RULE
        _put(ws, R["cum"], c, f"={prev}{R['cum']}+{c}{R['cash']}", colour=MUTED)
    ws[f"C{R['cash']}"].border = TOP_RULE

    # ---- the answer ----
    cash0 = f"{cols[0]}{R['cash']}"
    cash_range = f"{cols[0]}{R['cash']}:{last}{R['cash']}"
    _label(ws, R["roi"], "ROI", bold=True)
    # Excel's RATE is iterative and returns garbage rather than #NUM! when no
    # real rate exists, so guard it the way the app does: a return only exists
    # if the cash flows actually turn the outlay positive.
    # RATE takes a fractional nper, so a 6.25-year term annualises over 6.25.
    roi = _put(ws, R["roi"], cols[0],
               f'=IFERROR(IF(-SUM({cash_range})/{cash0}<=0,"n/a",'
               f'RATE($D${R["term_row"]},,{cash0},SUM({cash_range}))),"n/a")',
               PCT, GREEN, bold=True, colour=GREEN_INK)
    roi.border = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)

    _label(ws, R["total"], "Total Net Cash")
    _put(ws, R["total"], cols[0], f"=SUM({cash_range})")

    cum_range = f"{cols[0]}{R['cum']}:{last}{R['cum']}"
    frac_range = f"{cols[0]}{R['frac']}:{last}{R['frac']}"
    idx = f"MATCH(TRUE,INDEX({cum_range}>=0,0),0)"
    _label(ws, R["payback"], "Payback (years)")
    # The year it turns positive may be a stub, which spans less than a
    # year, so the interpolated fraction is scaled by that year's share.
    _put(ws, R["payback"], cols[0],
         f'=IF(MAX({cum_range})<0,"never",IF({cols[0]}{R["cum"]}>=0,0,'
         f'({idx}-2)+INDEX({frac_range},{idx})*(-INDEX({cum_range},{idx}-1))/'
         f'(INDEX({cum_range},{idx})-INDEX({cum_range},{idx}-1))))',
         "0.0")

    # ---- commission by category, the source of the blended rate ----
    _label(ws, R["cat_head"], "Contract Commission Rates", bold=True)
    for i, cat in enumerate(CATEGORIES):
        _label(ws, R["rate0"] + i, cat, indent=1)
        _put(ws, R["rate0"] + i, "D", p["commission_rates"][i] / 100.0,
             PCT, YELLOW)
    _label(ws, R["mix_head"], "Sales By Category", bold=True)
    for i, cat in enumerate(CATEGORIES):
        _label(ws, R["mix0"] + i, cat, indent=1)
        _put(ws, R["mix0"] + i, "D", p["sales_mix"][i] / 100.0, PCT, YELLOW)
    _label(ws, R["mix0"] + len(CATEGORIES), "Total", bold=True)
    _put(ws, R["mix0"] + len(CATEGORIES), "D",
         f"=SUM({mix_span})", PCT, bold=True)

    # ---- what each category costs to deliver, as a share of its own income ----
    if cogs_rates:
        _label(ws, R["cogs_head"], "COGS Rates By Category", bold=True)
        for i, cat in enumerate(COGS_CATEGORIES):
            _label(ws, R["cogs_rate0"] + i, cat, indent=1)
            _put(ws, R["cogs_rate0"] + i, "D", cogs_rates[i] / 100.0, PCT,
                 YELLOW)

    # Revenue-weighted: total NOI over total net revenue across the term,
    # spanning the operating years only.
    noi_span = f"{cols[1]}{R['noi']}:{last}{R['noi']}"
    net_span = f"{cols[1]}{R['net']}:{last}{R['net']}"
    _label(ws, R["avgmargin"], "Avg NOI Margin")
    _put(ws, R["avgmargin"], cols[0],
         f"=IFERROR(SUM({noi_span})/SUM({net_span}),0)", PCT)

    # ---- staffing roster: the workbook's C45:G59, with the per-year merit
    # build-out that its "Payroll Calculation" tab performs ----
    _label(ws, R["staff_head"], "Staff Roster", bold=True)
    heads = [("D", "Start month"), ("E", "Salary"), ("F", "Benefit load"),
             ("G", "With benefits")]
    for t_i in range(1, n + 1):
        heads.append((ycol(t_i), f"Year {t_i}"))
        ws.column_dimensions[ycol(t_i)].width = 14
    for col, cap in heads:
        h = ws[f"{col}{R['staff_head']}"]
        h.value = cap
        h.font = Font(name="Calibri", size=10, bold=True, color=MUTED)
        h.alignment = Alignment(horizontal="right")

    merit_ref = f"$D${R['merit']}"
    for i, title in enumerate(titles):
        row = R["staff0"] + i
        _label(ws, row, title, indent=1)
        _put(ws, row, "D", starts[i], "0", YELLOW)
        _put(ws, row, "E", p["staff_salaries"][i], MONEY, YELLOW)
        _put(ws, row, "F", p["staff_loads"][i] / 100.0, PCT)
        _put(ws, row, "G", f"=$E{row}*(1+$F{row})")
        for t_i in range(1, n + 1):
            # Months of service either side of the merit step this year spans.
            # A stub year is the opening months of its year, so its end is
            # pulled back by the year fraction - a raise landing inside the
            # stub is still dated by service month.
            fr = f"${cols[t_i]}${R['frac']}"
            start = f"MAX(0,12*({t_i}-1)-$D{row})"
            end = f"MAX(0,12*({t_i}-1)+12*{fr}-$D{row})"
            k = f"INT({start}/12)"
            cap = f"12*({k}+1)"
            at_k = f"(MIN({end},{cap})-{start})"
            at_k1 = f"({end}-MIN({end},{cap}))"
            _put(ws, row, ycol(t_i),
                 f"=$G{row}*((1+{merit_ref})^{k}*{at_k}"
                 f"+(1+{merit_ref})^({k}+1)*{at_k1})/12")

    total_row = R["staff0"] + n_staff
    _label(ws, total_row, "Roster cost", bold=True)
    for t_i in range(1, n + 1):
        yc = ycol(t_i)
        _put(ws, total_row, yc, f"=SUM({yc}{first}:{yc}{last_s})", MONEY,
             bold=True)

    # ---- what year 0 costs: capex plus the roster's pre-opening month,
    # a start-up expense and a signing bonus ----
    _label(ws, R["pay0"], "Year-0 Payroll", bold=True)
    _put(ws, R["pay0"], cols[0], f"=SUM($G${first}:$G${last_s})/12", MONEY)
    _label(ws, R["exp0"], "Year-0 Expenses")
    _put(ws, R["exp0"], cols[0],
         f"={cols[1]}{R['rate']}*$D${R['exp0_pct']}", MONEY)
    _label(ws, R["exp0_pct"], "Year-0 Expenses %", indent=1)
    _put(ws, R["exp0_pct"], "D", p.get("year0_expense_pct", 0.0) / 100.0,
         PCT, YELLOW)
    _label(ws, R["bonus"], "Signing Bonus")
    for t_i in range(n + 1):
        _put(ws, R["bonus"], cols[t_i], bonus[t_i], MONEY, YELLOW)
    # A one-off, so only year 0 carries it; later years net capex and the
    # recurring bonus off directly in Net Cash.
    _label(ws, R["invest"], "Initial Investment", bold=True)
    _put(ws, R["invest"], cols[0],
         f"={cols[0]}{R['capex']}+{cols[0]}{R['pay0']}"
         f"+{cols[0]}{R['exp0']}+{cols[0]}{R['bonus']}", bold=True)

    _label(ws, R["merit"], "Merit Increase")
    _put(ws, R["merit"], "D", p.get("merit", 0.0) / 100.0, PCT, YELLOW)
    _label(ws, R["incentive"], "Incentives & Bonus (of gross revenue)")
    _put(ws, R["incentive"], "D", p.get("incentive_pct", 0.0) / 100.0, PCT, YELLOW)
    _label(ws, R["profit"], "Profit Sharing (of net revenue)")
    _put(ws, R["profit"], "D", p.get("profit_share_pct", 0.0) / 100.0, PCT, YELLOW)

    # ---- what makes a partial year partial ----
    # The term is an input like any other: type 6.25 here and every figure
    # above reprices to six years and a three-month stub.
    _label(ws, R["term_row"], "Contract Term (years)", bold=True)
    _put(ws, R["term_row"], "D", term, "0.00", YELLOW)

    _label(ws, R["rate"], "Gross Revenue (annual run rate)")
    _put(ws, R["rate"], cols[0], f"={cols[0]}{R['rev']}", colour=MUTED)
    for t_i in range(1, n + 1):
        _put(ws, R["rate"], cols[t_i],
             f"={cols[t_i - 1]}{R['rate']}*(1+{cols[t_i]}{R['growth']})",
             colour=MUTED)

    _label(ws, R["frac"], "Year Fraction Billed", indent=1)
    # MEDIAN(0, x, 1) is Excel's clamp: whole years give 1, the stub gives
    # what is left of the term, and anything past it gives 0.
    _put(ws, R["frac"], cols[0], 1.0, PCT, colour=MUTED)
    for t_i in range(1, n + 1):
        _put(ws, R["frac"], cols[t_i],
             f"=MEDIAN(0,$D${R['term_row']}-{t_i - 1},1)", PCT, colour=MUTED)

    legend = ("Yellow cells are inputs - change any of them and every "
              "figure above, ROI included, recalculates.")
    if cogs_rates:
        legend += " The blue block is COGS broken out by income category."
    note = ws.cell(row=R["legend"], column=3, value=legend)
    note.font = Font(name="Calibri", size=9, italic=True, color=MUTED)
    ws.freeze_panes = f"D{R['year'] + 1}"
    return ws


def write_workbook(path, scenarios, active=0, names=None, note=""):
    """`scenarios` is a list of model-input dicts, one per scenario.

    `note` is the free-text label from the app header; it is repeated on
    every sheet so each one still says what it belongs to when read alone.
    """
    names = names or [f"Scenario {i + 1}" for i in range(len(scenarios))]
    wb = Workbook()
    wb.remove(wb.active)

    summary = wb.create_sheet("Comparison")
    summary.sheet_view.showGridLines = False
    summary.column_dimensions["A"].width = 2
    summary.column_dimensions["B"].width = 2
    summary.column_dimensions["C"].width = 26

    t = summary.cell(row=2, column=3, value="Hotel ROI Modeling Tool - scenario comparison")
    t.font = Font(name="Calibri", size=14, bold=True, color=INK)
    if note:
        _note_cell(summary, "D2", note)
    sub = summary.cell(row=3, column=3,
                       value="Every figure below is a live reference into the "
                             "scenario sheets.")
    sub.font = Font(name="Calibri", size=9, italic=True, color=MUTED)

    for i, name in enumerate(names):
        col = get_column_letter(4 + i)
        summary.column_dimensions[col].width = 16
        head = summary[f"{col}5"]
        head.value = name + (" (active)" if i == active else "")
        head.font = Font(name="Calibri", size=11, bold=True, color=MUTED)
        head.fill = HEAD_FILL
        head.alignment = Alignment(horizontal="right")

    rows = [("Contract Term (years)", "term", "0.00", False),
            ("ROI", R["roi"], PCT, True),
            ("Total Net Cash", R["total"], MONEY, False),
            ("Payback (years)", R["payback"], "0.0", False),
            ("Avg NOI Margin", R["avgmargin"], PCT, False)]
    for r_i, (label, src_row, fmt, is_roi) in enumerate(rows, start=6):
        c = summary.cell(row=r_i, column=3, value=label)
        c.font = Font(name="Calibri", size=11, bold=is_roi, color=INK)
        for i, name in enumerate(names):
            cell = summary.cell(row=r_i, column=4 + i)
            if src_row == "term":
                # A term may be fractional, so read the input cell itself
                # rather than counting year columns.
                cell.value = f"='{name}'!$D${R['term_row']}"
            else:
                cell.value = f"='{name}'!{COL[0]}{src_row}"
            cell.number_format = fmt
            cell.font = Font(name="Calibri", size=11, bold=is_roi,
                             color=GREEN_INK if is_roi else INK)
            if is_roi:
                cell.fill = GREEN

    for name, p in zip(names, scenarios):
        _sheet(wb, name, p, note)

    wb.save(path)
    return path
