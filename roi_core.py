"""The financial model from "App Model.xlsx", shared by both front ends.

The term need not be whole: 6.25 runs six operating years and a three-month
stub. Each year carries a weight (term_weights) that is its share of a full
year, and:

    every flow    is booked at its annual rate * that year's weight
    growth        compounds on the annualised run rate, not the booked figure
    ROI           is annualised over the fractional term: ^(1/6.25)

The workbook models a contract term of N operating years plus a Year 0 capex
outlay:

    Commission %     = SUMPRODUCT(commission rate by category, sales mix)
    Gross Revenue    Y0 = input, Yn = Yn-1 * (1 + growth_n)
    Commission       = Gross Revenue * commission %
    Net Revenue      = Gross Revenue - Commission
    COGS             = Net Revenue * COGS_n %
    Payroll          roster * (1 + merit)^(n-1) + incentive % * Gross
                     + profit sharing % * Net Revenue
    Expenses         = Net Revenue * expenses_n %
    Bonus Amort      = year-0 signing bonus / term + this year's bonus
    NOI              = Net Revenue - COGS - Payroll - Expenses - Bonus Amort
    Capex            Y0 = Gross Revenue Y1 * initial capex %, Yn = Gross Revenue n * capex %
    Initial Invest   = Capex + year-0 payroll + year-0 expense + signing bonus
    Net Cash         Y0 = -Initial Investment
                     Yn = NOI - Capex - bonus + Bonus Amort (non-cash)
    ROI              = RATE(N, , NetCash_Y0, SUM(NetCash_Y0:Y_N))

RATE with no payment reduces to a closed form:
    pv * (1 + r)^N + fv = 0   ->   r = (-fv / pv)^(1/N) - 1

Verified against the workbook's cached values: every line item matches to
floating-point precision and ROI agrees to within 1e-16.

This module is deliberately free of UI imports so the desktop app
(roi_model.py) and the Dash app (app_dash.py) cannot drift apart.
"""

MAX_YEARS = 10      # the longest contract term the app offers
DEFAULT_TERM = 6    # what the workbook itself models
YEARS = DEFAULT_TERM  # back-compat alias for callers that assume a fixed term
TERM_STEP = 0.25    # terms are set in quarter-year steps


def year_count(term):
    """How many operating-year columns a term needs. A fractional term still
    gets a whole column for its stub: 6.25 runs six years and a 3-month tail."""
    term = max(TERM_STEP, min(float(MAX_YEARS), float(term)))
    n = int(term)
    if term - n > 1e-9:
        n += 1
    return max(1, n)


def term_weights(term):
    """Each operating year's share of a full year, so a stub books only the
    months inside the term. The weights sum back to `term`."""
    term = max(TERM_STEP, min(float(MAX_YEARS), float(term)))
    n = year_count(term)
    w = [1.0] * n
    tail = term - (n - 1)
    if tail < 1.0:
        w[-1] = round(tail, 9)
    return w

# Commission is no longer a single rate: the workbook blends a per-category
# contract rate against the share of sales that category represents.
CATEGORIES = [
    "Labor Income",
    "Equipment Income",
    "Logistics Income",
    "Consumable / Sale Item Income",
    "Creative Content Income",
    "Rigging Income",
    "Power Income",
    "HSIA Income",
]

# Payroll is built from a staffing roster on the workbook's "Payroll
# Calculation" tab: each role costs salary plus a benefit load, the whole
# roster escalates by the merit increase each year, and incentives and profit
# sharing ride on top as a share of revenue. Title, salary, benefit load %.
STAFF = [
    # type, title, start month, salary, benefit load %
    ("Salaried Full Time Staff", "Director", -1, 78_000.0, 13.3),
    ("Salaried Full Time Staff", "Technical Coordinator 1", -1, 55_000.0, 13.3),
    ("Salaried Full Time Staff", "Technical Coordinator 2", -1, 0.0, 13.3),
    ("Salaried Full Time Staff", "Technical Coordinator 3", -1, 0.0, 13.3),
    ("Salaried Full Time Staff", "Technical Coordinator 4", -1, 0.0, 13.3),
    ("Hourly Full Time Staff", "Technician 1", -1, 45_670.0, 13.3),
    ("Hourly Full Time Staff", "Technician 2", -1, 0.0, 13.3),
    ("Hourly Full Time Staff", "Technician 3", -1, 0.0, 13.3),
    ("Hourly Full Time Staff", "Technician 4", -1, 0.0, 13.3),
    ("Hourly Full Time Staff", "Technician 5", -1, 0.0, 13.3),
    ("Hourly /Part Time / Seasonal Positions", "Technician 6", -1, 22_880.0, 8.85),
    ("Hourly /Part Time / Seasonal Positions", "Technician 7", -1, 0.0, 8.85),
    ("Hourly /Part Time / Seasonal Positions", "Technician 8", -1, 0.0, 8.85),
    ("Hourly /Part Time / Seasonal Positions", "Technician 9", -1, 0.0, 8.85),
]
STAFF_TYPES = [t for t, _, _, _, _ in STAFF]
STAFF_TITLES = [t for _, t, _, _, _ in STAFF]


def staff_base(salaries, loads):
    """Year-1 payroll before merit: every salary grossed up by its benefits."""
    return sum(s * (1 + l / 100.0) for s, l in zip(salaries, loads))


def role_cost(loaded, merit, start_month, n, weight=1.0):
    """One role's cost in operating year n, following the workbook's
    "Payroll Calculation" tab.

    Service months decide which merit step a role sits on, and a year can
    straddle two steps, so the two portions are weighted by months. A role
    starting before year 1 (a negative start month) has banked service and
    reaches its next raise sooner.

    `weight` is the year's share of a full year. A stub bills only its own
    months, counted from the year's start, so a merit step landing inside the
    stub is still dated by month rather than smeared across it.
    """
    start = max(0, 12 * (n - 1) - start_month)
    end = max(0, 12 * n - start_month)
    if weight < 1.0:
        end = max(0, min(end, 12 * (n - 1) + weight * 12 - start_month))
    if end - start <= 0:
        return 0.0
    k = int(start // 12)              # completed service years at year start
    cap = 12 * (k + 1)                # month the next merit step lands
    at_k = min(end, cap) - start
    at_k1 = end - min(end, cap)
    m = merit / 100.0
    return loaded * ((1 + m) ** k * at_k + (1 + m) ** (k + 1) * at_k1) / 12.0


# Opening values, mirroring the yellow input cells of App Model.xlsx.
# Per-year lists run to MAX_YEARS so a longer term still has values to use;
# only the first `term` entries are ever read.
DEFAULTS = {
    "term": DEFAULT_TERM,
    "gross_revenue": 967_000.0,
    "rev_growth": [0.0] + [3.0] * (MAX_YEARS - 1),      # % per year, years 1..N
    # % commission the contract pays on each category of income
    "commission_rates": [49.0, 49.0, 49.0, 49.0, 30.0, 30.0, 75.0, 75.0],
    # % of sales each category represents; these are meant to total 100
    "sales_mix": [40.03, 56.11, 1.96,
                  0.00, 0.14, 0.02,
                  0.00, 1.74],
    "cogs": 10.0,             # % of net revenue
    "expenses": 6.5,          # % of net revenue
    "staff_titles": [t for _, t, _, _, _ in STAFF],
    "staff_starts": [m for _, _, m, _, _ in STAFF],     # month relative to year 1
    "staff_salaries": [s for _, _, _, s, _ in STAFF],   # $ per role
    "staff_loads": [l for _, _, _, _, l in STAFF],      # % benefit load per role
    "merit": 3.0,             # % the roster escalates each year
    "incentive_pct": 1.6,     # % of gross revenue
    "profit_share_pct": 0.89,  # % of net revenue
    "capex_initial": 25.15,   # % of year-1 gross revenue
    "capex": [0.0] + [3.15] * (MAX_YEARS - 1),          # % of gross revenue
    # Year 0 also carries a month of payroll, a start-up expense and a
    # signing bonus; together with capex these make the Initial Investment.
    "year0_expense_pct": 0.7172547150951487,   # % of year-1 gross revenue
    # Year 0 has its own bonus cell (D76); every operating year carries one
    # too (E79:J79), and Net Cash nets it off each year.
    "signing_bonus": [60_000.0] + [0.0] * MAX_YEARS,    # $ per year
}


# A single blended rate, for callers that do not model the categories
# (the Dash front end still shows one Commission slider).
DEFAULTS["commission"] = sum(
    r * m for r, m in zip(DEFAULTS["commission_rates"], DEFAULTS["sales_mix"])
) / 100.0


def as_years(default, count):
    """A per-year list of exactly `count` entries.

    Accepts a single value or a list; a short list is padded with its last
    entry and a long one truncated, so changing the term never leaves a
    series the wrong length.
    """
    if not isinstance(default, (list, tuple)):
        return [default] * count
    vals = list(default)
    if len(vals) >= count:
        return vals[:count]
    return vals + [vals[-1] if vals else 0.0] * (count - len(vals))


def is_uniform(values):
    return len(set(values)) == 1


def lead_value(values):
    """What a per-year series collapses to on its single lead slider: the
    most common entry, so collapsing keeps the dominant rate."""
    return max(set(values), key=values.count)


def blended_commission(rates, mix):
    """The workbook's SUMPRODUCT of contract rates against the sales mix.

    Both arrive in display units (49.0 meaning 49%), so the product of two
    percentages is divided back down once.
    """
    return sum(r * m for r, m in zip(rates, mix)) / 100.0


# Line items as they appear in the workbook: label, series key, format.
# "_growth" is the revenue-growth input echoed back as a row.
ROWS = [
    ("Gross Revenue", "rev", "money"),
    ("Revenue Growth", "_growth", "pct"),
    ("Commission", "commission", "money"),
    ("Net Revenue", "net_rev", "money"),
    ("COGS", "cogs", "money"),
    ("Payroll", "payroll", "money"),
    ("Payroll % of Net Revenue", "payroll_pct", "pct"),
    ("Expenses", "expenses", "money"),
    ("Bonus Amortization", "amort", "money"),
    ("NOI", "noi", "money"),
    ("NOI Margin", "margin", "pct"),
    ("Capex Investment", "capex", "money"),
    ("Initial Investment", "invest", "money"),
    ("Net Cash", "net_cash", "money"),
    ("Cumulative Net Cash", "cumulative", "money"),
]

# Year 0 only carries these lines in the workbook; the rest stay blank.
YEAR0 = {"rev", "capex", "invest", "net_cash", "cumulative"}

# The reverse: lines the workbook computes for year 0 alone. Initial
# Investment is a one-off, so the operating years are left blank.
ONLY_YEAR0 = {"invest"}


def commission_of(p):
    """The blended commission rate for a parameter set, in display units."""
    rates, mix = p.get("commission_rates"), p.get("sales_mix")
    if rates and mix:
        return blended_commission(rates, mix)
    return p["commission"]


def compute(p):
    """Run the model. Percentages arrive in display units (3.0 means 3%).

    `term` is the contract length in operating years and need not be whole:
    6.25 runs six years and a three-month stub. Every returned series is
    year_count(term)+1 long, index 0 being Year 0, so a stub gets a column
    of its own carrying only the months inside the term. The year-varying
    parameters - rev_growth, cogs, expenses, capex and payroll_growth -
    accept either a single value applied to every year or a per-year list.
    """
    term = float(p.get("term", DEFAULT_TERM))
    term = max(TERM_STEP, min(float(MAX_YEARS), term))
    n = year_count(term)
    weight = term_weights(term)     # weight[t - 1] is operating year t's share

    growth = as_years(p["rev_growth"], n)
    cogs_pct = as_years(p["cogs"], n)
    exp_pct = as_years(p["expenses"], n)
    capex_pct = as_years(p["capex"], n)
    comm_pct = commission_of(p)
    merit = p.get("merit", 0.0)
    incentive = p.get("incentive_pct", 0.0)
    profit_share = p.get("profit_share_pct", 0.0)
    # Either a full roster or a single pre-computed year-1 staff cost.
    roster = "staff_salaries" in p
    if roster:
        loaded = [s * (1 + l / 100.0)
                  for s, l in zip(p["staff_salaries"], p["staff_loads"])]
        starts = p.get("staff_starts") or [0] * len(loaded)
    base = staff_base(p["staff_salaries"], p["staff_loads"]) if roster \
        else p.get("staff_base", 0.0)

    z = [0.0] * (n + 1)
    rev, comm, net_rev = z[:], z[:], z[:]
    cogs, payroll, expenses = z[:], z[:], z[:]
    noi, capex = z[:], z[:]
    rate = z[:]

    # `rate` is what a year would bill in full; `rev` is what the contract
    # actually books, which is less in a stub year. Growth compounds on the
    # run rate, never on the booked figure, so the escalator still steps a
    # whole notch at the anniversary and only the months billed prorate.
    rate[0] = p["gross_revenue"]
    rev[0] = rate[0]
    for t in range(1, n + 1):
        rate[t] = rate[t - 1] * (1 + growth[t - 1] / 100.0)
        rev[t] = rate[t] * weight[t - 1]
        comm[t] = rev[t] * comm_pct / 100.0
        net_rev[t] = rev[t] - comm[t]
        cogs[t] = net_rev[t] * cogs_pct[t - 1] / 100.0
        expenses[t] = net_rev[t] * exp_pct[t - 1] / 100.0

    # Roster escalated by merit, plus the revenue-linked overlays. The
    # overlays ride on `rev`, so they prorate with it; the roster prorates
    # inside role_cost, which bills the stub months at their merit step.
    for t in range(1, n + 1):
        if roster:
            staff = sum(role_cost(b, merit, st, t, weight[t - 1])
                        for b, st in zip(loaded, starts))
        else:
            # Same month-accurate rule, applied to the aggregate: with the
            # default start of 0 this is exactly base * (1 + merit)^(t-1).
            staff = role_cost(base, merit, p.get("staff_start", 0), t,
                              weight[t - 1])
        payroll[t] = (staff + rev[t] * incentive / 100.0
                      + net_rev[t] * profit_share / 100.0)

    # The year-0 signing bonus is written off over the term, and each year's
    # own bonus with it. This reduces NOI but is not a cash movement, so it is
    # added straight back below - it changes reported margin, never cash.
    bonus_raw = p.get("signing_bonus", 0.0)
    if isinstance(bonus_raw, (list, tuple)):
        bonus = (list(bonus_raw) + [0.0] * (n + 1))[:n + 1]
    else:
        bonus = [float(bonus_raw)] + [0.0] * n
    # A recurring bonus is an annual cost, so a stub year pays its share.
    # The year-0 one is a one-off and is never prorated.
    bonus = [bonus[0]] + [bonus[t] * weight[t - 1] for t in range(1, n + 1)]
    # Spread over the fractional term: the weights sum back to `term`, so
    # the write-off still lands on exactly the year-0 bonus.
    amort = [0.0] * (n + 1)
    for t in range(1, n + 1):
        amort[t] = bonus[0] / term * weight[t - 1] + bonus[t]

    for t in range(1, n + 1):
        noi[t] = net_rev[t] - cogs[t] - payroll[t] - expenses[t] - amort[t]

    # The workbook keys the year-0 outlay off year-1 gross revenue (=+E4*D72).
    # That is a full year of billing, so it reads the run rate.
    capex[0] = rate[1] * p["capex_initial"] / 100.0
    for t in range(1, n + 1):
        capex[t] = rev[t] * capex_pct[t - 1] / 100.0

    # Year 0 carries the roster's pre-opening month, a start-up expense and a
    # signing bonus on top of capex. Later years are capex plus any bonus.
    if roster:
        pay0 = sum(role_cost(b, merit, st, 0) for b, st in zip(loaded, starts))
    else:
        pay0 = role_cost(base, merit, p.get("staff_start", 0), 0)
    exp0 = rate[1] * p.get("year0_expense_pct", 0.0) / 100.0

    invest = [capex[t] + bonus[t] for t in range(n + 1)]
    invest[0] += pay0 + exp0

    # Year 0 is the whole initial outlay. Later years take capex and the
    # year's bonus off, then add the non-cash amortization back.
    net_cash = [noi[0] - invest[0]] + [
        noi[t] - capex[t] - bonus[t] + amort[t] for t in range(1, n + 1)]

    cumulative, run = [], 0.0
    for v in net_cash:
        run += v
        cumulative.append(run)

    total = cumulative[-1]
    pv = net_cash[0]

    # RATE(term, , pv, total) with pmt = 0
    roi = None
    if pv != 0:
        ratio = -total / pv
        if ratio > 0:
            roi = ratio ** (1.0 / term) - 1.0

    # Payback year, linearly interpolated across the year it turns positive.
    # A stub spans only its own length, so the fraction is scaled by that
    # year's weight: a quarter of the way into a 3-month stub is 0.0625
    # years, not a quarter of a year.
    payback = None
    for t in range(1, n + 1):
        if cumulative[t] >= 0:
            prev = cumulative[t - 1]
            step = cumulative[t] - prev
            into = (-prev / step if step else 0.0)
            payback = (t - 1) + weight[t - 1] * into
            break

    margin = [noi[t] / net_rev[t] if net_rev[t] else 0.0
              for t in range(n + 1)]
    # The workbook's row 41: what payroll is costing as a share of net revenue.
    payroll_pct = [payroll[t] / net_rev[t] if net_rev[t] else 0.0
                   for t in range(n + 1)]
    # Revenue-weighted: total NOI over total net revenue across the term, so
    # a bigger year counts for more than a small one. Year 0 has no revenue
    # and so takes no part in it.
    span_net = sum(net_rev[1:])
    avg_margin = (sum(noi[1:]) / span_net) if span_net else 0.0

    return {
        "term": term, "years": n, "weights": weight, "rate": rate,
        "commission_pct": comm_pct, "staff_base": base,
        "invest": invest, "amort": amort,
        "year0_payroll": pay0, "year0_expense": exp0,
        "year0_bonus": bonus[0],
        "margin": margin, "avg_margin": avg_margin,
        "payroll_pct": payroll_pct,
        "rev": rev, "commission": comm, "net_rev": net_rev, "cogs": cogs,
        "payroll": payroll, "expenses": expenses, "noi": noi, "capex": capex,
        "net_cash": net_cash, "cumulative": cumulative,
        "_growth": [0.0] + [g / 100.0 for g in growth],
        "total": total, "roi": roi, "payback": payback,
    }


def money(v):
    return f"({abs(v):,.0f})" if v < 0 else f"{v:,.0f}"


def pct(v, dp=1):
    return f"{v * 100:.{dp}f}%"
