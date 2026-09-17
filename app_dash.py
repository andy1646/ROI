"""Hotel ROI Modeling Tool - Plotly Dash front end for the model in roi_core.py.

Run it:      python app_dash.py       ->  http://127.0.0.1:8050

Same model and the same design language as the desktop build (roi_model.py);
styling lives in assets/roi.css, which Dash picks up automatically.
"""

import webbrowser

import plotly.graph_objects as go
from dash import ALL, Dash, Input, Output, dcc, html

from roi_core import (DEFAULTS, ROWS, YEAR0, YEARS, as_years, compute,
                      is_uniform, lead_value, money, pct, staff_base)

# palette, kept in step with assets/roi.css
ACCENT_2 = "#10b981"
NEG_2 = "#f43f5e"
INK_2 = "#334155"
MUTED = "#8794a7"
HAIR = "#e8ecf1"
CARD = "#ffffff"

FONT = ('"Segoe UI Variable Text", "Segoe UI", system-ui, -apple-system, '
        '"Helvetica Neue", Arial, sans-serif')

# The three parameters the workbook varies year by year. Each gets one lead
# slider plus an optional per-year breakdown, exactly as the desktop app does.
GROUPS = {
    "growth":   dict(label="Revenue Growth", lo=-10, hi=30, step=0.25,
                     default=DEFAULTS["rev_growth"], first=1, count=YEARS),
    "cogs":     dict(label="COGS", lo=0, hi=60, step=0.25,
                     default=DEFAULTS["cogs"], first=1, count=YEARS,
                     note="% of net revenue"),
    "expenses": dict(label="Expenses", lo=0, hi=60, step=0.25,
                     default=DEFAULTS["expenses"], first=1, count=YEARS,
                     note="% of net revenue"),
    "capex":    dict(label="Ongoing Capital Equipment", lo=0, hi=30, step=0.25,
                     default=DEFAULTS["capex"], first=1, count=YEARS),
}
# Fixed order: the callbacks below pack their inputs and outputs by position.
GROUP_ORDER = ("growth", "cogs", "expenses", "capex")


# ------------------------------------------------------------ components ---
def slider(sid, label, lo, hi, step, value, kind="pct", note=None):
    """A labelled slider. Dash 4's slider renders its own numeric box, which
    doubles as the editable value field the desktop app has."""
    slider_id = sid if isinstance(sid, str) else {**sid, "el": "slider"}
    return html.Div(className="ctl", children=[
        html.Div(className="row", children=[html.Label(label)]),
        dcc.Slider(
            id=slider_id,
            min=lo, max=hi, step=step, value=value,
            marks=None, updatemode="drag", tooltip=None,
            allow_direct_input=True, className="slider",
        ),
        html.Div(className="ends", children=[
            html.Span(f"{lo:,.0f}" if kind == "money" else f"{lo:g}"),
            html.Span(note, className="mid") if note else html.Span(),
            html.Span(f"{hi:,.0f}" if kind == "money" else f"{hi:g}"),
        ]),
    ])


def group(key):
    """Lead slider + collapsible per-year sliders."""
    g = GROUPS[key]
    values = as_years(g["default"], g["count"])
    years = [
        slider({"type": f"yr-{key}", "i": i}, f"Year {g['first'] + i}",
               g["lo"], g["hi"], g["step"], values[i])
        for i in range(g["count"])
    ]
    # A series the workbook varies year by year opens already expanded.
    varies = not is_uniform(values)
    return html.Div(className="ctl-group", children=[
        slider(f"{key}-lead", g["label"], g["lo"], g["hi"], g["step"],
               lead_value(values), note=g.get("note")),
        html.Div(className="toggle", children=[
            dcc.Checklist(
                id=f"{key}-toggle", value=["on"] if varies else [],
                options=[{"label": f" set each year individually "
                                   f"({g['count']} years)", "value": "on"}],
            )
        ]),
        html.Div(years, id=f"{key}-panel", className="per-year",
                 style={"display": "block" if varies else "none"}),
    ])


def section(title, *children):
    return html.Div(className="section",
                    children=[html.H2(title), *children])


D = DEFAULTS
controls = html.Div(className="card", children=[
    section("Revenue",
            slider("rev", "Gross Revenue (Year 0)", 50_000, 5_000_000, 10_000,
                   D["gross_revenue"], kind="money"),
            group("growth")),
    section("Cost Structure",
            slider("commission", "Commission", 0, 90, 0.25, D["commission"],
                   note="% of gross revenue"),
            group("cogs"),
            group("expenses")),
    section("Payroll",
            slider("staffbase", "Roster Cost (Year 1)", 0, 2_000_000, 500,
                   round(staff_base(D["staff_salaries"], D["staff_loads"]), 2),
                   kind="money"),
            slider("merit", "Merit Increase", 0, 15, 0.25, D["merit"],
                   note="a year, across the roster"),
            slider("incentive", "Incentives & Bonus", 0, 10, 0.25,
                   D["incentive_pct"], note="% of gross revenue"),
            slider("profit", "Profit Sharing", 0, 10, 0.25,
                   D["profit_share_pct"], note="% of net revenue")),
    section("Capital",
            slider("capex0", "Capital Equipment (Pre Move In)", 0, 150, 0.25,
                   D["capex_initial"], note="% of year-1 gross revenue"),
            slider("year0exp", "Pre Move In Expenses", 0, 10, 0.05,
                   D["year0_expense_pct"], note="% of year-1 gross revenue"),
            slider("bonus", "Signing Bonus (Pre Move In)", 0, 500_000, 1_000,
                   D["signing_bonus"][0], kind="money"),
            slider("bonusyr", "Signing Bonus (Yearly)", 0, 500_000, 1_000,
                   D["signing_bonus"][1], kind="money"),
            group("capex")),
])

STATS = [("invest", "Year-0 Capex"), ("total", "Total Net Cash"),
         ("payback", "Payback"), ("margin", "Avg NOI Margin")]

hero = html.Div(className="hero", children=[
    html.Div(className="lede", children=[
        html.P("Return on Investment", className="eyebrow"),
        html.P("-", className="roi", id="roi"),
        html.P("annualized on the year-0 capex, over six years", className="note"),
    ]),
    html.Div(className="stats", children=[
        html.Div(className="stat", children=[
            html.P(cap, className="k"),
            html.P("-", className="v", id=f"stat-{key}"),
        ]) for key, cap in STATS
    ]),
])

results = html.Div(className="stack", children=[
    hero,
    html.Div(className="card", children=[
        html.Div(className="tbl-wrap", children=[
            html.Table(className="model", children=[
                html.Colgroup(
                    [html.Col(style={"width": "20%"})]
                    + [html.Col(style={"width": f"{80 / (YEARS + 1):.4f}%"})
                       for _ in range(YEARS + 1)]),
                html.Thead(html.Tr(
                    [html.Th("")] + [html.Th(f"Year {i}") for i in range(YEARS + 1)])),
                html.Tbody(id="tbody"),
            ])
        ])
    ]),
    html.Div(className="card", children=[
        html.P("Net cash by year", className="chart-title"),
        dcc.Graph(id="chart", config={"displayModeBar": False},
                  style={"height": "232px"}),
    ]),
])

app = Dash(__name__, title="Hotel ROI Modeling Tool", update_title=None)
app.layout = html.Div(className="wrap", children=[
    html.Div(className="head", children=[
        html.Div([
            html.H1("Hotel ROI Modeling Tool"),
            html.P("Six-year cash-flow model  ·  App Model.xlsx", className="sub"),
        ]),
        html.Div(className="spacer"),
        html.Button("Reset", id="reset", className="btn", n_clicks=0),
    ]),
    html.Div(className="grid", children=[controls, results]),
])


# ------------------------------------------------------------- callbacks ---
for key in GROUPS:
    # Show or hide a per-year panel.
    app.callback(
        Output(f"{key}-panel", "style"),
        Input(f"{key}-toggle", "value"),
    )(lambda v: {"display": "block"} if v else {"display": "none"})

def _default_series():
    return {k: as_years(GROUPS[k]["default"], GROUPS[k]["count"])
            for k in GROUP_ORDER}


@app.callback(
    Output("rev", "value"), Output("commission", "value"),
    Output("staffbase", "value"), Output("merit", "value"),
    Output("incentive", "value"), Output("profit", "value"),
    Output("capex0", "value"), Output("year0exp", "value"),
    Output("bonus", "value"),
    Output("bonusyr", "value"),
    *[Output(f"{k}-lead", "value") for k in GROUP_ORDER],
    *[Output(f"{k}-toggle", "value") for k in GROUP_ORDER],
    *[Output({"type": f"yr-{k}", "i": ALL, "el": "slider"}, "value")
      for k in GROUP_ORDER],
    Input("reset", "n_clicks"),
    prevent_initial_call=True,
)
def reset(_n):
    """Restore every control to the workbook's opening state, per-year series
    and their expanded/collapsed state included."""
    series = _default_series()
    return (
        D["gross_revenue"], D["commission"],
        round(staff_base(D["staff_salaries"], D["staff_loads"]), 2),
        D["merit"], D["incentive_pct"], D["profit_share_pct"],
        D["capex_initial"], D["year0_expense_pct"],
        D["signing_bonus"][0], D["signing_bonus"][1],
        *[lead_value(series[k]) for k in GROUP_ORDER],
        *[(["on"] if not is_uniform(series[k]) else []) for k in GROUP_ORDER],
        *[series[k] for k in GROUP_ORDER],
    )


def figure(net_cash):
    """Net cash by year: rounded bars, a zero rule, values called out."""
    hi, lo = max(max(net_cash), 0.0), min(min(net_cash), 0.0)
    span = (hi - lo) or 1.0
    fig = go.Figure(go.Bar(
        x=[f"Year {i}" for i in range(YEARS + 1)],
        y=net_cash,
        marker=dict(color=[ACCENT_2 if v >= 0 else NEG_2 for v in net_cash],
                    cornerradius=5),
        text=[money(v) for v in net_cash],
        textposition="outside",
        textfont=dict(size=11, family=FONT,
                      color=[INK_2 if v >= 0 else NEG_2 for v in net_cash]),
        cliponaxis=False,
        hovertemplate="%{x}<br>Net cash %{customdata}<extra></extra>",
        customdata=[("$" + money(v)) for v in net_cash],
        width=0.44,
    ))
    fig.update_layout(
        margin=dict(l=8, r=8, t=22, b=4),
        paper_bgcolor=CARD, plot_bgcolor=CARD,
        font=dict(family=FONT, size=12, color=MUTED),
        showlegend=False, bargap=0.5,
        xaxis=dict(showgrid=False, showline=False, ticks="",
                   tickfont=dict(size=11.5, color=MUTED)),
        # Headroom top and bottom so the outside value labels are never clipped.
        yaxis=dict(visible=False, zeroline=True, zerolinecolor=HAIR,
                   zerolinewidth=1,
                   range=[lo - 0.22 * span, hi + 0.22 * span]),
        hoverlabel=dict(bgcolor="#101828", bordercolor="#101828",
                        font=dict(color="#e9eef5", family=FONT, size=12)),
    )
    return fig


@app.callback(
    Output("roi", "children"), Output("roi", "className"),
    Output("stat-invest", "children"), Output("stat-total", "children"),
    Output("stat-total", "className"), Output("stat-payback", "children"),
    Output("stat-margin", "children"),
    Output("tbody", "children"), Output("chart", "figure"),
    Input("rev", "value"), Input("commission", "value"),
    Input("staffbase", "value"), Input("merit", "value"),
    Input("incentive", "value"), Input("profit", "value"),
    Input("capex0", "value"), Input("year0exp", "value"),
    Input("bonus", "value"),
    Input("bonusyr", "value"),
    *[Input(f"{k}-lead", "value") for k in GROUP_ORDER],
    *[Input({"type": f"yr-{k}", "i": ALL, "el": "slider"}, "value")
      for k in GROUP_ORDER],
    *[Input(f"{k}-toggle", "value") for k in GROUP_ORDER],
)
def recalc(rev, commission, staffbase, merit, incentive, profit,
           capex0, year0exp, bonus, bonusyr, *group_values):
    # group_values arrives as leads, then per-year lists, then toggles -
    # one entry per group, in GROUP_ORDER.
    n = len(GROUP_ORDER)
    leads = dict(zip(GROUP_ORDER, group_values[:n]))
    years = dict(zip(GROUP_ORDER, group_values[n:2 * n]))
    shown = dict(zip(GROUP_ORDER, group_values[2 * n:3 * n]))

    def series(key):
        return (list(years[key]) if shown[key]
                else [leads[key]] * GROUPS[key]["count"])

    r = compute({
        "gross_revenue": rev,
        "rev_growth": series("growth"),
        "commission": commission,
        "cogs": series("cogs"),
        "expenses": series("expenses"),
        "staff_base": staffbase,
        "merit": merit,
        "incentive_pct": incentive,
        "profit_share_pct": profit,
        "capex_initial": capex0,
        "capex": series("capex"),
        # Year 0 also carries a month of payroll, a start-up expense and the
        # signing bonus, exactly as the desktop build does.
        "staff_start": -1,
        "year0_expense_pct": year0exp,
        "signing_bonus": [bonus] + [bonusyr] * YEARS,
    })

    roi_txt = pct(r["roi"], 2) if r["roi"] is not None else "n/a"
    # ROI colour bands, judged on the value as displayed (2 decimals):
    # red below 6.37%, yellow from 6.37% up to 10%, green at 10% and above.
    roi_pct = round(r["roi"] * 100, 2) if r["roi"] is not None else None
    if roi_pct is None or roi_pct < 6.37:
        roi_cls = "roi neg"
    elif roi_pct < 10:
        roi_cls = "roi mid"
    else:
        roi_cls = "roi"

    rows = []
    for label, key, kind in ROWS:
        cells = []
        for t in range(YEARS + 1):
            if t == 0 and key not in YEAR0:
                cells.append("")
            else:
                cells.append(money(r[key][t]) if kind == "money"
                             else pct(r[key][t]))
        # r- prefixed so a row class can never collide with a layout
        # class elsewhere in the stylesheet.
        tag = {"_growth": "r-soft", "net_rev": "r-head", "noi": "r-head",
               "margin": "r-soft", "payroll_pct": "r-soft",
               "net_cash": "r-cash", "cumulative": "r-soft"}.get(key, "")
        rows.append(html.Tr(className=tag,
                            children=[html.Td(label)] + [html.Td(c) for c in cells]))

    return (
        roi_txt, roi_cls,
        "$" + money(-r["net_cash"][0]),
        "$" + money(r["total"]), "v" if r["total"] >= 0 else "v neg",
        f"{r['payback']:.1f} yrs" if r["payback"] is not None else "never",
        pct(r["avg_margin"]),
        rows, figure(r["net_cash"]),
    )


if __name__ == "__main__":
    url = "http://127.0.0.1:8050"
    print(f"Hotel ROI Modeling Tool (Dash) running at {url}   -   Ctrl+C to stop")
    webbrowser.open(url)
    app.run(debug=False, port=8050)
