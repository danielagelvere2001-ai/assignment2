"""Regression exams from the Restaurant case (Excel Part 4).

The three sheets are merged the way the case asks (one row per trading day:
total sales, weather, game day, sales one year earlier), the merged table is
written to data/Restaurant_merged.csv so anyone can re-run every model with
Excel's Data Analysis → Regression, and ordinary least squares is fitted here
with the same definitions Excel uses (two-sided t-test p-values, R²).
"""
import csv
import datetime as dt
import math
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


# ------------------------------------------------------------------ statistics

def _betacf(a, b, x):
    qab, qap, qam = a + b, a + 1, a - 1
    c, d = 1.0, 1 - qab * x / qap
    d = 1 / (d if abs(d) > 1e-300 else 1e-300)
    h = d
    for m in range(1, 300):
        m2 = 2 * m
        aa = m * (b - m) * x / ((qam + m2) * (a + m2))
        d = 1 + aa * d
        d = 1 / (d if abs(d) > 1e-300 else 1e-300)
        c = 1 + aa / c if abs(1 + aa / c) > 1e-300 else 1e-300
        h *= d * c
        aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2))
        d = 1 + aa * d
        d = 1 / (d if abs(d) > 1e-300 else 1e-300)
        c = 1 + aa / c if abs(1 + aa / c) > 1e-300 else 1e-300
        delta = d * c
        h *= delta
        if abs(delta - 1) < 1e-12:
            break
    return h


def _betai(a, b, x):
    if x <= 0:
        return 0.0
    if x >= 1:
        return 1.0
    lbeta = math.lgamma(a + b) - math.lgamma(a) - math.lgamma(b) + a * math.log(x) + b * math.log(1 - x)
    if x < (a + 1) / (a + b + 2):
        return math.exp(lbeta) * _betacf(a, b, x) / a
    return 1 - math.exp(lbeta) * _betacf(b, a, 1 - x) / b


def t_pvalue(t, df):
    """Two-sided p-value of a t statistic (what Excel's regression output reports)."""
    return _betai(df / 2, 0.5, df / (df + t * t))


def ols(y, X, names):
    n, k = len(y), len(X[0])
    xtx = [[sum(X[r][i] * X[r][j] for r in range(n)) for j in range(k)] for i in range(k)]
    xty = [sum(X[r][i] * y[r] for r in range(n)) for i in range(k)]
    inv = _invert(xtx)
    beta = [sum(inv[i][j] * xty[j] for j in range(k)) for i in range(k)]
    fitted = [sum(beta[i] * X[r][i] for i in range(k)) for r in range(n)]
    sse = sum((y[r] - fitted[r]) ** 2 for r in range(n))
    ybar = sum(y) / n
    sst = sum((v - ybar) ** 2 for v in y)
    df = n - k
    s2 = sse / df
    coefs = []
    for i in range(k):
        se = math.sqrt(s2 * inv[i][i])
        t = beta[i] / se
        coefs.append(dict(term=names[i], estimate=beta[i], se=se, t=t, p=t_pvalue(t, df)))
    return dict(n=n, df=df, r2=1 - sse / sst, coefs=coefs)


def _invert(m):
    k = len(m)
    a = [row[:] + [1.0 if i == j else 0.0 for j in range(k)] for i, row in enumerate(m)]
    for col in range(k):
        piv = max(range(col, k), key=lambda r: abs(a[r][col]))
        a[col], a[piv] = a[piv], a[col]
        p = a[col][col]
        a[col] = [v / p for v in a[col]]
        for r in range(k):
            if r != col:
                f = a[r][col]
                a[r] = [vr - f * vc for vr, vc in zip(a[r], a[col])]
    return [row[k:] for row in a]


# ------------------------------------------------------------------ data

def merged_days(con):
    rows = con.execute("""
        SELECT s.date, SUM(s.netTotal), SUM(s.netTotal_364), w.tmax, w.prcp, COALESCE(g.gameDay, 0)
        FROM sales s JOIN weather w ON w.date = s.date LEFT JOIN games g ON g.date = s.date
        GROUP BY s.date ORDER BY s.date""").fetchall()
    days = []
    by_date = {}
    for d, sales, s364, tmax, prcp, game in rows:
        day = dt.datetime.strptime(d[:10], "%Y-%m-%d").date()
        prcp = None if prcp == "NA" else prcp  # 5 days have no precipitation reading
        rec = dict(date=day.isoformat(), sales=sales, sales_364=s364, tmax=tmax, prcp=prcp, gameDay=int(game),
                   weekend=int(day.weekday() >= 4))  # Friday–Sunday
        by_date[day] = rec
        days.append(rec)
    for rec in days:
        prev = by_date.get(dt.date.fromisoformat(rec["date"]) - dt.timedelta(days=7))
        rec["sales_lag7"] = prev["sales"] if prev else None
    return days


def write_merged(days):
    path = ROOT / "data" / "Restaurant_merged.csv"
    cols = ["date", "sales", "sales_364", "tmax", "prcp", "gameDay", "weekend", "sales_lag7"]
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(cols)
        for r in days:
            w.writerow(["" if r[c] is None else (round(r[c], 4) if isinstance(r[c], float) else r[c]) for c in cols])


# ------------------------------------------------------------------ exams

VARS = {
    "sales": ("Daily sales", "total net sales for the day ($)"),
    "tmax": ("Max temperature", "highest temperature of the day (°F)"),
    "prcp": ("Precipitation", "precipitation (inches)"),
    "gameDay": ("Game day", "1 = university football game that day, 0 = no game (reference group)"),
    "weekend": ("Weekend", "1 = Friday, Saturday or Sunday, 0 = Monday–Thursday (reference group)"),
    "sales_364": ("Sales one year ago", "sales on the same weekday 364 days earlier ($)"),
    "sales_lag7": ("Sales 7 days earlier", "sales on the same weekday one week earlier ($)"),
    "sales_364:gameDay": ("Sales one year ago × Game day", "interaction of the two variables"),
}

MODELS = [
    ("R01", "Temperature and sales", ["tmax"]),
    ("R02", "Game days and sales", ["gameDay"]),
    ("R03", "Weekend and temperature", ["weekend", "tmax"]),
    ("R04", "Weather and game days", ["tmax", "prcp", "gameDay"]),
    ("R05", "Last year's sales and game days", ["sales_364", "gameDay", "sales_364:gameDay"]),
    ("R06", "Predicting next week", ["sales_lag7"]),
]


def _value(rec, term):
    if ":" in term:
        a, b = term.split(":")
        return rec[a] * rec[b]
    return rec[term]


def _r(x, d=2):
    return round(x + 0.0, d)


def _exam(eid, title, terms, days):
    use = [d for d in days if all(d[v] is not None for t in terms for v in t.split(":"))]
    y = [d["sales"] for d in use]
    X = [[1.0] + [float(_value(d, t)) for t in terms] for d in use]
    fit = ols(y, X, ["Intercept"] + terms)
    b = [c["estimate"] for c in fit["coefs"]]
    eq = "Predicted daily sales = b₀ + " + " + ".join(f"b{i + 1}[{VARS[t][0]}]" for i, t in enumerate(terms))
    qs = []

    # 1. Prediction
    if terms == ["sales_lag7"]:
        last = use[-1]
        x = {"sales_lag7": last["sales"]}
        prompt = (f"The last day in the data is {last['date']} with sales of ${last['sales']:,.0f}. "
                  "Predict sales 7 days later.")
    else:
        x = {"tmax": 70, "prcp": 0.5, "gameDay": 1, "weekend": 1, "sales_364": 3000}
        x["sales_364:gameDay"] = x["sales_364"] * x["gameDay"]
        shown = [t for t in terms if ":" not in t]
        prompt = "Predict daily sales for a day with " + ", ".join(f"{VARS[t][0].lower()} = {x[t]:g}" for t in shown) + "."
    pred = b[0] + sum(b[i + 1] * x[t] for i, t in enumerate(terms))
    sub = " + ".join([f"{_r(b[0])}"] + [f"({_r(b[i + 1], 4)} × {x[t]:g})" for i, t in enumerate(terms)])
    qs.append(dict(id=f"{eid}-P", kind="num", prompt=prompt, answer=_r(pred), tol=max(1.0, abs(pred) * 0.005),
                   explain=f"Substitute into the model: {sub} = {_r(pred)}."))

    # 2. Coefficient interpretation (first non-interaction slope)
    i, t = 1, terms[0]
    unit = 10 if t == "tmax" else (1000 if t.startswith("sales") else 1)
    eff = b[i] * unit
    what = {10: "10 °F higher", 1000: "$1,000 higher", 1: "1 instead of 0"}[unit]
    extra = " for a day without a game (game day = 0)" if "sales_364:gameDay" in terms else ", holding the other variables constant" if len(terms) > 1 else ""
    qs.append(dict(id=f"{eid}-C", kind="num",
                   prompt=f"By how much do predicted daily sales change when {VARS[t][0].lower()} is {what}{extra}?",
                   answer=_r(eff), tol=max(0.5, abs(eff) * 0.005),
                   explain=f"The change is b{i} × {unit} = {_r(b[i], 4)} × {unit} = {_r(eff)}."))

    # 3. Significance of the last term
    j = len(terms)
    p = fit["coefs"][j]["p"]
    sig = p < 0.05
    qs.append(dict(id=f"{eid}-S", kind="choice",
                   prompt=f"At the 5% significance level, is {VARS[terms[-1]][0].lower()} statistically significant in this model?",
                   options=["Yes", "No"], answer="Yes" if sig else "No",
                   explain=f"Its p-value is {p:.3f}, which is {'below' if sig else 'above'} 0.05, so "
                           f"{'we reject' if sig else 'we cannot reject'} the hypothesis that its coefficient is zero."))

    # 4. Interaction or explanatory power
    if "sales_364:gameDay" in terms:
        k = terms.index("sales_364:gameDay") + 1
        slope_game = b[1] + b[k]
        qs.append(dict(id=f"{eid}-I", kind="num",
                       prompt="On a game day, by how much do predicted sales change for each $1 higher sales one year ago?",
                       answer=_r(slope_game, 4), tol=0.0006,
                       explain=f"On a game day the slope is b1 + b3 = {_r(b[1], 4)} + ({_r(b[k], 4)}) = {_r(slope_game, 4)}."))
    else:
        qs.append(dict(id=f"{eid}-R", kind="num",
                       prompt="What share of the variation in daily sales does this model explain? Give R² as a percentage (one decimal).",
                       answer=_r(fit["r2"] * 100, 1), tol=0.051,
                       explain=f"R² = {fit['r2']:.4f}, so the model explains {fit['r2'] * 100:.1f}% of the variation."))

    table = [dict(param=f"b{i}", term=("Intercept" if c["term"] == "Intercept" else VARS[c["term"]][0]),
                  estimate=_r(c["estimate"], 4), p=round(c["p"], 3)) for i, c in enumerate(fit["coefs"])]
    variables = [dict(name=VARS[t][0], desc=VARS[t][1]) for t in ["sales"] + [t for t in terms if ":" not in t]]
    return dict(id=eid, title=title, equation=eq, variables=variables, table=table, n=fit["n"],
                r2=round(fit["r2"], 4), terms=terms, questions=qs)


def build_regression(con):
    days = merged_days(con)
    write_merged(days)
    return dict(source="Restaurant_merged.csv", exams=[_exam(e, t, terms, days) for e, t, terms in MODELS])
