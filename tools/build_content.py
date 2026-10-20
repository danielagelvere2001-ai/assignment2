"""Compute the answer key and write content/questions.json and content/exams.json.

For every question this records:
  * the reference method (Excel steps + check queries, or the SQL query),
  * the computed answer,
  * a fingerprint of method + answer. A human review in content/reviews.json is
    tied to that fingerprint, so if a question or dataset changes later the
    review automatically becomes "needs re-review" instead of vouching for an
    answer nobody has looked at,
  * the instructor's published answer where one exists, and whether the
    computed answer agrees with it.

Run from the repo root:  python tools/build_content.py
"""
import csv
import datetime as dt
import hashlib
import json
import random
import re
import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import refdb  # noqa: E402
from questions import DATASETS, Q  # noqa: E402
from regression import build_regression  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "content"
MAX_ROWS = 1000  # keep answer files small enough to load quickly

# ------------------------------------------------------------------ Python checks
# A few datasets need cleaning that is clearer in Python than in one SQL query.

def _codes3(con):
    out = []
    for (code,) in con.execute("SELECT EmployeeCode FROM codes3"):
        m = re.fullmatch(r"([A-Za-z]+)(\d+)-(\d+)([A-Za-z]+)", code.strip())
        if not m:
            raise ValueError(f"Unparsed Case 3 code: {code!r}")
        loc, emp, plant, pay = m.groups()
        loc = {"Canada": "CAN", "Japan": "JAP", "US": "USA"}.get(loc, loc)
        pay = {"W": "W", "We": "W", "M": "M", "Mo": "M"}[pay]
        out.append((loc, int(emp), int(plant), pay))
    return out


def _etl1():
    with open(ROOT / "data" / "ETL_Case1.csv", newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))[1:]
    clean = []
    for r in rows:
        if len(r) == 9:  # "Apple, Inc." split into two cells
            r = [r[0], r[1] + "," + r[2]] + r[3:]
        clean.append(r)
    return rows, clean


def _money(text):
    t = text.strip().replace(",", "")
    if t in ("", "-"):
        return 0.0
    neg = t.startswith("(") and t.endswith(")")
    t = t.strip("()")
    mult = 1
    if t.endswith(" M"):
        mult, t = 1e6, t[:-2]
    elif t.endswith(" K"):
        mult, t = 1e3, t[:-2]
    v = float(t) * mult
    return -v if neg else v


def _dedup(rows):
    seen, out = set(), []
    for r in rows:
        if tuple(r) not in seen:
            seen.add(tuple(r))
            out.append(r)
    return out


def py_check(con, name):
    key, _, arg = name.partition(":")
    if key.startswith("codes3"):
        rows = _codes3(con)
        if key == "codes3_weekly":
            return sum(1 for r in rows if r[3] == "W")
        if key == "codes3_country":
            return sum(1 for r in rows if r[0] == arg)
        if key == "codes3_distinct_lp":
            return len({f"{r[0]}-{r[2]}" for r in rows})
        if key == "codes3_empsum":
            return sum(r[1] for r in rows)
    raw, clean = _etl1()
    dd = _dedup(clean)
    if key == "etl1_rows":
        return len(raw)
    if key == "etl1_rows_dedup":
        return len(_dedup(raw))
    if key == "etl1_apple_cols":
        return max(len(r) for r in raw)
    if key == "etl1_cash_2018":
        return sum(_money(r[3]) for r in dd if r[2] == "2018")
    if key == "etl1_amzn_dep":
        return sum(_money(r[7]) for r in dd if r[1].startswith("Amazon"))
    if key == "etl1_ar_2017":
        return sum(_money(r[4]) for r in dd if r[2] == "2017")
    if key == "etl1_bad_ticker":
        return sum(1 for r in raw if not r[0].strip() or " " in r[0])
    raise KeyError(name)


# ------------------------------------------------------------------ helpers

def scalar(con, query):
    if query.startswith("PY:"):
        return py_check(con, query[3:])
    row = con.execute(query).fetchone()
    return row[0] if row else None


def norm(v):
    if isinstance(v, float):
        return round(v, 6) if abs(v - round(v)) > 1e-9 else int(round(v))
    return v


def fingerprint(obj):
    return hashlib.sha256(json.dumps(obj, sort_keys=True, default=str).encode()).hexdigest()[:16]


def agrees(computed, instructor, kind_or_none):
    """Loose agreement test used only to raise a flag for the human reviewer."""
    comp = computed if isinstance(computed, list) else [computed]
    inst = instructor if isinstance(instructor, list) else [instructor]
    if len(inst) > len(comp):
        return False
    for a, b in zip(comp, inst):
        if isinstance(b, (int, float)) and not isinstance(b, bool):
            try:
                a = float(a)
            except (TypeError, ValueError):
                return False
            if abs(a - b) > max(1e-4, abs(b) * 1e-3):
                return False
        elif str(a).strip().lower() != str(b).strip().lower():
            return False
    return True


# ------------------------------------------------------------------ build

def build_questions(con):
    out = []
    for q in Q:
        item = {k: q[k] for k in ("id", "ds", "subject", "prompt", "topics", "minutes")}
        if q["subject"] == "Excel":
            fields = []
            for ch in q["checks"]:
                v = norm(scalar(con, ch["sql"]))
                if v is None:
                    raise ValueError(f"{q['id']}: check {ch['label']!r} returned nothing")
                fields.append(dict(label=ch["label"], kind=ch["kind"], value=v, tol=ch["tol"],
                                   method=None if ch["sql"].startswith("PY:") else " ".join(ch["sql"].split())))
            item["fields"] = fields
            item["steps"] = q["steps"]
            computed = [f["value"] for f in fields]
            method = [f["method"] or "python:" + ch["sql"] for f, ch in zip(fields, q["checks"])]
        else:
            # Run against the exact file students download, opened read-only.
            cur = dataset_con(q["ds"]).execute(q["sql"])
            cols = [d[0] for d in cur.description]
            rows = [[norm(v) for v in r] for r in cur.fetchall()]
            if len(rows) > MAX_ROWS:
                raise ValueError(f"{q['id']}: {len(rows)} rows, more than {MAX_ROWS}")
            if not rows:
                raise ValueError(f"{q['id']}: empty result")
            item["sql"] = q["sql"].strip()
            item["result"] = dict(columns=cols, rows=rows)
            item["order"] = q["order"]
            computed = rows[0] if len(rows) == 1 else None
            method = q["sql"]
        item["fingerprint"] = fingerprint([q["prompt"], method, item.get("fields"), item.get("result")])
        if q.get("instructor"):
            inst = q["instructor"]
            ok = computed is not None and agrees(computed, inst["value"], None)
            item["instructor"] = dict(value=inst["value"], source=inst["source"], agrees=ok, note=inst.get("note"))
        out.append(item)
    return out


_CONS = {}


def dataset_con(ds):
    if ds not in _CONS:
        path = (ROOT / "data" / DATASETS[ds]["db"]).as_posix()
        _CONS[ds] = sqlite3.connect(f"file:{path}?mode=ro", uri=True)
    return _CONS[ds]


def build_exams(questions):
    """Data manipulation exams: two Excel questions on one file + three SQL questions on one database."""
    rng = random.Random(2026)
    by_ds = {}
    for q in questions:
        by_ds.setdefault(q["ds"], []).append(q)
    excel_ds = [d for d in by_ds if DATASETS[d]["subject"] == "Excel" and len(by_ds[d]) >= 2]
    sql_ds = [d for d in by_ds if DATASETS[d]["subject"] == "SQL"]
    exams, seen = [], set()
    while len(exams) < 60:
        e, s = rng.choice(excel_ds), rng.choice(sql_ds)
        xs = rng.sample(by_ds[e], 2)
        ss = rng.sample(by_ds[s], 3)
        ids = tuple(sorted(q["id"] for q in xs + ss))
        if ids in seen:
            continue
        seen.add(ids)
        exams.append(dict(id=f"E{len(exams) + 1:02d}", datasets=[e, s], questions=[q["id"] for q in xs + ss],
                          minutes=8))
    return exams


def main():
    OUT.mkdir(exist_ok=True)
    con = refdb.build()
    questions = build_questions(con)
    exams = build_exams(questions)
    regression = build_regression(con)
    for exam in regression["exams"]:
        for rq in exam["questions"]:
            rq["fingerprint"] = fingerprint([exam["table"], exam["n"], rq["prompt"], rq["answer"]])
    built = dt.date.today().isoformat()
    json_dump(OUT / "questions.json", dict(built=built, datasets=DATASETS, questions=questions))
    json_dump(OUT / "exams.json", dict(built=built, exams=exams))
    json_dump(OUT / "regression.json", dict(built=built, **regression))
    reviews = OUT / "reviews.json"
    if not reviews.exists():
        json_dump(reviews, dict(reviews={}))

    flags = [q for q in questions if q.get("instructor") and not q["instructor"]["agrees"]]
    print(f"{len(questions)} questions, {len(exams)} exams, {len(regression['exams'])} regression exams")
    print(f"{sum(1 for q in questions if q.get('instructor'))} with an instructor answer, {len(flags)} disagree:")
    for q in flags:
        got = [f["value"] for f in q["fields"]] if "fields" in q else q["result"]["rows"][:3]
        print(f"  {q['id']}: instructor {q['instructor']['value']!r} vs computed {got!r}")


def json_dump(path, obj):
    path.write_text(json.dumps(obj, indent=1, ensure_ascii=False, default=str) + "\n", encoding="utf-8")


if __name__ == "__main__":
    main()
