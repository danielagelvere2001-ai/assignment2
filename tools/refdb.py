"""Load every course dataset into one SQLite reference database.

The answer key is computed from this database. Excel datasets are loaded
as-is (no cleaning beyond what the question itself asks for) so that the
reference answer matches what a student gets from the downloaded file.
SQL datasets are attached unchanged from data/.
"""
import csv
import datetime as dt
import sqlite3
from pathlib import Path

import openpyxl

DATA = Path(__file__).resolve().parents[1] / "data"
CACHE = Path(__file__).resolve().parent / ".cache"


def _iso(value):
    if isinstance(value, dt.datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    return value


def _us_datetime(text):
    """'7/18/2016 6:00' -> '2016-07-18 06:00:00' (film permits are US format)."""
    if not text:
        return None
    for fmt in ("%m/%d/%Y %H:%M", "%m/%d/%Y"):
        try:
            return dt.datetime.strptime(text, fmt).strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            pass
    raise ValueError(text)


def _load_rows(con, table, header, rows):
    cols = ", ".join(f'"{h}"' for h in header)
    con.execute(f'CREATE TABLE "{table}" ({cols})')
    marks = ", ".join("?" for _ in header)
    con.executemany(f'INSERT INTO "{table}" VALUES ({marks})', rows)


def _load_sheet(con, table, path, sheet, ncols=None):
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True)[sheet]
    it = ws.iter_rows(values_only=True)
    header = [h for h in next(it) if h is not None]
    n = ncols or len(header)
    rows = [tuple(_iso(v) for v in r[:n]) for r in it if any(v not in (None, "") for v in r[:n])]
    _load_rows(con, table, header[:n], rows)


def _load_pipe_csv(con, table, path, delimiter="|"):
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f, delimiter=delimiter))
    _load_rows(con, table, rows[0], [tuple(r) for r in rows[1:]])


def build():
    CACHE.mkdir(exist_ok=True)
    path = CACHE / "ref.db"
    path.unlink(missing_ok=True)
    con = sqlite3.connect(path)

    # NYC film permits (Excel Part 1) - read the raw text export.
    with open(DATA / "NYC_Film_Permits.txt", newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    header = [h.replace("(s)", "s") for h in rows[0]]
    film = []
    for r in rows[1:]:
        r = list(r)
        for i in (2, 3, 4):
            r[i] = _us_datetime(r[i])
        film.append(tuple(r))
    _load_rows(con, "film", header, film)

    # ETL cases (Excel Part 2)
    _load_sheet(con, "codes2", DATA / "ETL_Case2.xlsx", "Case 2")
    _load_sheet(con, "codes3", DATA / "ETL_Case3.xlsx", "Case 3 data")
    for sheet, table, n in (("JELineItems", "je_lines", None), ("BusinessUnits", "business_units", None),
                            ("ChartOfAccounts", "chart_of_accounts", 4), ("PreparerInfo", "preparers", None),
                            ("Source", "sources", None)):
        _load_sheet(con, table, DATA / "ETL_Case4.xlsx", sheet, n)

    # IntegrateCo (Excel Part 3) - pipe-delimited despite the .csv extension.
    _load_pipe_csv(con, "pay15", DATA / "IntegrateCo_2015payroll.csv")
    _load_pipe_csv(con, "pay16", DATA / "IntegrateCo_2016payroll.csv")
    _load_pipe_csv(con, "jobcodes", DATA / "IntegrateCo_jobcodes.csv")
    _load_pipe_csv(con, "locations", DATA / "IntegrateCo_locations.csv")

    # Restaurant regression case (Excel Part 4)
    for sheet, table in (("salesData", "sales"), ("weatherData", "weather"), ("gameData", "games")):
        _load_sheet(con, table, DATA / "Restaurant_data.xlsm", sheet)

    con.commit()
    for alias, name in (("songs", "songs.db"), ("movies", "movies.db"), ("chinook", "chinook.sqlite"),
                        ("pcards", "pcards.db"), ("fiftyville", "fiftyville.db")):
        con.execute("ATTACH ? AS " + alias, (str(DATA / name),))
    return con


if __name__ == "__main__":
    con = build()
    for (name,) in con.execute("SELECT name FROM main.sqlite_master WHERE type='table'"):
        print(name, con.execute(f'SELECT COUNT(*) FROM "{name}"').fetchone()[0])
