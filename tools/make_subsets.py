"""Build the downloadable practice datasets from the course materials.

The two large SQLite files are cut down so they load in a browser:
  movies.db  -> movies with at least 5,000 IMDb votes, plus their stars,
                directors and people (both Kevin Bacons are kept on purpose).
  pcards.db  -> calendar years 2013 and 2014 only (the years the P-card
                case tests), all columns kept.
Every other file is copied unchanged.

Run from the repo root:  python tools/make_subsets.py
"""
import gzip
import shutil
import sqlite3
from pathlib import Path

MATERIALS = Path(__file__).resolve().parents[2] / "materials"
OUT = Path(__file__).resolve().parents[1] / "data"

COPIES = {
    "1. Excel Part 1/PWC Excel Assignment/NYC_Film_Permits.txt": "NYC_Film_Permits.txt",
    "1. Excel Part 1/PWC Excel Assignment/NYC_Film_Permits.xlsx": "NYC_Film_Permits.xlsx",
    "1. Excel part 2/ETL Assignment/data/case 1/Analytics_mindset_case_studies_ETL_Case1.csv": "ETL_Case1.csv",
    "1. Excel part 2/ETL Assignment/data/case 2/Analytics_mindset_case_studies_ETL_Case2_Excel.xlsx": "ETL_Case2.xlsx",
    "1. Excel part 2/ETL Assignment/data/case 3/Analytics_mindset_case_studies_ETL_Case3_Excel.xlsx": "ETL_Case3.xlsx",
    "1. Excel part 2/ETL Assignment/data/case 4/Analytics_mindset_case_studies_ETL_Case4_Excel.xlsx": "ETL_Case4.xlsx",
    "1. Excel part 3/IntegrateCo Assignment/Analytics_mindset_case_studies_IntegrateCo_2015payroll.csv": "IntegrateCo_2015payroll.csv",
    "1. Excel part 3/IntegrateCo Assignment/Analytics_mindset_case_studies_IntegrateCo_2016payroll.csv": "IntegrateCo_2016payroll.csv",
    "1. Excel part 3/IntegrateCo Assignment/Analytics_mindset_case_studies_IntegrateCo_jobcodes.csv": "IntegrateCo_jobcodes.csv",
    "1. Excel part 3/IntegrateCo Assignment/Analytics_mindset_case_studies_IntegrateCo_locations.csv": "IntegrateCo_locations.csv",
    "1. Excel part 4/Regression Assignment/Restaurant/data.xlsm": "Restaurant_data.xlsm",
    "2. SQL part 1/Assignments/cs50 assignment/songs/songs.db": "songs.db",
    "2. SQL part 1/Assignments/pwc assignment/Chinook_Sqlite.sqlite": "chinook.sqlite",
    "2. SQL part 1/Assignments/cs50 assignment/mystery (OPTIONAL)/fiftyville.db": "fiftyville.db",
}


def subset_movies(src_path, dst_path):
    dst_path.unlink(missing_ok=True)
    con = sqlite3.connect(dst_path)
    con.execute("ATTACH ? AS src", (str(src_path),))
    for (sql,) in con.execute("SELECT sql FROM src.sqlite_master WHERE type='table'").fetchall():
        con.execute(sql)
    con.executescript("""
        CREATE TEMP TABLE keep AS SELECT movie_id AS id FROM src.ratings WHERE votes >= 5000;
        INSERT INTO movies SELECT * FROM src.movies WHERE id IN (SELECT id FROM keep);
        INSERT INTO ratings SELECT * FROM src.ratings WHERE movie_id IN (SELECT id FROM keep);
        INSERT INTO stars SELECT * FROM src.stars WHERE movie_id IN (SELECT id FROM keep);
        INSERT INTO directors SELECT * FROM src.directors WHERE movie_id IN (SELECT id FROM keep);
        INSERT INTO people SELECT * FROM src.people
            WHERE id IN (SELECT person_id FROM stars UNION SELECT person_id FROM directors)
               OR name = 'Kevin Bacon';
        CREATE INDEX stars_movie ON stars(movie_id);
        CREATE INDEX stars_person ON stars(person_id);
        CREATE INDEX directors_movie ON directors(movie_id);
        CREATE INDEX directors_person ON directors(person_id);
        CREATE INDEX ratings_movie ON ratings(movie_id);
    """)
    con.commit()
    con.execute("DETACH src")
    con.execute("VACUUM")
    con.close()


def subset_pcards(src_path, dst_path):
    dst_path.unlink(missing_ok=True)
    con = sqlite3.connect(dst_path)
    con.execute("ATTACH ? AS src", (str(src_path),))
    (sql,) = con.execute("SELECT sql FROM src.sqlite_master WHERE name='pcards'").fetchone()
    con.execute(sql)
    con.execute("INSERT INTO pcards SELECT * FROM src.pcards WHERE Year IN (2013, 2014)")
    con.commit()
    con.execute("DETACH src")
    con.execute("VACUUM")
    con.close()


def gzip_copy(path):
    with open(path, "rb") as f, gzip.open(str(path) + ".gz", "wb", compresslevel=9) as g:
        shutil.copyfileobj(f, g)


def main():
    OUT.mkdir(exist_ok=True)
    for src, name in COPIES.items():
        shutil.copyfile(MATERIALS / src, OUT / name)
    subset_movies(MATERIALS / "2. SQL part 1/Assignments/cs50 assignment/movies/movies.db", OUT / "movies.db")
    subset_pcards(MATERIALS / "2. SQL part 2/Pcards Assignment/pcards.db", OUT / "pcards.db")
    # Compressed copies are what the in-browser SQL runner loads.
    for name in ("songs.db", "chinook.sqlite", "movies.db", "pcards.db", "fiftyville.db"):
        gzip_copy(OUT / name)
    for p in sorted(OUT.iterdir()):
        print(f"{p.stat().st_size / 1e6:8.2f} MB  {p.name}")


if __name__ == "__main__":
    main()
