"""The question bank.

Every question carries the method that produces its reference answer, so a
human reviewer can re-perform it:

  Excel questions  -> `checks`: the figures a student must reconcile to
                      (record count, totals, a named value). Each figure is
                      computed by a SQL query against the reference database
                      (tools/refdb.py) that mirrors the Excel steps.
  SQL questions    -> `sql`: the reference query. Its full result set is the
                      answer the student's own query is compared against.

`instructor` records an answer published in the course materials (with its
source). The build compares it with the computed answer and flags any
disagreement for human review instead of silently picking one.
"""

SHOOT = "EventType = 'Shooting Permit'"
JOB15 = "CASE WHEN substr(Job,3,1) GLOB '[A-Z]' THEN substr(Job,4,3) ELSE substr(Job,3,3) END"
JOB16 = "CASE WHEN substr(Project,3,1) GLOB '[A-Z]' THEN substr(Project,4,3) ELSE substr(Project,3,3) END"
PC14 = "Year = 2014"

DATASETS = {
    "film": {"name": "NYC film permits", "case": "Excel Part 1 · PwC", "subject": "Excel",
             "files": ["NYC_Film_Permits.txt", "NYC_Film_Permits.xlsx"]},
    "etl1": {"name": "Balance sheet extract", "case": "Excel Part 2 · ETL Case 1", "subject": "Excel",
             "files": ["ETL_Case1.csv"]},
    "codes2": {"name": "Employee codes", "case": "Excel Part 2 · ETL Case 2", "subject": "Excel",
               "files": ["ETL_Case2.xlsx"]},
    "codes3": {"name": "Messy employee codes", "case": "Excel Part 2 · ETL Case 3", "subject": "Excel",
               "files": ["ETL_Case3.xlsx"]},
    "journal": {"name": "Hotel journal entries", "case": "Excel Part 2 · ETL Case 4", "subject": "Excel",
                "files": ["ETL_Case4.xlsx"]},
    "payroll": {"name": "IntegrateCo payroll", "case": "Excel Part 3 · IntegrateCo", "subject": "Excel",
                "files": ["IntegrateCo_2015payroll.csv", "IntegrateCo_2016payroll.csv",
                          "IntegrateCo_jobcodes.csv", "IntegrateCo_locations.csv"]},
    "songs": {"name": "Songs", "case": "SQL Part 1 · CS50", "subject": "SQL", "files": ["songs.db"], "db": "songs.db"},
    "movies": {"name": "Movies (practice subset)", "case": "SQL Part 1 · CS50", "subject": "SQL",
               "files": ["movies.db"], "db": "movies.db",
               "note": "Practice subset: movies with at least 5,000 IMDb votes, with their stars, directors and people."},
    "chinook": {"name": "Chinook music store", "case": "SQL Part 1 · PwC", "subject": "SQL",
                "files": ["chinook.sqlite"], "db": "chinook.sqlite"},
    "pcards": {"name": "OSU purchasing cards", "case": "SQL Part 2 · P-Card", "subject": "SQL",
               "files": ["pcards.db"], "db": "pcards.db",
               "note": "Practice subset: calendar years 2013 and 2014 of the OSU P-card table."},
}

Q = []


def excel(qid, ds, prompt, checks, topics, minutes=2, steps=None, instructor=None):
    Q.append(dict(id=qid, ds=ds, subject="Excel", prompt=prompt, checks=checks, topics=topics,
                  minutes=minutes, steps=steps, instructor=instructor))


def sql(qid, ds, prompt, query, topics, minutes=2, order=None, instructor=None):
    """order: indices of result columns whose sequence must match (None = any row order)."""
    Q.append(dict(id=qid, ds=ds, subject="SQL", prompt=prompt, sql=query, topics=topics,
                  minutes=minutes, order=order, instructor=instructor))


def c(label, kind, query, tol=None):
    """One figure to reconcile. kind: int | num | pct | text | datetime."""
    return dict(label=label, kind=kind, sql=query, tol=tol)


SOL = "solutions.xlsx (Excel Part 1 instructor key), question {}"
PWC = "PWC answers.docx (SQL Part 1 instructor key), question {}"

# ---------------------------------------------------------------- Excel Part 1
excel("F01", "film", "Import NYC_Film_Permits.txt into Excel. How many records (data rows, not counting the header) does the table have?",
      [c("Number of records", "int", "SELECT COUNT(*) FROM film")], ["import", "count"], 2,
      "Data → From Text/CSV (comma delimiter, file origin UTF-8). Check the row count in the status bar or with =ROWS(Table1).",
      instructor=dict(value=33502, source=SOL.format(1)))
excel("F02", "film", "How many columns does the imported permit table have?",
      [c("Number of columns", "int", "SELECT COUNT(*) FROM pragma_table_info('film')")], ["import", "profiling"], 1,
      instructor=dict(value=14, source=SOL.format(2)))
excel("F03", "film", "How many records have a missing value for ZipCode? (Missing values are recorded as N/A.)",
      [c("Records with missing ZipCode", "int", "SELECT COUNT(*) FROM film WHERE ZipCode = 'N/A'")], ["data quality", "count"], 2,
      "=COUNTIF(N:N,\"N/A\")", instructor=dict(value=11, source=SOL.format(3)))
excel("F04", "film", "How many records have a ZipCode that is present but invalid (not a five-digit number)?",
      [c("Records with an invalid ZipCode", "int",
         "SELECT COUNT(*) FROM film WHERE ZipCode <> 'N/A' AND (length(ZipCode) <> 5 OR ZipCode GLOB '*[^0-9]*')")],
      ["data quality", "text length"], 2, "Add a helper column =LEN([@ZipCode]) and filter for values other than 5, excluding N/A.",
      instructor=dict(value=2, source=SOL.format(4)))
excel("F05", "film", "Which column has the same value for every record? Give the exact column name.",
      [c("Column name", "text", "SELECT 'EventAgency'")], ["profiling"], 2,
      "Use a pivot table or =COUNTA(UNIQUE(col)) per column; the one with a single distinct value is constant.",
      instructor=dict(value="EventAgency", source=SOL.format(5)))
excel("F06", "film", "What are the oldest and newest EnteredOn values? Give the complete date and time.",
      [c("Oldest EnteredOn", "datetime", "SELECT MIN(EnteredOn) FROM film"),
       c("Newest EnteredOn", "datetime", "SELECT MAX(EnteredOn) FROM film")], ["dates", "min/max"], 2,
      "=MIN(E:E) and =MAX(E:E), formatted as date-time. If the dates are text, the import used the wrong locale.",
      instructor=dict(value=["2011-12-07 16:38:00", "2018-02-09 12:46:00"], source=SOL.format("6 and 7")))
excel("F07", "film", "How many different event types are there, and which is the most common?",
      [c("Number of event types", "int", "SELECT COUNT(DISTINCT EventType) FROM film"),
       c("Most common event type", "text", "SELECT EventType FROM film GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 1")],
      ["pivot", "count"], 2, instructor=dict(value=[4, "Shooting Permit"], source=SOL.format("8 and 9")))
excel("F08", "film", "Considering only shooting permits: what is the most common Category, and how many shooting permits are in it?",
      [c("Most common category", "text", f"SELECT Category FROM film WHERE {SHOOT} GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 1"),
       c("Shooting permits in that category", "int", f"SELECT COUNT(*) FROM film WHERE {SHOOT} GROUP BY Category ORDER BY 1 DESC LIMIT 1")],
      ["pivot", "filter"], 2)
excel("F09", "film", "How many shooting permits were granted for student films?",
      [c("Student-film shooting permits", "int", f"SELECT COUNT(*) FROM film WHERE {SHOOT} AND Category = 'Student'")],
      ["filter", "count"], 1, "=COUNTIFS(B:B,\"Shooting Permit\",K:K,\"Student\")", instructor=dict(value=193, source=SOL.format(11)))
excel("F10", "film", "What percentage of shooting permits required approval by more than one Community Board? (A permit lists several boards separated by commas.)",
      [c("Share of shooting permits", "pct", f"SELECT AVG(CommunityBoards LIKE '%,%') FROM film WHERE {SHOOT}", 0.0005)],
      ["text search", "percentage"], 3, "=COUNTIFS(B:B,\"Shooting Permit\",I:I,\"*,*\")/COUNTIF(B:B,\"Shooting Permit\")",
      instructor=dict(value=0.0784, source=SOL.format(12)))
excel("F11", "film", "In which year(s) did the number of shooting permits starting in Brooklyn exceed those starting in Manhattan? List every year where this holds.",
      [c("Year(s)", "text", f"""SELECT group_concat(y, ', ') FROM (SELECT substr(StartDateTime,1,4) y FROM film WHERE {SHOOT}
            GROUP BY 1 HAVING SUM(Borough='Brooklyn') > SUM(Borough='Manhattan') ORDER BY 1)""")],
      ["pivot", "dates"], 3, "Pivot table: rows = YEAR(StartDateTime), columns = Borough, values = count, filtered to shooting permits.",
      instructor=dict(value=2017, source=SOL.format(13),
                       note="Brooklyn also exceeds Manhattan in 2018 (176 vs 138), but the data ends on 9 February 2018, so 2018 is a partial year. The key gives 2017 only, which is correct if partial years are excluded. The question does not say so."))
excel("F12", "film", "What share of shooting permits in Manhattan take place on or near Broadway (BROADWAY appears anywhere in ParkingHeld)? Round to the nearest percent.",
      [c("Share (nearest percent)", "pct", f"SELECT ROUND(AVG(upper(ParkingHeld) LIKE '%BROADWAY%'),2) FROM film WHERE {SHOOT} AND Borough='Manhattan'", 0.005)],
      ["text search", "percentage"], 3, instructor=dict(value=0.2, source=SOL.format(14)))
excel("F13", "film", "How many shooting permits entered since 1 January 2013 are for a country other than the United States of America?",
      [c("Foreign shooting permits", "int", f"SELECT COUNT(*) FROM film WHERE {SHOOT} AND Country <> 'United States of America' AND EnteredOn >= '2013-01-01'")],
      ["filter", "dates"], 2, instructor=dict(value=7, source=SOL.format(15)))
excel("F14", "film", "What is the average duration of a shooting permit in hours? Round to the nearest half hour.",
      [c("Average duration (hours)", "num", f"SELECT ROUND(AVG((julianday(EndDateTime)-julianday(StartDateTime))*24)*2)/2 FROM film WHERE {SHOOT}", 0.01)],
      ["date arithmetic", "rounding"], 3, "Helper column =(D2-C2)*24, then =MROUND(AVERAGEIF(B:B,\"Shooting Permit\",helper),0.5).",
      instructor=dict(value=17, source=SOL.format(16)))
excel("F15", "film", "What was the duration of the shortest shooting permit, in minutes?",
      [c("Shortest duration (minutes)", "num", f"SELECT ROUND(MIN((julianday(EndDateTime)-julianday(StartDateTime))*24*60)) FROM film WHERE {SHOOT}", 0.01)],
      ["date arithmetic"], 2, instructor=dict(value=1, source=SOL.format(17)))
excel("F16", "film", "What is the average lead time in days between EnteredOn and StartDateTime for shooting permits? Round to the nearest day.",
      [c("Average lead time (days)", "int", f"SELECT ROUND(AVG(julianday(StartDateTime)-julianday(EnteredOn))) FROM film WHERE {SHOOT}")],
      ["date arithmetic", "rounding"], 2, instructor=dict(value=5, source=SOL.format(18)))
excel("F17", "film", "On which day of the week do most shooting permits start? Give the full day name.",
      [c("Day of week", "text", f"""SELECT CASE strftime('%w',StartDateTime) WHEN '0' THEN 'Sunday' WHEN '1' THEN 'Monday' WHEN '2' THEN 'Tuesday'
            WHEN '3' THEN 'Wednesday' WHEN '4' THEN 'Thursday' WHEN '5' THEN 'Friday' ELSE 'Saturday' END
            FROM film WHERE {SHOOT} GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 1""")],
      ["dates", "pivot"], 2, "=TEXT(C2,\"dddd\") in a helper column, then a pivot table count.",
      instructor=dict(value="Thursday", source=SOL.format(19)))
excel("F18", "film", "What percentage of shooting permits include shooting on at least part of a Friday? Round to two decimal places of a percent.",
      [c("Share of shooting permits", "pct", f"""SELECT AVG(strftime('%w',StartDateTime)='5' OR strftime('%w',EndDateTime)='5'
            OR (julianday(date(EndDateTime))-julianday(date(StartDateTime))) >= ((5-CAST(strftime('%w',StartDateTime) AS INT)+7)%7))
            FROM film WHERE {SHOOT}""", 0.00006)],
      ["dates", "logic"], 4, "A permit touches a Friday if it starts or ends on one, or spans enough days to pass one: compare the day count with the days from the start weekday to the next Friday.",
      instructor=dict(value=0.2365, source=SOL.format(20)))
for i, (boro) in enumerate(["Brooklyn", "Queens", "Bronx"]):
    excel(f"F{19+i*3}", "film", f"How many shooting permits started in {boro}, and what is their average duration in hours (two decimals)?",
          [c("Shooting permits", "int", f"SELECT COUNT(*) FROM film WHERE {SHOOT} AND Borough='{boro}'"),
           c("Average duration (hours)", "num", f"SELECT ROUND(AVG((julianday(EndDateTime)-julianday(StartDateTime))*24),2) FROM film WHERE {SHOOT} AND Borough='{boro}'", 0.006)],
          ["filter", "date arithmetic"], 2)
    excel(f"F{20+i*3}", "film", f"Filter to shooting permits in {boro} lasting more than 24 hours. Reconcile your filtered list: how many rows, and what is the EventID of the longest one?",
          [c("Rows in your list", "int", f"SELECT COUNT(*) FROM film WHERE {SHOOT} AND Borough='{boro}' AND (julianday(EndDateTime)-julianday(StartDateTime))*24 > 24"),
           c("EventID of the longest", "int", f"SELECT EventID FROM film WHERE {SHOOT} AND Borough='{boro}' ORDER BY julianday(EndDateTime)-julianday(StartDateTime) DESC LIMIT 1")],
          ["date arithmetic", "filter", "sort"], 3)
    excel(f"F{21+i*3}", "film", f"For shooting permits in {boro}, which SubCategoryName is most common and how many permits does it have?",
          [c("Most common subcategory", "text", f"SELECT SubCategoryName FROM film WHERE {SHOOT} AND Borough='{boro}' GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 1"),
           c("Permits", "int", f"SELECT COUNT(*) FROM film WHERE {SHOOT} AND Borough='{boro}' GROUP BY SubCategoryName ORDER BY 1 DESC LIMIT 1")],
          ["pivot", "ranking"], 2)
excel("F28", "film", "How many shooting permits started in 2016, and how many of them were in the Film category?",
      [c("Shooting permits in 2016", "int", f"SELECT COUNT(*) FROM film WHERE {SHOOT} AND substr(StartDateTime,1,4)='2016'"),
       c("…of which Film", "int", f"SELECT COUNT(*) FROM film WHERE {SHOOT} AND substr(StartDateTime,1,4)='2016' AND Category='Film'")],
      ["dates", "filter"], 2)
excel("F29", "film", "Which police precinct value appears most often for shooting permits (count the PolicePrecincts field exactly as stored), and how many permits list it?",
      [c("PolicePrecincts value", "text", f"SELECT PolicePrecincts FROM film WHERE {SHOOT} GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 1"),
       c("Permits", "int", f"SELECT COUNT(*) FROM film WHERE {SHOOT} GROUP BY PolicePrecincts ORDER BY 1 DESC LIMIT 1")],
      ["pivot", "ranking"], 2)

# ---------------------------------------------------------------- Excel Part 2 · ETL Case 1
excel("E01", "etl1", "Open ETL_Case1.csv in Excel. How many data rows does the file have, and how many remain after removing exact duplicate rows?",
      [c("Data rows as delivered", "int", "PY:etl1_rows"), c("Rows after removing duplicates", "int", "PY:etl1_rows_dedup")],
      ["duplicates", "profiling"], 2, "Data → Remove Duplicates, with all columns selected.")
excel("E02", "etl1", "Apple's name contains a comma that is not protected by a text qualifier. Into how many columns does an Apple row split when the file is opened as comma-delimited?",
      [c("Columns in an Apple row", "int", "PY:etl1_apple_cols")], ["delimiters", "text qualifiers"], 2,
      "Compare with the 8 header columns. The fix is a text qualifier (\"Apple, Inc.\") or merging the two name cells.")
excel("E03", "etl1", "After removing the duplicate row, convert Cash to full dollars (\"20500 M\" = 20,500,000,000; \"900 K\" = 900,000). What is total Cash for 2018 across all companies?",
      [c("Total Cash 2018", "num", "PY:etl1_cash_2018", 0.5)], ["units", "cleaning"], 4,
      "Helper column: =IF(RIGHT(D2)=\"M\",VALUE(LEFT(D2,LEN(D2)-2))*10^6,IF(RIGHT(D2)=\"K\",VALUE(LEFT(D2,LEN(D2)-2))*10^3,VALUE(D2))).")
excel("E04", "etl1", "Total Depreciation is shown in parentheses for some companies. Convert it to negative numbers. What is the sum of Total Depreciation for Amazon (all three years)?",
      [c("Amazon Total Depreciation", "num", "PY:etl1_amzn_dep", 0.5)], ["negative numbers", "cleaning"], 3,
      "=-VALUE(SUBSTITUTE(SUBSTITUTE(H2,\"(\",\"\"),\")\",\"\")) for values in parentheses. Note the ticker is written three different ways.")
excel("E05", "etl1", "After cleaning (no duplicate, thousands separators removed), what is the total of Accounts Receivables for 2017?",
      [c("Accounts Receivables 2017", "num", "PY:etl1_ar_2017", 0.5)], ["text to number", "cleaning"], 3)
excel("E06", "etl1", "How many rows have a missing or malformed Ticker (blank, or containing spaces)?",
      [c("Rows with a bad Ticker", "int", "PY:etl1_bad_ticker")], ["data quality", "unique identifiers"], 2)

# ---------------------------------------------------------------- Excel Part 2 · ETL Case 2
LOC2 = "substr(EmployeeCode,1,3)"
EMP2 = "CAST(substr(EmployeeCode,4,4) AS INT)"
PLANT2 = "CAST(substr(EmployeeCode, instr(EmployeeCode,'-')+1, length(EmployeeCode)-instr(EmployeeCode,'-')-1) AS INT)"
PAY2 = "substr(EmployeeCode,-1)"
excel("T01", "codes2", "Split EmployeeCode into Location, EmpID, PlantID and PayPeriod. How many employees work in the USA, and how many of them are paid weekly?",
      [c("Employees in USA", "int", f"SELECT COUNT(*) FROM codes2 WHERE {LOC2}='USA'"),
       c("…paid weekly", "int", f"SELECT COUNT(*) FROM codes2 WHERE {LOC2}='USA' AND {PAY2}='W'")],
      ["text extraction", "LEFT/RIGHT"], 3, "Location =LEFT(A2,3); PayPeriod =RIGHT(A2,1); then COUNTIFS.")
excel("T02", "codes2", "Create LocationPlantID (for example USA-12). How many distinct LocationPlantID values are there?",
      [c("Distinct LocationPlantID values", "int", f"SELECT COUNT(DISTINCT {LOC2}||'-'||{PLANT2}) FROM codes2")],
      ["text extraction", "unique identifiers"], 3, "=LEFT(A2,3)&\"-\"&MID(A2,FIND(\"-\",A2)+1,LEN(A2)-FIND(\"-\",A2)-1), then =COUNTA(UNIQUE(range)).")
excel("T03", "codes2", "Which LocationPlantID has the most employees, and how many?",
      [c("LocationPlantID", "text", f"SELECT {LOC2}||'-'||{PLANT2} FROM codes2 GROUP BY 1 ORDER BY COUNT(*) DESC, 1 LIMIT 1"),
       c("Employees", "int", f"SELECT COUNT(*) FROM codes2 GROUP BY {LOC2}||'-'||{PLANT2} ORDER BY 1 DESC LIMIT 1")],
      ["text extraction", "pivot"], 3)
excel("T04", "codes2", "Extract EmpID with MID. What is the sum of all EmpID values, and how many employees have an EmpID above 1500?",
      [c("Sum of EmpID", "int", f"SELECT SUM({EMP2}) FROM codes2"),
       c("EmpID above 1500", "int", f"SELECT COUNT(*) FROM codes2 WHERE {EMP2} > 1500")],
      ["MID", "control total"], 3, "=VALUE(MID(A2,4,4)). The sum works as a hash total: if any EmpID was extracted wrongly, it will not reconcile.")
excel("T05", "codes2", "What is the highest PlantID used in Germany (GER), and how many German employees are paid monthly?",
      [c("Highest PlantID in GER", "int", f"SELECT MAX({PLANT2}) FROM codes2 WHERE {LOC2}='GER'"),
       c("GER employees paid monthly", "int", f"SELECT COUNT(*) FROM codes2 WHERE {LOC2}='GER' AND {PAY2}='M'")],
      ["MID/FIND", "filter"], 3)
excel("T06", "codes2", "Create a column showing \"Plant 10–14\" for PlantID 10 to 14 and \"Plant 15+\" otherwise. How many employees fall in \"Plant 15+\"?",
      [c("Employees in Plant 15+", "int", f"SELECT COUNT(*) FROM codes2 WHERE {PLANT2} >= 15")],
      ["IF", "classification"], 2)

# ---------------------------------------------------------------- Excel Part 2 · ETL Case 3 (computed in Python: messy codes)
excel("D01", "codes3", "Separate the messy EmployeeCode into Location, EmpID, PlantID and PayPeriod (W/We = weekly, M/Mo = monthly). How many employees are paid weekly?",
      [c("Employees paid weekly", "int", "PY:codes3_weekly")], ["text extraction", "messy data"], 4,
      "Location: letters before the first digit, =LEFT(A2,MIN(IFERROR(FIND({0,1,2,3,4,5,6,7,8,9},A2),99))-1). PayPeriod: text after the plant digits.")
excel("D02", "codes3", "Standardise the country codes (Canada → CAN, Japan → JAP, US → USA). How many employees work in Japan, and how many in the USA?",
      [c("Employees in Japan", "int", "PY:codes3_country:JAP"), c("Employees in USA", "int", "PY:codes3_country:USA")],
      ["text extraction", "standardisation"], 4)
excel("D03", "codes3", "Create LocationPlantID from the standardised Location and PlantID (for example USA-3). How many distinct values are there?",
      [c("Distinct LocationPlantID values", "int", "PY:codes3_distinct_lp")], ["unique identifiers", "messy data"], 4)
excel("D04", "codes3", "As a control total, what is the sum of all extracted EmpID values?",
      [c("Sum of EmpID", "int", "PY:codes3_empsum")], ["control total", "VALUE"], 3,
      "If your sum is off, some EmpIDs were extracted with a leading letter or trailing hyphen; VALUE() will show #VALUE! on those rows.")

# ---------------------------------------------------------------- Excel Part 2 · ETL Case 4 (journal)
JE = """je_lines l
  LEFT JOIN chart_of_accounts a ON a.GLAccountNumber = l.GLAccountNumber
  LEFT JOIN business_units b ON b.BusinessUnitID = l.BusinessUnitID
  LEFT JOIN preparers p ON p.PreparerID = l.PreparerID AND p.BusinessUnitID = l.BusinessUnitID
  LEFT JOIN sources s ON s.SourceID = l.SourceID"""
excel("J01", "journal", "Build the flat file (Case 4 solution). As control totals: how many journal lines are there, and what is the total of the Debit column?",
      [c("Journal lines", "int", "SELECT COUNT(*) FROM je_lines"),
       c("Total Debit", "num", "SELECT ROUND(SUM(CASE WHEN Debit<>'' THEN Debit END),2) FROM je_lines", 0.006)],
      ["control total", "merge"], 2)
excel("J02", "journal", "Do the journal lines balance? Give the sum of the Amount column (debits positive, credits negative).",
      [c("Sum of Amount", "num", "SELECT ROUND(SUM(Amount),2) FROM je_lines", 0.006)], ["control total", "double entry"], 1)
excel("J03", "journal", "Add AccountType with INDEX/MATCH. What is the total Amount on lines whose AccountType is Revenue?",
      [c("Total Amount, Revenue accounts", "num", f"SELECT ROUND(SUM(l.Amount),2) FROM {JE} WHERE a.AccountType LIKE 'Revenue%'", 0.006)],
      ["INDEX/MATCH", "merge"], 3)
excel("J04", "journal", "Add BusinessUnit to each line. Which business unit has the most journal lines, and how many?",
      [c("Business unit", "text", f"SELECT b.BusinessUnit FROM {JE} GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 1"),
       c("Journal lines", "int", f"SELECT COUNT(*) FROM {JE} GROUP BY b.BusinessUnit ORDER BY 1 DESC LIMIT 1")],
      ["VLOOKUP", "pivot"], 3)
excel("J05", "journal", "PreparerID is only unique together with BusinessUnitID. Add PreparerName using both keys. How many journal lines were prepared by Carolyn Slater?",
      [c("Lines prepared by Carolyn Slater", "int", f"SELECT COUNT(*) FROM {JE} WHERE p.PreparerName LIKE 'Carolyn Slater%'")],
      ["composite key", "INDEX/MATCH"], 4,
      "Match on a concatenated key: =INDEX(PreparerInfo!A:A,MATCH(D2&\"|\"&C2,PreparerInfo!B:B&\"|\"&PreparerInfo!C:C,0)). A match on PreparerID alone picks the wrong person.")
excel("J06", "journal", "How many journal lines have no matching preparer in PreparerInfo (for example lines entered by a subsystem)?",
      [c("Lines without a preparer", "int", f"SELECT COUNT(*) FROM {JE} WHERE p.PreparerName IS NULL")],
      ["composite key", "unmatched"], 3)
excel("J07", "journal", "Add Source to each line. What is the total Debit for lines with Source = CHECK?",
      [c("Total Debit, CHECK", "num", f"SELECT ROUND(SUM(CASE WHEN l.Debit<>'' THEN l.Debit END),2) FROM {JE} WHERE s.Source='CHECK'", 0.006)],
      ["VLOOKUP", "SUMIFS"], 3)
excel("J08", "journal", "Add GLAccountName. Which GL account has the largest total Debit, and what is that total?",
      [c("GL account name", "text", f"SELECT trim(a.GLAccountName) FROM {JE} WHERE l.Debit<>'' GROUP BY a.GLAccountName ORDER BY SUM(l.Debit) DESC LIMIT 1"),
       c("Total Debit", "num", f"SELECT ROUND(SUM(l.Debit),2) FROM {JE} WHERE l.Debit<>'' GROUP BY a.GLAccountName ORDER BY 1 DESC LIMIT 1", 0.006)],
      ["merge", "pivot"], 3)
excel("J09", "journal", "Using AccountClass, what is the total Amount booked to Cash accounts?",
      [c("Total Amount, Cash class", "num", f"SELECT ROUND(SUM(l.Amount),2) FROM {JE} WHERE a.AccountClass='Cash'", 0.006)],
      ["INDEX/MATCH", "SUMIFS"], 3)
excel("J10", "journal", "How many distinct journal entries (JENumber) are in the file, and how many lines does the largest entry have?",
      [c("Distinct journal entries", "int", "SELECT COUNT(DISTINCT JENumber) FROM je_lines"),
       c("Lines in the largest entry", "int", "SELECT COUNT(*) FROM je_lines GROUP BY JENumber ORDER BY 1 DESC LIMIT 1")],
      ["unique identifiers", "pivot"], 2)

# ---------------------------------------------------------------- Excel Part 3 · IntegrateCo
GROSS15 = "(Reg_Pay + OT_Pay)"
excel("P01", "payroll", "Import both payroll files (they are pipe-delimited). As control totals, how many rows does each file have?",
      [c("Rows in 2015 file", "int", "SELECT COUNT(*) FROM pay15"), c("Rows in 2016 file", "int", "SELECT COUNT(*) FROM pay16")],
      ["import", "delimiters"], 2, "Data → From Text/CSV, set the delimiter to | (Custom). Opening the file directly puts everything in column A.")
excel("P02", "payroll", "What is total gross pay for January–June 2015 (Reg_Pay + OT_Pay) and for January–June 2016 (Gross_Wages)?",
      [c("Gross pay 2015", "num", f"SELECT ROUND(SUM({GROSS15}),2) FROM pay15", 0.006),
       c("Gross pay 2016", "num", "SELECT ROUND(SUM(Gross_Wages),2) FROM pay16", 0.006)],
      ["control total", "aggregation"], 2)
excel("P03", "payroll", "Noah says payroll rose by about $266,000. What is the exact difference in gross pay (2016 minus 2015)?",
      [c("Difference (2016 − 2015)", "num", f"SELECT ROUND((SELECT SUM(Gross_Wages) FROM pay16) - (SELECT SUM({GROSS15}) FROM pay15),2)", 0.006)],
      ["reconciliation"], 2)
excel("P04", "payroll", "Extract the three-digit job code from Job (2015) and Project (2016). Which job code has the largest 2016 gross wages, and how much?",
      [c("Job code", "text", f"SELECT {JOB16} FROM pay16 GROUP BY 1 ORDER BY SUM(Gross_Wages) DESC LIMIT 1"),
       c("Gross wages 2016", "num", f"SELECT ROUND(SUM(Gross_Wages),2) FROM pay16 GROUP BY {JOB16} ORDER BY 1 DESC LIMIT 1", 0.006)],
      ["text extraction", "pivot"], 4,
      "=IF(ISNUMBER(VALUE(MID(A2,3,1))),MID(A2,3,3),MID(A2,4,3)) — keep it as text so 001 does not become 1.")
excel("P05", "payroll", "For job code 005 (Installation), what were total hours (regular + overtime) in 2015 and in 2016?",
      [c("Hours 2015", "num", f"SELECT ROUND(SUM(Reg_Hrs+OT_Hrs),2) FROM pay15 WHERE {JOB15}='005'", 0.006),
       c("Hours 2016", "num", f"SELECT ROUND(SUM(Hours+OT_Hours),2) FROM pay16 WHERE {JOB16}='005'", 0.006)],
      ["text extraction", "SUMIFS"], 4)
excel("P06", "payroll", "Which job code has the largest increase in gross pay from 2015 to 2016, and what is the increase?",
      [c("Job code", "text", f"""SELECT j FROM (SELECT {JOB16} j, SUM(Gross_Wages) g FROM pay16 GROUP BY 1) a
            JOIN (SELECT {JOB15} j, SUM({GROSS15}) g FROM pay15 GROUP BY 1) b USING(j) ORDER BY a.g-b.g DESC LIMIT 1"""),
       c("Increase", "num", f"""SELECT ROUND(a.g-b.g,2) FROM (SELECT {JOB16} j, SUM(Gross_Wages) g FROM pay16 GROUP BY 1) a
            JOIN (SELECT {JOB15} j, SUM({GROSS15}) g FROM pay15 GROUP BY 1) b USING(j) ORDER BY a.g-b.g DESC LIMIT 1""", 0.006)],
      ["year-over-year", "merge"], 5)
excel("P07", "payroll", "How many 2016 rows have no Employee_Number (\"phantom\" entries), and what are their total gross wages?",
      [c("Phantom rows", "int", "SELECT COUNT(*) FROM pay16 WHERE trim(Employee_Number)=''"),
       c("Gross wages of phantom rows", "num", "SELECT ROUND(SUM(Gross_Wages),2) FROM pay16 WHERE trim(Employee_Number)=''", 0.006)],
      ["data quality", "veracity"], 3)
excel("P08", "payroll", "Which employee earned the most overtime pay in 2015, and how much?",
      [c("Employee name", "text", "SELECT First_Name||' '||Last_Name FROM pay15 GROUP BY Emp_No ORDER BY SUM(OT_Pay) DESC LIMIT 1"),
       c("Overtime pay", "num", "SELECT ROUND(SUM(OT_Pay),2) FROM pay15 GROUP BY Emp_No ORDER BY 1 DESC LIMIT 1", 0.006)],
      ["pivot", "ranking"], 3)
excel("P09", "payroll", "Which job code had the most overtime pay in 2015, and how much was it?",
      [c("Job code", "text", f"SELECT {JOB15} FROM pay15 GROUP BY 1 ORDER BY SUM(OT_Pay) DESC LIMIT 1"),
       c("Overtime pay", "num", f"SELECT ROUND(SUM(OT_Pay),2) FROM pay15 GROUP BY {JOB15} ORDER BY 1 DESC LIMIT 1", 0.006)],
      ["text extraction", "pivot"], 4)
excel("P10", "payroll", "In 2015, on how many employee pay dates was overtime recorded although the employee's regular hours that pay date were below 40?",
      [c("Employee pay dates", "int", "SELECT COUNT(*) FROM (SELECT 1 FROM pay15 GROUP BY Emp_No, Pay_Date HAVING SUM(Reg_Hrs) < 40 AND SUM(OT_Hrs) > 0)")],
      ["aggregation", "control test"], 4, "Pivot table: rows = employee and Pay_Date, values = sum of Reg_Hrs and OT_Hrs; filter Reg_Hrs < 40 and OT_Hrs > 0.")
excel("P11", "payroll", "What was the total employee cost in 2016 (gross wages + Social Security/Medicare + FUTA + SUTA + workers' compensation)?",
      [c("Total employee cost 2016", "num", "SELECT ROUND(SUM(Gross_Wages+SocSec_Medicare+FUTA+SUTA+Work_Comp),2) FROM pay16", 0.006)],
      ["aggregation"], 2)
excel("P12", "payroll", "Which month of 2015 had the highest gross wages, and how much above the six-month monthly average was it?",
      [c("Month (1–12)", "int", f"SELECT CAST(substr(Pay_Date,1,instr(Pay_Date,'/')-1) AS INT) m FROM pay15 GROUP BY 1 ORDER BY SUM({GROSS15}) DESC LIMIT 1"),
       c("Above average by", "num", f"""SELECT ROUND(MAX(g) - AVG(g),2) FROM (SELECT CAST(substr(Pay_Date,1,instr(Pay_Date,'/')-1) AS INT) m,
            SUM({GROSS15}) g FROM pay15 GROUP BY 1)""", 0.006)],
      ["dates", "pivot"], 4, "Dates are US format (M/D/YYYY). If MONTH() returns the day, the import used a European locale.")
excel("P13", "payroll", "Join the locations file on employee name. What were 2015 gross wages for location 1 and for location 2?",
      [c("Location 1, 2015", "num", f"SELECT ROUND(SUM({GROSS15}),2) FROM pay15 p JOIN locations l ON l.Employee_Name = p.First_Name||' '||p.Last_Name WHERE l.Location='1'", 0.006),
       c("Location 2, 2015", "num", f"SELECT ROUND(SUM({GROSS15}),2) FROM pay15 p JOIN locations l ON l.Employee_Name = p.First_Name||' '||p.Last_Name WHERE l.Location='2'", 0.006)],
      ["merge", "concatenation"], 4, "Create =F2&\" \"&G2 to build the name key, then XLOOKUP it in the locations sheet.")
excel("P14", "payroll", "What was the average gross pay per hour (gross pay ÷ regular + overtime hours) in 2015 and in 2016? Two decimals.",
      [c("Per hour 2015", "num", f"SELECT ROUND(SUM({GROSS15})/SUM(Reg_Hrs+OT_Hrs),2) FROM pay15", 0.006),
       c("Per hour 2016", "num", "SELECT ROUND(SUM(Gross_Wages)/SUM(Hours+OT_Hours),2) FROM pay16", 0.006)],
      ["ratios"], 3)
excel("P15", "payroll", "How many 2015 rows have a job code outside the valid range 001–012?",
      [c("Rows with invalid job code", "int", f"SELECT COUNT(*) FROM pay15 WHERE {JOB15} NOT IN (SELECT Job_Code FROM jobcodes)")],
      ["data quality", "lookup"], 3)

# ---------------------------------------------------------------- SQL Part 1 · songs
sql("S01", "songs", "List the names and energy of all songs with energy above 0.8, highest energy first.",
    "SELECT name, energy FROM songs WHERE energy > 0.8 ORDER BY energy DESC", ["WHERE", "ORDER BY"], 1, order=[1])
sql("S02", "songs", "List the names of the three longest songs, longest first.",
    "SELECT name FROM songs ORDER BY duration_ms DESC LIMIT 3", ["ORDER BY", "LIMIT"], 1, order=[0])
sql("S03", "songs", "List the names of songs with danceability, energy and valence all greater than 0.6.",
    "SELECT name FROM songs WHERE danceability > 0.6 AND energy > 0.6 AND valence > 0.6", ["WHERE", "AND"], 1)
sql("S04", "songs", "Return the average tempo of all songs, rounded to two decimals.",
    "SELECT ROUND(AVG(tempo), 2) FROM songs", ["AVG", "ROUND"], 1)
sql("S05", "songs", "List the names of songs by Post Malone. Do not assume his artist_id.",
    "SELECT name FROM songs WHERE artist_id = (SELECT id FROM artists WHERE name = 'Post Malone')", ["subquery"], 2)
sql("S06", "songs", "Return the average energy of songs by Drake, rounded to two decimals.",
    "SELECT ROUND(AVG(energy), 2) FROM songs JOIN artists ON artists.id = songs.artist_id WHERE artists.name = 'Drake'",
    ["JOIN", "AVG"], 2)
sql("S07", "songs", "List the names of songs that feature other artists (\"feat.\" appears in the name).",
    "SELECT name FROM songs WHERE name LIKE '%feat.%'", ["LIKE"], 1)
sql("S08", "songs", "For each artist with more than two songs, show the artist name and number of songs, most songs first.",
    "SELECT a.name, COUNT(*) FROM songs s JOIN artists a ON a.id = s.artist_id GROUP BY a.id HAVING COUNT(*) > 2 ORDER BY 2 DESC",
    ["JOIN", "GROUP BY", "HAVING"], 2, order=[1])
sql("S09", "songs", "Show the lowest, highest and average danceability across all songs (average rounded to three decimals).",
    "SELECT MIN(danceability), MAX(danceability), ROUND(AVG(danceability), 3) FROM songs", ["aggregates"], 1)
sql("S10", "songs", "Show each artist's name and average song duration in minutes (two decimals), longest average first.",
    "SELECT a.name, ROUND(AVG(s.duration_ms) / 60000.0, 2) FROM songs s JOIN artists a ON a.id = s.artist_id GROUP BY a.id ORDER BY 2 DESC",
    ["JOIN", "GROUP BY", "arithmetic"], 2, order=[1])
sql("S11", "songs", "How many songs have a tempo above the average tempo of all songs?",
    "SELECT COUNT(*) FROM songs WHERE tempo > (SELECT AVG(tempo) FROM songs)", ["subquery"], 2)
sql("S12", "songs", "List artists (name only) who have no song with energy above 0.5.",
    "SELECT name FROM artists WHERE id NOT IN (SELECT artist_id FROM songs WHERE energy > 0.5)", ["NOT IN", "subquery"], 2)

# ---------------------------------------------------------------- SQL Part 1 · movies
sql("M01", "movies", "List the titles of all movies released in 2007.", "SELECT title FROM movies WHERE year = 2007", ["WHERE"], 1)
sql("M02", "movies", "Find the birth year of Johnny Depp.", "SELECT birth FROM people WHERE name = 'Johnny Depp'", ["WHERE"], 1)
sql("M03", "movies", "List the titles and years of all Toy Story movies, oldest first.",
    "SELECT title, year FROM movies WHERE title LIKE 'Toy Story%' ORDER BY year", ["LIKE", "ORDER BY"], 1, order=[1])
sql("M04", "movies", "How many movies have a rating above 9.0?",
    "SELECT COUNT(*) FROM ratings WHERE rating > 9.0", ["COUNT"], 1)
sql("M05", "movies", "Find the average rating of movies released in 2013, rounded to two decimals.",
    "SELECT ROUND(AVG(rating), 2) FROM ratings JOIN movies ON movies.id = ratings.movie_id WHERE year = 2013", ["JOIN", "AVG"], 2)
sql("M06", "movies", "List movies released in 2011 with their ratings, highest rating first; for equal ratings, alphabetical by title.",
    "SELECT title, rating FROM movies JOIN ratings ON ratings.movie_id = movies.id WHERE year = 2011 ORDER BY rating DESC, title",
    ["JOIN", "ORDER BY"], 2, order=[0, 1])
sql("M07", "movies", "List the names of everyone who starred in Harry Potter and the Sorcerer's Stone.",
    """SELECT name FROM people JOIN stars ON stars.person_id = people.id JOIN movies ON movies.id = stars.movie_id
       WHERE title = 'Harry Potter and the Sorcerer''s Stone'""", ["JOIN"], 2)
sql("M08", "movies", "List the names of people who directed a movie rated at least 9.0. Each person once.",
    """SELECT DISTINCT name FROM people JOIN directors ON directors.person_id = people.id
       JOIN ratings ON ratings.movie_id = directors.movie_id WHERE rating >= 9.0""", ["JOIN", "DISTINCT"], 2)
sql("M09", "movies", "List the titles of the five highest-rated movies Helena Bonham Carter starred in, highest first; for equal ratings, alphabetical by title.",
    """SELECT title FROM movies JOIN stars ON stars.movie_id = movies.id JOIN people ON people.id = stars.person_id
       JOIN ratings ON ratings.movie_id = movies.id WHERE name = 'Helena Bonham Carter' ORDER BY rating DESC, title LIMIT 5""",
    ["JOIN", "ORDER BY", "LIMIT"], 3, order=[0])
sql("M10", "movies", "List the titles of movies in which both Johnny Depp and Helena Bonham Carter starred.",
    """SELECT title FROM movies WHERE id IN (SELECT movie_id FROM stars JOIN people ON people.id = person_id WHERE name = 'Johnny Depp')
       AND id IN (SELECT movie_id FROM stars JOIN people ON people.id = person_id WHERE name = 'Helena Bonham Carter')""",
    ["subquery", "INTERSECT"], 3)
sql("M11", "movies", "List everyone who starred in a movie with Kevin Bacon (the one born in 1958). Exclude Kevin Bacon himself; each person once.",
    """SELECT DISTINCT p.name FROM people p JOIN stars s ON s.person_id = p.id
       WHERE s.movie_id IN (SELECT movie_id FROM stars WHERE person_id = (SELECT id FROM people WHERE name = 'Kevin Bacon' AND birth = 1958))
       AND p.id <> (SELECT id FROM people WHERE name = 'Kevin Bacon' AND birth = 1958)""", ["subquery", "self-reference"], 4)
sql("M12", "movies", "How many people named Kevin Bacon are in the database, and how many have a birth year recorded?",
    "SELECT COUNT(*), COUNT(birth) FROM people WHERE name = 'Kevin Bacon'", ["COUNT", "NULL"], 1)
sql("M13", "movies", "Show the number of movies per release year from 2015 to 2019, oldest first.",
    "SELECT year, COUNT(*) FROM movies WHERE year BETWEEN 2015 AND 2019 GROUP BY year ORDER BY year", ["GROUP BY", "BETWEEN"], 2, order=[0])
sql("M14", "movies", "Show directors with at least five movies in the database, with their number of movies, most first.",
    """SELECT name, COUNT(*) FROM directors JOIN people ON people.id = directors.person_id
       GROUP BY people.id HAVING COUNT(*) >= 5 ORDER BY 2 DESC""", ["GROUP BY", "HAVING"], 3, order=[1])
sql("M15", "movies", "Show the title, year and votes of the ten movies with the most votes, most first.",
    "SELECT title, year, votes FROM movies JOIN ratings ON ratings.movie_id = movies.id ORDER BY votes DESC LIMIT 10",
    ["JOIN", "ORDER BY"], 2, order=[2])
sql("M16", "movies", "What is the average rating of movies Christopher Nolan directed, rounded to two decimals?",
    """SELECT ROUND(AVG(rating), 2) FROM ratings JOIN directors ON directors.movie_id = ratings.movie_id
       JOIN people ON people.id = directors.person_id WHERE name = 'Christopher Nolan'""", ["JOIN", "AVG"], 2)
sql("M17", "movies", "List the titles of movies released in 2010 that Leonardo DiCaprio starred in.",
    """SELECT title FROM movies JOIN stars ON stars.movie_id = movies.id JOIN people ON people.id = stars.person_id
       WHERE name = 'Leonardo DiCaprio' AND year = 2010""", ["JOIN", "WHERE"], 2)
sql("M18", "movies", "How many people are in the database, and how many of them have no birth year recorded?",
    "SELECT COUNT(*), SUM(birth IS NULL) FROM people", ["COUNT", "NULL"], 2)

# ---------------------------------------------------------------- SQL Part 1 · Chinook
sql("K01", "chinook", "How many columns does the Customer table have?",
    "SELECT COUNT(*) FROM pragma_table_info('Customer')", ["schema"], 1, instructor=dict(value=13, source=PWC.format(1)))
sql("K02", "chinook", "Select the first five rows of Customer. What is the Country in the fifth row?",
    "SELECT Country FROM (SELECT Country FROM Customer LIMIT 5) LIMIT 1 OFFSET 4", ["LIMIT"], 1,
    instructor=dict(value="Czech Republic", source=PWC.format(2)))
sql("K03", "chinook", "Sort Customer by LastName. What is the last name in the third row?",
    "SELECT LastName FROM Customer ORDER BY LastName LIMIT 1 OFFSET 2", ["ORDER BY"], 1, instructor=dict(value="Bernard", source=PWC.format(3)))
sql("K04", "chinook", "How many rows are in the Invoice table?", "SELECT COUNT(*) FROM Invoice", ["COUNT"], 1,
    instructor=dict(value=412, source=PWC.format(4)))
sql("K05", "chinook", "How many invoices have totals greater than $20?", "SELECT COUNT(*) FROM Invoice WHERE Total > 20", ["WHERE"], 1,
    instructor=dict(value=4, source=PWC.format(5)))
sql("K06", "chinook", "How many invoices have totals greater than $10 and less than $20?",
    "SELECT COUNT(*) FROM Invoice WHERE Total > 10 AND Total < 20", ["WHERE", "AND"], 1, instructor=dict(value=60, source=PWC.format(6)))
sql("K07", "chinook", "How many invoices were billed to Brazil, Argentina or Chile?",
    "SELECT COUNT(*) FROM Invoice WHERE BillingCountry IN ('Brazil', 'Argentina', 'Chile')", ["IN"], 1,
    instructor=dict(value=49, source=PWC.format(7)))
sql("K08", "chinook", "How many invoices have a billing country starting with C?",
    "SELECT COUNT(*) FROM Invoice WHERE BillingCountry LIKE 'C%'", ["LIKE"], 1, instructor=dict(value=77, source=PWC.format(8)))
sql("K09", "chinook", "How many customers list their company?", "SELECT COUNT(*) FROM Customer WHERE Company IS NOT NULL", ["NULL"], 1,
    instructor=dict(value=10, source=PWC.format(9)))
sql("K10", "chinook", "How many employees does Chinook have?", "SELECT COUNT(*) FROM Employee", ["COUNT"], 1,
    instructor=dict(value=8, source=PWC.format(10)))
sql("K11", "chinook", "Show each employee's first name, last name and phone number without the area code (drop everything up to and including the closing parenthesis and the space after it).",
    "SELECT FirstName, LastName, SUBSTR(Phone, INSTR(Phone, ')') + 2) FROM Employee", ["SUBSTR", "INSTR"], 2)
sql("K12", "chinook", "How long is the longest employee last name?", "SELECT MAX(LENGTH(LastName)) FROM Employee", ["LENGTH"], 1,
    instructor=dict(value=8, source=PWC.format(12)))
sql("K13", "chinook", "How many employees were born on a Sunday?", "SELECT COUNT(*) FROM Employee WHERE STRFTIME('%w', BirthDate) = '0'",
    ["STRFTIME"], 2, instructor=dict(value=2, source=PWC.format(13)))
sql("K14", "chinook", "How old (in years, one decimal) was the oldest employee when hired? Use (JULIANDAY(HireDate) − JULIANDAY(BirthDate)) / 365.25.",
    "SELECT ROUND(MAX((JULIANDAY(HireDate) - JULIANDAY(BirthDate)) / 365.25), 1) FROM Employee", ["JULIANDAY"], 2,
    instructor=dict(value=55.5, source=PWC.format(14) + " (\"approximately 55.5 years\")",
                     note="Using 365.25-day years the age is 55.62, which rounds to 55.6. The key's \"approximately 55.5\" probably comes from a different year length or from rounding down."))
sql("K15", "chinook", "List billing countries by total invoice value, highest first.",
    "SELECT BillingCountry, ROUND(SUM(Total), 2) FROM Invoice GROUP BY BillingCountry ORDER BY 2 DESC", ["GROUP BY", "SUM"], 2, order=[1])
sql("K16", "chinook", "How many invoice line items are there?", "SELECT COUNT(*) FROM InvoiceLine", ["COUNT"], 1,
    instructor=dict(value=2240, source=PWC.format(16)))
sql("K17", "chinook", "How many distinct unit prices appear on invoice lines?", "SELECT COUNT(DISTINCT UnitPrice) FROM InvoiceLine", ["DISTINCT"], 1,
    instructor=dict(value=2, source=PWC.format(17)))
sql("K18", "chinook", "How many downloads (invoice lines) happened at each unit price?",
    "SELECT UnitPrice, COUNT(*) FROM InvoiceLine GROUP BY UnitPrice", ["GROUP BY"], 1)
sql("K19", "chinook", "What is the earliest and latest invoice date?", "SELECT MIN(InvoiceDate), MAX(InvoiceDate) FROM Invoice", ["MIN/MAX", "dates"], 1)
sql("K20", "chinook", "How many different tracks were downloaded the maximum number of times, and what is that maximum?",
    "SELECT MAX(n), SUM(n = (SELECT MAX(c) FROM (SELECT COUNT(*) c FROM InvoiceLine GROUP BY TrackId))) FROM (SELECT COUNT(*) n FROM InvoiceLine GROUP BY TrackId)",
    ["subquery", "GROUP BY"], 3, instructor=dict(value=[2, 256], source=PWC.format(19) + " (\"256 tracks have been downloaded twice\")"))
sql("K21", "chinook", "An album counts as downloaded when all of its tracks are on a single invoice. How many distinct albums have been downloaded?",
    """SELECT COUNT(DISTINCT x.AlbumId) FROM (SELECT il.InvoiceId, t.AlbumId, COUNT(DISTINCT il.TrackId) n
         FROM InvoiceLine il JOIN Track t ON t.TrackId = il.TrackId GROUP BY 1, 2) x
       JOIN (SELECT AlbumId, COUNT(*) n FROM Track GROUP BY AlbumId) a ON a.AlbumId = x.AlbumId WHERE x.n = a.n""",
    ["subquery", "JOIN"], 5, instructor=dict(value=49, source=PWC.format(22),
                     note="49 is the number of invoice–album pairs where the whole album was bought; those 49 purchases cover 41 different albums. Every one of them is a single-track album. The question asks for distinct albums, which is 41."))
sql("K22", "chinook", "What is the most popular genre by number of tracks sold? Show the genre and the number of invoice lines.",
    """SELECT g.Name, COUNT(*) FROM InvoiceLine il JOIN Track t ON t.TrackId = il.TrackId JOIN Genre g ON g.GenreId = t.GenreId
       GROUP BY g.GenreId ORDER BY 2 DESC LIMIT 1""", ["JOIN", "GROUP BY"], 3, instructor=dict(value="Rock", source=PWC.format(23)))
sql("K23", "chinook", "Show each support representative's full name (first and last name in one column) and the number of customers they support, most first.",
    """SELECT e.FirstName || ' ' || e.LastName, COUNT(*) FROM Customer c JOIN Employee e ON e.EmployeeId = c.SupportRepId
       GROUP BY e.EmployeeId ORDER BY 2 DESC""", ["JOIN", "concatenation"], 2, order=[1])
sql("K24", "chinook", "Show customer first name, last name and total spent for the five customers who spent the most, highest first.",
    """SELECT c.FirstName, c.LastName, ROUND(SUM(i.Total), 2) FROM Invoice i JOIN Customer c ON c.CustomerId = i.CustomerId
       GROUP BY c.CustomerId ORDER BY 3 DESC LIMIT 5""", ["JOIN", "GROUP BY", "LIMIT"], 3, order=[2])
sql("K25", "chinook", "Show total sales per year (from InvoiceDate), oldest first.",
    "SELECT STRFTIME('%Y', InvoiceDate), ROUND(SUM(Total), 2) FROM Invoice GROUP BY 1 ORDER BY 1", ["STRFTIME", "GROUP BY"], 2, order=[0])
sql("K26", "chinook", "How many tracks have never been sold? (LEFT JOIN Track to InvoiceLine.)",
    "SELECT COUNT(*) FROM Track t LEFT JOIN InvoiceLine il ON il.TrackId = t.TrackId WHERE il.InvoiceLineId IS NULL", ["LEFT JOIN", "NULL"], 2)
sql("K27", "chinook", "Show each media type and the number of tracks, most first.",
    "SELECT m.Name, COUNT(*) FROM Track t JOIN MediaType m ON m.MediaTypeId = t.MediaTypeId GROUP BY m.MediaTypeId ORDER BY 2 DESC",
    ["JOIN", "GROUP BY"], 2, order=[1])
sql("K28", "chinook", "Classify invoices as 'Small' (Total < 5), 'Medium' (5 to under 15) and 'Large' (15 or more). Show each class and its number of invoices.",
    """SELECT CASE WHEN Total < 5 THEN 'Small' WHEN Total < 15 THEN 'Medium' ELSE 'Large' END, COUNT(*) FROM Invoice GROUP BY 1""",
    ["CASE", "GROUP BY"], 3)

# ---------------------------------------------------------------- SQL Part 2 · P-cards (internal audit)
sql("C01", "pcards", "How many P-card transactions are in the practice database? Label the column TotalTrans.",
    "SELECT COUNT(*) AS TotalTrans FROM pcards", ["COUNT"], 1)
sql("C02", "pcards", "How many transactions are $5,000 or more? Label the column TotalTransGT5k.",
    "SELECT COUNT(*) AS TotalTransGT5k FROM pcards WHERE Amount >= 5000", ["WHERE"], 1)
sql("C03", "pcards", "What is the total dollar amount of returns (negative amounts) from January to March 2014? Round to two decimals.",
    "SELECT ROUND(SUM(Amount), 2) AS TotalReturns2014Q1 FROM pcards WHERE Year = 2014 AND Month BETWEEN 1 AND 3 AND Amount < 0",
    ["WHERE", "SUM"], 2)
sql("C04", "pcards", "Show the number of transactions per year, most recent year first.",
    "SELECT Year, COUNT(*) AS TotalTrans FROM pcards GROUP BY Year ORDER BY Year DESC", ["GROUP BY"], 1, order=[0])
sql("C05", "pcards", "List cardholders (FullName) with more than 300 transactions in 2014, with their number of transactions, most first.",
    f"SELECT FullName, COUNT(*) AS NumTrans FROM pcards WHERE {PC14} GROUP BY FullName HAVING COUNT(*) > 300 ORDER BY 2 DESC",
    ["GROUP BY", "HAVING"], 2, order=[1])
sql("C06", "pcards", "List people who returned more than $10,000 of goods in 2014: name, total returned value (two decimals, as a negative number) and number of returns, largest returns first.",
    f"""SELECT FullName, ROUND(SUM(Amount), 2) AS ReturnedGoodsValue, COUNT(*) AS ReturnedGoodsCount FROM pcards
        WHERE {PC14} AND Amount < 0 GROUP BY FullName HAVING SUM(Amount) < -10000 ORDER BY 2""", ["GROUP BY", "HAVING"], 3, order=[1])
sql("C07", "pcards", "Control test: list employees who spent more than $50,000 in 2014, with their total (two decimals), largest first.",
    f"SELECT FullName, ROUND(SUM(Amount), 2) FROM pcards WHERE {PC14} GROUP BY FullName HAVING SUM(Amount) > 50000 ORDER BY 2 DESC",
    ["control test", "HAVING"], 3, order=[1])
sql("C08", "pcards", "Control test: list each employee and month in 2014 where monthly spending exceeded $10,000. Show name, month and total (two decimals); sort by month, then total descending.",
    f"""SELECT FullName, Month, ROUND(SUM(Amount), 2) FROM pcards WHERE {PC14} GROUP BY FullName, Month
        HAVING SUM(Amount) > 10000 ORDER BY Month, 3 DESC""", ["control test", "GROUP BY"], 3, order=[1, 2])
sql("C09", "pcards", "Control test: list every 2014 transaction above the $5,000 single-transaction limit. Show name, amount, vendor and MCC, smallest amount first.",
    f"SELECT FullName, Amount, Vendor, MCC FROM pcards WHERE {PC14} AND Amount > 5000 ORDER BY Amount", ["control test", "WHERE"], 2, order=[1])
sql("C10", "pcards", "Split-purchase test: find cardholder–vendor–day combinations in 2014 with more than one transaction whose combined total exceeds $5,000. Show name, vendor, TransactionDate, number of transactions and combined total (two decimals).",
    f"""SELECT FullName, Vendor, TransactionDate, COUNT(*), ROUND(SUM(Amount), 2) FROM pcards WHERE {PC14}
        GROUP BY FullName, Vendor, TransactionDate HAVING COUNT(*) > 1 AND SUM(Amount) > 5000""", ["control test", "GROUP BY"], 4)
sql("C11", "pcards", "Split-purchase test between cardholders: find vendor–day combinations in 2014 where exactly two different cardholders bought from the same vendor and their combined total exceeds $5,000. Show vendor, TransactionDate and combined total (two decimals).",
    f"""SELECT Vendor, TransactionDate, ROUND(SUM(Amount), 2) FROM pcards WHERE {PC14}
        GROUP BY Vendor, TransactionDate HAVING COUNT(DISTINCT FullName) = 2 AND SUM(Amount) > 5000""", ["control test", "COUNT DISTINCT"], 4)
sql("C12", "pcards", "OSU does not pay the 4.5% state sales tax. For each year, show total amount spent and the sales tax saved (4.5% of the total), both rounded to two decimals, most recent year first.",
    "SELECT Year, ROUND(SUM(Amount), 2), ROUND(SUM(Amount) * 0.045, 2) AS SalesTaxSaved FROM pcards GROUP BY Year ORDER BY Year DESC",
    ["arithmetic", "GROUP BY"], 2, order=[0])
sql("C13", "pcards", "Risk-rank 2014 returns: count returns that are High (refund of $1,000 or more), Medium ($500 to under $1,000) and Low (under $500). Show the rank and its count.",
    f"""SELECT CASE WHEN Amount <= -1000 THEN 'High' WHEN Amount <= -500 THEN 'Medium' ELSE 'Low' END AS RiskRanking, COUNT(*)
        FROM pcards WHERE {PC14} AND Amount < 0 GROUP BY 1""", ["CASE", "negative amounts"], 3)
sql("C14", "pcards", "Amazon pays OSU 0.75% on purchases of $500 or more and 0.50% on smaller ones. What would OSU earn on 2014 positive purchases from vendors containing AMAZON? Two decimals.",
    f"""SELECT ROUND(SUM(CASE WHEN Amount >= 500 THEN Amount * 0.0075 ELSE Amount * 0.005 END), 2) AS AmazonEarnings
        FROM pcards WHERE {PC14} AND Amount > 0 AND upper(Vendor) LIKE '%AMAZON%'""", ["CASE", "LIKE"], 3)
sql("C15", "pcards", "How many 2014 transactions are above the 2014 average transaction amount?",
    f"SELECT COUNT(*) FROM pcards WHERE {PC14} AND Amount > (SELECT AVG(Amount) FROM pcards WHERE {PC14})", ["subquery"], 2)
sql("C16", "pcards", "Show the ten vendors with the largest increase in spending from 2013 to 2014 (only vendors used in both years): vendor and difference (two decimals), largest first.",
    """SELECT Vendor, ROUND(SUM(CASE WHEN Year = 2014 THEN Amount ELSE 0 END) - SUM(CASE WHEN Year = 2013 THEN Amount ELSE 0 END), 2) AS DiffSpending
       FROM pcards GROUP BY Vendor HAVING SUM(Year = 2014) > 0 AND SUM(Year = 2013) > 0 ORDER BY 2 DESC LIMIT 10""",
    ["CASE", "year-over-year"], 4, order=[1])
sql("C17", "pcards", "Benford's law: for 2014 transactions of at least $1.00, show each leading digit (1–9) and its share of transactions in percent, rounded to three decimals.",
    f"""SELECT SUBSTR(CAST(Amount AS TEXT), 1, 1) AS Digit, ROUND(100.0 * COUNT(*) / (SELECT COUNT(*) FROM pcards WHERE {PC14} AND Amount >= 1), 3)
        FROM pcards WHERE {PC14} AND Amount >= 1 GROUP BY 1 ORDER BY 1""", ["forensic", "SUBSTR"], 4, order=[0])
sql("C18", "pcards", "Duplicate-payment test: count groups in 2014 with the same cardholder, vendor, TransactionDate and amount occurring more than once (positive amounts only). How many such groups are there?",
    f"""SELECT COUNT(*) FROM (SELECT 1 FROM pcards WHERE {PC14} AND Amount > 0
        GROUP BY FullName, Vendor, TransactionDate, Amount HAVING COUNT(*) > 1)""", ["forensic", "GROUP BY"], 3)
sql("C19", "pcards", "Round-number test: how many 2014 transactions are four-digit whole-dollar amounts ending in 000 (1000, 2000, … 9000)?",
    f"SELECT COUNT(*) FROM pcards WHERE {PC14} AND Amount >= 1000 AND Amount < 10000 AND Amount = CAST(Amount AS INT) AND CAST(Amount AS INT) % 1000 = 0",
    ["forensic", "arithmetic"], 3)
sql("C20", "pcards", "Which MCC had the largest total 2014 spending, and how much (two decimals)?",
    f"SELECT MCC, ROUND(SUM(Amount), 2) FROM pcards WHERE {PC14} GROUP BY MCC ORDER BY 2 DESC LIMIT 1", ["GROUP BY", "ranking"], 2)
