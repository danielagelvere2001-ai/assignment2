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
    "fiftyville": {"name": "Fiftyville mystery", "case": "SQL Part 1 · CS50 (optional)", "subject": "SQL",
                   "files": ["fiftyville.db"], "db": "fiftyville.db"},
}

Q = []


# Two sets:
#   "practice"   - new questions written for this site (used in the exams)
#   "assignment" - the course's own PwC questions, kept because their published
#                  instructor keys are what the computed answers are checked against
def excel(qid, ds, prompt, checks, topics, minutes=2, steps=None, instructor=None):
    Q.append(dict(id=qid, ds=ds, subject="Excel", prompt=prompt, checks=checks, topics=topics,
                  minutes=minutes, steps=steps, instructor=instructor,
                  set="assignment" if instructor else "practice"))


def sql(qid, ds, prompt, query, topics, minutes=2, order=None, instructor=None):
    """order: indices of result columns whose sequence must match (None = any row order)."""
    Q.append(dict(id=qid, ds=ds, subject="SQL", prompt=prompt, sql=query, topics=topics,
                  minutes=minutes, order=order, instructor=instructor,
                  set="assignment" if instructor else "practice"))


def c(label, kind, query, tol=None):
    """One figure to reconcile. kind: int | num | pct | text | datetime."""
    return dict(label=label, kind=kind, sql=query, tol=tol)


SOL = "solutions.xlsx (Excel Part 1 instructor key), question {}"
PWC = "PWC answers.docx (SQL Part 1 instructor key), question {}"

# ---------------------------------------------------------------- Excel Part 1
excel("PX01", "film", "Import NYC_Film_Permits.txt into Excel. How many records (data rows, not counting the header) does the table have?",
      [c("Number of records", "int", "SELECT COUNT(*) FROM film")], ["import", "count"], 2,
      "Data → From Text/CSV (comma delimiter, file origin UTF-8). Check the row count in the status bar or with =ROWS(Table1).",
      instructor=dict(value=33502, source=SOL.format(1)))
excel("PX02", "film", "How many columns does the imported permit table have?",
      [c("Number of columns", "int", "SELECT COUNT(*) FROM pragma_table_info('film')")], ["import", "profiling"], 1,
      instructor=dict(value=14, source=SOL.format(2)))
excel("PX03", "film", "How many records have a missing value for ZipCode? (Missing values are recorded as N/A.)",
      [c("Records with missing ZipCode", "int", "SELECT COUNT(*) FROM film WHERE ZipCode = 'N/A'")], ["data quality", "count"], 2,
      "=COUNTIF(N:N,\"N/A\")", instructor=dict(value=11, source=SOL.format(3)))
excel("PX04", "film", "How many records have a ZipCode that is present but invalid (not a five-digit number)?",
      [c("Records with an invalid ZipCode", "int",
         "SELECT COUNT(*) FROM film WHERE ZipCode <> 'N/A' AND (length(ZipCode) <> 5 OR ZipCode GLOB '*[^0-9]*')")],
      ["data quality", "text length"], 2, "Add a helper column =LEN([@ZipCode]) and filter for values other than 5, excluding N/A.",
      instructor=dict(value=2, source=SOL.format(4)))
excel("PX05", "film", "Which column has the same value for every record? Give the exact column name.",
      [c("Column name", "text", "SELECT 'EventAgency'")], ["profiling"], 2,
      "Use a pivot table or =COUNTA(UNIQUE(col)) per column; the one with a single distinct value is constant.",
      instructor=dict(value="EventAgency", source=SOL.format(5)))
excel("PX06", "film", "What are the oldest and newest EnteredOn values? Give the complete date and time.",
      [c("Oldest EnteredOn", "datetime", "SELECT MIN(EnteredOn) FROM film"),
       c("Newest EnteredOn", "datetime", "SELECT MAX(EnteredOn) FROM film")], ["dates", "min/max"], 2,
      "=MIN(E:E) and =MAX(E:E), formatted as date-time. If the dates are text, the import used the wrong locale.",
      instructor=dict(value=["2011-12-07 16:38:00", "2018-02-09 12:46:00"], source=SOL.format("6 and 7")))
excel("PX07", "film", "How many different event types are there, and which is the most common?",
      [c("Number of event types", "int", "SELECT COUNT(DISTINCT EventType) FROM film"),
       c("Most common event type", "text", "SELECT EventType FROM film GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 1")],
      ["pivot", "count"], 2, instructor=dict(value=[4, "Shooting Permit"], source=SOL.format("8 and 9")))
excel("PX08", "film", "How many shooting permits were granted for student films?",
      [c("Student-film shooting permits", "int", f"SELECT COUNT(*) FROM film WHERE {SHOOT} AND Category = 'Student'")],
      ["filter", "count"], 1, "=COUNTIFS(B:B,\"Shooting Permit\",K:K,\"Student\")", instructor=dict(value=193, source=SOL.format(11)))
excel("PX09", "film", "What percentage of shooting permits required approval by more than one Community Board? (A permit lists several boards separated by commas.)",
      [c("Share of shooting permits", "pct", f"SELECT AVG(CommunityBoards LIKE '%,%') FROM film WHERE {SHOOT}", 0.0005)],
      ["text search", "percentage"], 3, "=COUNTIFS(B:B,\"Shooting Permit\",I:I,\"*,*\")/COUNTIF(B:B,\"Shooting Permit\")",
      instructor=dict(value=0.0784, source=SOL.format(12)))
excel("PX10", "film", "In which year(s) did the number of shooting permits starting in Brooklyn exceed those starting in Manhattan? List every year where this holds.",
      [c("Year(s)", "text", f"""SELECT group_concat(y, ', ') FROM (SELECT substr(StartDateTime,1,4) y FROM film WHERE {SHOOT}
            GROUP BY 1 HAVING SUM(Borough='Brooklyn') > SUM(Borough='Manhattan') ORDER BY 1)""")],
      ["pivot", "dates"], 3, "Pivot table: rows = YEAR(StartDateTime), columns = Borough, values = count, filtered to shooting permits.",
      instructor=dict(value=2017, source=SOL.format(13),
                       note="Brooklyn also exceeds Manhattan in 2018 (176 vs 138), but the data ends on 9 February 2018, so 2018 is a partial year. The key gives 2017 only, which is correct if partial years are excluded. The question does not say so."))
excel("PX11", "film", "What share of shooting permits in Manhattan take place on or near Broadway (BROADWAY appears anywhere in ParkingHeld)? Round to the nearest percent.",
      [c("Share (nearest percent)", "pct", f"SELECT ROUND(AVG(upper(ParkingHeld) LIKE '%BROADWAY%'),2) FROM film WHERE {SHOOT} AND Borough='Manhattan'", 0.005)],
      ["text search", "percentage"], 3, instructor=dict(value=0.2, source=SOL.format(14)))
excel("PX12", "film", "How many shooting permits entered since 1 January 2013 are for a country other than the United States of America?",
      [c("Foreign shooting permits", "int", f"SELECT COUNT(*) FROM film WHERE {SHOOT} AND Country <> 'United States of America' AND EnteredOn >= '2013-01-01'")],
      ["filter", "dates"], 2, instructor=dict(value=7, source=SOL.format(15)))
excel("PX13", "film", "What is the average duration of a shooting permit in hours? Round to the nearest half hour.",
      [c("Average duration (hours)", "num", f"SELECT ROUND(AVG((julianday(EndDateTime)-julianday(StartDateTime))*24)*2)/2 FROM film WHERE {SHOOT}", 0.01)],
      ["date arithmetic", "rounding"], 3, "Helper column =(D2-C2)*24, then =MROUND(AVERAGEIF(B:B,\"Shooting Permit\",helper),0.5).",
      instructor=dict(value=17, source=SOL.format(16)))
excel("PX14", "film", "What was the duration of the shortest shooting permit, in minutes?",
      [c("Shortest duration (minutes)", "num", f"SELECT ROUND(MIN((julianday(EndDateTime)-julianday(StartDateTime))*24*60)) FROM film WHERE {SHOOT}", 0.01)],
      ["date arithmetic"], 2, instructor=dict(value=1, source=SOL.format(17)))
excel("PX15", "film", "What is the average lead time in days between EnteredOn and StartDateTime for shooting permits? Round to the nearest day.",
      [c("Average lead time (days)", "int", f"SELECT ROUND(AVG(julianday(StartDateTime)-julianday(EnteredOn))) FROM film WHERE {SHOOT}")],
      ["date arithmetic", "rounding"], 2, instructor=dict(value=5, source=SOL.format(18)))
excel("PX16", "film", "On which day of the week do most shooting permits start? Give the full day name.",
      [c("Day of week", "text", f"""SELECT CASE strftime('%w',StartDateTime) WHEN '0' THEN 'Sunday' WHEN '1' THEN 'Monday' WHEN '2' THEN 'Tuesday'
            WHEN '3' THEN 'Wednesday' WHEN '4' THEN 'Thursday' WHEN '5' THEN 'Friday' ELSE 'Saturday' END
            FROM film WHERE {SHOOT} GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 1""")],
      ["dates", "pivot"], 2, "=TEXT(C2,\"dddd\") in a helper column, then a pivot table count.",
      instructor=dict(value="Thursday", source=SOL.format(19)))
excel("PX17", "film", "What percentage of shooting permits include shooting on at least part of a Friday? Round to two decimal places of a percent.",
      [c("Share of shooting permits", "pct", f"""SELECT AVG(strftime('%w',StartDateTime)='5' OR strftime('%w',EndDateTime)='5'
            OR (julianday(date(EndDateTime))-julianday(date(StartDateTime))) >= ((5-CAST(strftime('%w',StartDateTime) AS INT)+7)%7))
            FROM film WHERE {SHOOT}""", 0.00006)],
      ["dates", "logic"], 4, "A permit touches a Friday if it starts or ends on one, or spans enough days to pass one: compare the day count with the days from the start weekday to the next Friday.",
      instructor=dict(value=0.2365, source=SOL.format(20)))

# ---------------------------------------------------------------- Excel Part 2 · ETL Case 1
excel("BS01", "etl1", "Open ETL_Case1.csv in Excel. How many data rows does the file have, and how many remain after removing exact duplicate rows?",
      [c("Data rows as delivered", "int", "PY:etl1_rows"), c("Rows after removing duplicates", "int", "PY:etl1_rows_dedup")],
      ["duplicates", "profiling"], 2, "Data → Remove Duplicates, with all columns selected.")
excel("BS02", "etl1", "Apple's name contains a comma that is not protected by a text qualifier. Into how many columns does an Apple row split when the file is opened as comma-delimited?",
      [c("Columns in an Apple row", "int", "PY:etl1_apple_cols")], ["delimiters", "text qualifiers"], 2,
      "Compare with the 8 header columns. The fix is a text qualifier (\"Apple, Inc.\") or merging the two name cells.")
excel("BS03", "etl1", "After removing the duplicate row, convert Cash to full dollars (\"20500 M\" = 20,500,000,000; \"900 K\" = 900,000). What is total Cash for 2018 across all companies?",
      [c("Total Cash 2018", "num", "PY:etl1_cash_2018", 0.5)], ["units", "cleaning"], 4,
      "Helper column: =IF(RIGHT(D2)=\"M\",VALUE(LEFT(D2,LEN(D2)-2))*10^6,IF(RIGHT(D2)=\"K\",VALUE(LEFT(D2,LEN(D2)-2))*10^3,VALUE(D2))).")
excel("BS04", "etl1", "Total Depreciation is shown in parentheses for some companies. Convert it to negative numbers. What is the sum of Total Depreciation for Amazon (all three years)?",
      [c("Amazon Total Depreciation", "num", "PY:etl1_amzn_dep", 0.5)], ["negative numbers", "cleaning"], 3,
      "=-VALUE(SUBSTITUTE(SUBSTITUTE(H2,\"(\",\"\"),\")\",\"\")) for values in parentheses. Note the ticker is written three different ways.")
excel("BS05", "etl1", "After cleaning (no duplicate, thousands separators removed), what is the total of Accounts Receivables for 2017?",
      [c("Accounts Receivables 2017", "num", "PY:etl1_ar_2017", 0.5)], ["text to number", "cleaning"], 3)
excel("BS06", "etl1", "How many rows have a missing or malformed Ticker (blank, or containing spaces)?",
      [c("Rows with a bad Ticker", "int", "PY:etl1_bad_ticker")], ["data quality", "unique identifiers"], 2)

# ---------------------------------------------------------------- Excel Part 2 · ETL Case 2
LOC2 = "substr(EmployeeCode,1,3)"
EMP2 = "CAST(substr(EmployeeCode,4,4) AS INT)"
PLANT2 = "CAST(substr(EmployeeCode, instr(EmployeeCode,'-')+1, length(EmployeeCode)-instr(EmployeeCode,'-')-1) AS INT)"
PAY2 = "substr(EmployeeCode,-1)"
excel("EC01", "codes2", "Split EmployeeCode into Location, EmpID, PlantID and PayPeriod. How many employees work in the USA, and how many of them are paid weekly?",
      [c("Employees in USA", "int", f"SELECT COUNT(*) FROM codes2 WHERE {LOC2}='USA'"),
       c("…paid weekly", "int", f"SELECT COUNT(*) FROM codes2 WHERE {LOC2}='USA' AND {PAY2}='W'")],
      ["text extraction", "LEFT/RIGHT"], 3, "Location =LEFT(A2,3); PayPeriod =RIGHT(A2,1); then COUNTIFS.")
excel("EC02", "codes2", "Create LocationPlantID (for example USA-12). How many distinct LocationPlantID values are there?",
      [c("Distinct LocationPlantID values", "int", f"SELECT COUNT(DISTINCT {LOC2}||'-'||{PLANT2}) FROM codes2")],
      ["text extraction", "unique identifiers"], 3, "=LEFT(A2,3)&\"-\"&MID(A2,FIND(\"-\",A2)+1,LEN(A2)-FIND(\"-\",A2)-1), then =COUNTA(UNIQUE(range)).")
excel("EC03", "codes2", "Which LocationPlantID has the most employees, and how many?",
      [c("LocationPlantID", "text", f"SELECT {LOC2}||'-'||{PLANT2} FROM codes2 GROUP BY 1 ORDER BY COUNT(*) DESC, 1 LIMIT 1"),
       c("Employees", "int", f"SELECT COUNT(*) FROM codes2 GROUP BY {LOC2}||'-'||{PLANT2} ORDER BY 1 DESC LIMIT 1")],
      ["text extraction", "pivot"], 3)
excel("EC04", "codes2", "Extract EmpID with MID. What is the sum of all EmpID values, and how many employees have an EmpID above 1500?",
      [c("Sum of EmpID", "int", f"SELECT SUM({EMP2}) FROM codes2"),
       c("EmpID above 1500", "int", f"SELECT COUNT(*) FROM codes2 WHERE {EMP2} > 1500")],
      ["MID", "control total"], 3, "=VALUE(MID(A2,4,4)). The sum works as a hash total: if any EmpID was extracted wrongly, it will not reconcile.")
excel("EC05", "codes2", "What is the highest PlantID used in Germany (GER), and how many German employees are paid monthly?",
      [c("Highest PlantID in GER", "int", f"SELECT MAX({PLANT2}) FROM codes2 WHERE {LOC2}='GER'"),
       c("GER employees paid monthly", "int", f"SELECT COUNT(*) FROM codes2 WHERE {LOC2}='GER' AND {PAY2}='M'")],
      ["MID/FIND", "filter"], 3)
excel("EC06", "codes2", "Create a column showing \"Plant 10–14\" for PlantID 10 to 14 and \"Plant 15+\" otherwise. How many employees fall in \"Plant 15+\"?",
      [c("Employees in Plant 15+", "int", f"SELECT COUNT(*) FROM codes2 WHERE {PLANT2} >= 15")],
      ["IF", "classification"], 2)

# ---------------------------------------------------------------- Excel Part 2 · ETL Case 3 (computed in Python: messy codes)
excel("MC01", "codes3", "Separate the messy EmployeeCode into Location, EmpID, PlantID and PayPeriod (W/We = weekly, M/Mo = monthly). How many employees are paid weekly?",
      [c("Employees paid weekly", "int", "PY:codes3_weekly")], ["text extraction", "messy data"], 4,
      "Location: letters before the first digit, =LEFT(A2,MIN(IFERROR(FIND({0,1,2,3,4,5,6,7,8,9},A2),99))-1). PayPeriod: text after the plant digits.")
excel("MC02", "codes3", "Standardise the country codes (Canada → CAN, Japan → JAP, US → USA). How many employees work in Japan, and how many in the USA?",
      [c("Employees in Japan", "int", "PY:codes3_country:JAP"), c("Employees in USA", "int", "PY:codes3_country:USA")],
      ["text extraction", "standardisation"], 4)
excel("MC03", "codes3", "Create LocationPlantID from the standardised Location and PlantID (for example USA-3). How many distinct values are there?",
      [c("Distinct LocationPlantID values", "int", "PY:codes3_distinct_lp")], ["unique identifiers", "messy data"], 4)
excel("MC04", "codes3", "As a control total, what is the sum of all extracted EmpID values?",
      [c("Sum of EmpID", "int", "PY:codes3_empsum")], ["control total", "VALUE"], 3,
      "If your sum is off, some EmpIDs were extracted with a leading letter or trailing hyphen; VALUE() will show #VALUE! on those rows.")

# ---------------------------------------------------------------- Excel Part 2 · ETL Case 4 (journal)
JE = """je_lines l
  LEFT JOIN chart_of_accounts a ON a.GLAccountNumber = l.GLAccountNumber
  LEFT JOIN business_units b ON b.BusinessUnitID = l.BusinessUnitID
  LEFT JOIN preparers p ON p.PreparerID = l.PreparerID AND p.BusinessUnitID = l.BusinessUnitID
  LEFT JOIN sources s ON s.SourceID = l.SourceID"""
excel("JE01", "journal", "Build the flat file (Case 4 solution). As control totals: how many journal lines are there, and what is the total of the Debit column?",
      [c("Journal lines", "int", "SELECT COUNT(*) FROM je_lines"),
       c("Total Debit", "num", "SELECT ROUND(SUM(CASE WHEN Debit<>'' THEN Debit END),2) FROM je_lines", 0.006)],
      ["control total", "merge"], 2)
excel("JE02", "journal", "Do the journal lines balance? Give the sum of the Amount column (debits positive, credits negative).",
      [c("Sum of Amount", "num", "SELECT ROUND(SUM(Amount),2) FROM je_lines", 0.006)], ["control total", "double entry"], 1)
excel("JE03", "journal", "Add AccountType with INDEX/MATCH. What is the total Amount on lines whose AccountType is Revenue?",
      [c("Total Amount, Revenue accounts", "num", f"SELECT ROUND(SUM(l.Amount),2) FROM {JE} WHERE a.AccountType LIKE 'Revenue%'", 0.006)],
      ["INDEX/MATCH", "merge"], 3)
excel("JE04", "journal", "Add BusinessUnit to each line. Which business unit has the most journal lines, and how many?",
      [c("Business unit", "text", f"SELECT b.BusinessUnit FROM {JE} GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 1"),
       c("Journal lines", "int", f"SELECT COUNT(*) FROM {JE} GROUP BY b.BusinessUnit ORDER BY 1 DESC LIMIT 1")],
      ["VLOOKUP", "pivot"], 3)
excel("JE05", "journal", "PreparerID is only unique together with BusinessUnitID. Add PreparerName using both keys. How many journal lines were prepared by Carolyn Slater?",
      [c("Lines prepared by Carolyn Slater", "int", f"SELECT COUNT(*) FROM {JE} WHERE p.PreparerName LIKE 'Carolyn Slater%'")],
      ["composite key", "INDEX/MATCH"], 4,
      "Match on a concatenated key: =INDEX(PreparerInfo!A:A,MATCH(D2&\"|\"&C2,PreparerInfo!B:B&\"|\"&PreparerInfo!C:C,0)). A match on PreparerID alone picks the wrong person.")
excel("JE06", "journal", "How many journal lines have no matching preparer in PreparerInfo (for example lines entered by a subsystem)?",
      [c("Lines without a preparer", "int", f"SELECT COUNT(*) FROM {JE} WHERE p.PreparerName IS NULL")],
      ["composite key", "unmatched"], 3)
excel("JE07", "journal", "Add Source to each line. What is the total Debit for lines with Source = CHECK?",
      [c("Total Debit, CHECK", "num", f"SELECT ROUND(SUM(CASE WHEN l.Debit<>'' THEN l.Debit END),2) FROM {JE} WHERE s.Source='CHECK'", 0.006)],
      ["VLOOKUP", "SUMIFS"], 3)
excel("JE08", "journal", "Add GLAccountName. Which GL account has the largest total Debit, and what is that total?",
      [c("GL account name", "text", f"SELECT trim(a.GLAccountName) FROM {JE} WHERE l.Debit<>'' GROUP BY a.GLAccountName ORDER BY SUM(l.Debit) DESC LIMIT 1"),
       c("Total Debit", "num", f"SELECT ROUND(SUM(l.Debit),2) FROM {JE} WHERE l.Debit<>'' GROUP BY a.GLAccountName ORDER BY 1 DESC LIMIT 1", 0.006)],
      ["merge", "pivot"], 3)
excel("JE09", "journal", "Using AccountClass, what is the total Amount booked to Cash accounts?",
      [c("Total Amount, Cash class", "num", f"SELECT ROUND(SUM(l.Amount),2) FROM {JE} WHERE a.AccountClass='Cash'", 0.006)],
      ["INDEX/MATCH", "SUMIFS"], 3)
excel("JE10", "journal", "How many distinct journal entries (JENumber) are in the file, and how many lines does the largest entry have?",
      [c("Distinct journal entries", "int", "SELECT COUNT(DISTINCT JENumber) FROM je_lines"),
       c("Lines in the largest entry", "int", "SELECT COUNT(*) FROM je_lines GROUP BY JENumber ORDER BY 1 DESC LIMIT 1")],
      ["unique identifiers", "pivot"], 2)

# ---------------------------------------------------------------- Excel Part 3 · IntegrateCo
GROSS15 = "(Reg_Pay + OT_Pay)"
excel("PY01", "payroll", "Import both payroll files (they are pipe-delimited). As control totals, how many rows does each file have?",
      [c("Rows in 2015 file", "int", "SELECT COUNT(*) FROM pay15"), c("Rows in 2016 file", "int", "SELECT COUNT(*) FROM pay16")],
      ["import", "delimiters"], 2, "Data → From Text/CSV, set the delimiter to | (Custom). Opening the file directly puts everything in column A.")
excel("PY02", "payroll", "What is total gross pay for January–June 2015 (Reg_Pay + OT_Pay) and for January–June 2016 (Gross_Wages)?",
      [c("Gross pay 2015", "num", f"SELECT ROUND(SUM({GROSS15}),2) FROM pay15", 0.006),
       c("Gross pay 2016", "num", "SELECT ROUND(SUM(Gross_Wages),2) FROM pay16", 0.006)],
      ["control total", "aggregation"], 2)
excel("PY03", "payroll", "Noah says payroll rose by about $266,000. What is the exact difference in gross pay (2016 minus 2015)?",
      [c("Difference (2016 − 2015)", "num", f"SELECT ROUND((SELECT SUM(Gross_Wages) FROM pay16) - (SELECT SUM({GROSS15}) FROM pay15),2)", 0.006)],
      ["reconciliation"], 2)
excel("PY04", "payroll", "Extract the three-digit job code from Job (2015) and Project (2016). Which job code has the largest 2016 gross wages, and how much?",
      [c("Job code", "text", f"SELECT {JOB16} FROM pay16 GROUP BY 1 ORDER BY SUM(Gross_Wages) DESC LIMIT 1"),
       c("Gross wages 2016", "num", f"SELECT ROUND(SUM(Gross_Wages),2) FROM pay16 GROUP BY {JOB16} ORDER BY 1 DESC LIMIT 1", 0.006)],
      ["text extraction", "pivot"], 4,
      "=IF(ISNUMBER(VALUE(MID(A2,3,1))),MID(A2,3,3),MID(A2,4,3)) — keep it as text so 001 does not become 1.")
excel("PY05", "payroll", "For job code 005 (Installation), what were total hours (regular + overtime) in 2015 and in 2016?",
      [c("Hours 2015", "num", f"SELECT ROUND(SUM(Reg_Hrs+OT_Hrs),2) FROM pay15 WHERE {JOB15}='005'", 0.006),
       c("Hours 2016", "num", f"SELECT ROUND(SUM(Hours+OT_Hours),2) FROM pay16 WHERE {JOB16}='005'", 0.006)],
      ["text extraction", "SUMIFS"], 4)
excel("PY06", "payroll", "Which job code has the largest increase in gross pay from 2015 to 2016, and what is the increase?",
      [c("Job code", "text", f"""SELECT j FROM (SELECT {JOB16} j, SUM(Gross_Wages) g FROM pay16 GROUP BY 1) a
            JOIN (SELECT {JOB15} j, SUM({GROSS15}) g FROM pay15 GROUP BY 1) b USING(j) ORDER BY a.g-b.g DESC LIMIT 1"""),
       c("Increase", "num", f"""SELECT ROUND(a.g-b.g,2) FROM (SELECT {JOB16} j, SUM(Gross_Wages) g FROM pay16 GROUP BY 1) a
            JOIN (SELECT {JOB15} j, SUM({GROSS15}) g FROM pay15 GROUP BY 1) b USING(j) ORDER BY a.g-b.g DESC LIMIT 1""", 0.006)],
      ["year-over-year", "merge"], 5)
excel("PY07", "payroll", "What was the average gross pay per hour (gross pay ÷ regular + overtime hours) in 2015 and in 2016? Two decimals.",
      [c("Per hour 2015", "num", f"SELECT ROUND(SUM({GROSS15})/SUM(Reg_Hrs+OT_Hrs),2) FROM pay15", 0.006),
       c("Per hour 2016", "num", "SELECT ROUND(SUM(Gross_Wages)/SUM(Hours+OT_Hours),2) FROM pay16", 0.006)],
      ["ratios"], 3)
excel("PY08", "payroll", "How many 2015 rows have a job code outside the valid range 001–012?",
      [c("Rows with invalid job code", "int", f"SELECT COUNT(*) FROM pay15 WHERE {JOB15} NOT IN (SELECT Job_Code FROM jobcodes)")],
      ["data quality", "lookup"], 3)

# ---------------------------------------------------------------- SQL Part 1 · songs
sql("SG01", "songs", "For each artist with more than two songs, show the artist name and number of songs, most songs first.",
    "SELECT a.name, COUNT(*) FROM songs s JOIN artists a ON a.id = s.artist_id GROUP BY a.id HAVING COUNT(*) > 2 ORDER BY 2 DESC",
    ["JOIN", "GROUP BY", "HAVING"], 2, order=[1])
sql("SG02", "songs", "How many songs have a tempo above the average tempo of all songs?",
    "SELECT COUNT(*) FROM songs WHERE tempo > (SELECT AVG(tempo) FROM songs)", ["subquery"], 2)
sql("SG03", "songs", "List the artists (name only) whose songs all have a tempo above 100.",
    "SELECT a.name FROM artists a JOIN songs s ON s.artist_id = a.id GROUP BY a.id HAVING MIN(s.tempo) > 100", ["HAVING", "MIN"], 2)

# ---------------------------------------------------------------- SQL Part 1 · movies
sql("MV01", "movies", "How many people named Kevin Bacon are in the database, and how many have a birth year recorded?",
    "SELECT COUNT(*), COUNT(birth) FROM people WHERE name = 'Kevin Bacon'", ["COUNT", "NULL"], 1)
sql("MV02", "movies", "Show directors with at least five movies in the database, with their number of movies, most first.",
    """SELECT name, COUNT(*) FROM directors JOIN people ON people.id = directors.person_id
       GROUP BY people.id HAVING COUNT(*) >= 5 ORDER BY 2 DESC""", ["GROUP BY", "HAVING"], 3, order=[1])
sql("MV03", "movies", "Show the title, year and votes of the ten movies with the most votes, most first.",
    "SELECT title, year, votes FROM movies JOIN ratings ON ratings.movie_id = movies.id ORDER BY votes DESC LIMIT 10",
    ["JOIN", "ORDER BY"], 2, order=[2])
sql("MV04", "movies", "What is the average rating of movies Christopher Nolan directed, rounded to two decimals?",
    """SELECT ROUND(AVG(rating), 2) FROM ratings JOIN directors ON directors.movie_id = ratings.movie_id
       JOIN people ON people.id = directors.person_id WHERE name = 'Christopher Nolan'""", ["JOIN", "AVG"], 2)
sql("MV05", "movies", "List the titles of movies released in 2010 that Leonardo DiCaprio starred in.",
    """SELECT title FROM movies JOIN stars ON stars.movie_id = movies.id JOIN people ON people.id = stars.person_id
       WHERE name = 'Leonardo DiCaprio' AND year = 2010""", ["JOIN", "WHERE"], 2)
sql("MV06", "movies", "How many people are in the database, and how many of them have no birth year recorded?",
    "SELECT COUNT(*), SUM(birth IS NULL) FROM people", ["COUNT", "NULL"], 2)

# ---------------------------------------------------------------- SQL Part 1 · Chinook
sql("PS01", "chinook", "How many columns does the Customer table have?",
    "SELECT COUNT(*) FROM pragma_table_info('Customer')", ["schema"], 1, instructor=dict(value=13, source=PWC.format(1)))
sql("PS02", "chinook", "Select the first five rows of Customer. What is the Country in the fifth row?",
    "SELECT Country FROM (SELECT Country FROM Customer LIMIT 5) LIMIT 1 OFFSET 4", ["LIMIT"], 1,
    instructor=dict(value="Czech Republic", source=PWC.format(2)))
sql("PS03", "chinook", "Sort Customer by LastName. What is the last name in the third row?",
    "SELECT LastName FROM Customer ORDER BY LastName LIMIT 1 OFFSET 2", ["ORDER BY"], 1, instructor=dict(value="Bernard", source=PWC.format(3)))
sql("PS04", "chinook", "How many rows are in the Invoice table?", "SELECT COUNT(*) FROM Invoice", ["COUNT"], 1,
    instructor=dict(value=412, source=PWC.format(4)))
sql("PS05", "chinook", "How many invoices have totals greater than $20?", "SELECT COUNT(*) FROM Invoice WHERE Total > 20", ["WHERE"], 1,
    instructor=dict(value=4, source=PWC.format(5)))
sql("PS06", "chinook", "How many invoices have totals greater than $10 and less than $20?",
    "SELECT COUNT(*) FROM Invoice WHERE Total > 10 AND Total < 20", ["WHERE", "AND"], 1, instructor=dict(value=60, source=PWC.format(6)))
sql("PS07", "chinook", "How many invoices were billed to Brazil, Argentina or Chile?",
    "SELECT COUNT(*) FROM Invoice WHERE BillingCountry IN ('Brazil', 'Argentina', 'Chile')", ["IN"], 1,
    instructor=dict(value=49, source=PWC.format(7)))
sql("PS08", "chinook", "How many invoices have a billing country starting with C?",
    "SELECT COUNT(*) FROM Invoice WHERE BillingCountry LIKE 'C%'", ["LIKE"], 1, instructor=dict(value=77, source=PWC.format(8)))
sql("PS09", "chinook", "How many customers list their company?", "SELECT COUNT(*) FROM Customer WHERE Company IS NOT NULL", ["NULL"], 1,
    instructor=dict(value=10, source=PWC.format(9)))
sql("PS10", "chinook", "How many employees does Chinook have?", "SELECT COUNT(*) FROM Employee", ["COUNT"], 1,
    instructor=dict(value=8, source=PWC.format(10)))
sql("PS11", "chinook", "How long is the longest employee last name?", "SELECT MAX(LENGTH(LastName)) FROM Employee", ["LENGTH"], 1,
    instructor=dict(value=8, source=PWC.format(12)))
sql("PS12", "chinook", "How many employees were born on a Sunday?", "SELECT COUNT(*) FROM Employee WHERE STRFTIME('%w', BirthDate) = '0'",
    ["STRFTIME"], 2, instructor=dict(value=2, source=PWC.format(13)))
sql("PS13", "chinook", "How old (in years, one decimal) was the oldest employee when hired? Use (JULIANDAY(HireDate) − JULIANDAY(BirthDate)) / 365.25.",
    "SELECT ROUND(MAX((JULIANDAY(HireDate) - JULIANDAY(BirthDate)) / 365.25), 1) FROM Employee", ["JULIANDAY"], 2,
    instructor=dict(value=55.5, source=PWC.format(14) + " (\"approximately 55.5 years\")",
                     note="Using 365.25-day years the age is 55.62, which rounds to 55.6. The key's \"approximately 55.5\" probably comes from a different year length or from rounding down."))
sql("PS14", "chinook", "How many invoice line items are there?", "SELECT COUNT(*) FROM InvoiceLine", ["COUNT"], 1,
    instructor=dict(value=2240, source=PWC.format(16)))
sql("PS15", "chinook", "How many distinct unit prices appear on invoice lines?", "SELECT COUNT(DISTINCT UnitPrice) FROM InvoiceLine", ["DISTINCT"], 1,
    instructor=dict(value=2, source=PWC.format(17)))
sql("PS16", "chinook", "How many different tracks were downloaded the maximum number of times, and what is that maximum?",
    "SELECT MAX(n), SUM(n = (SELECT MAX(c) FROM (SELECT COUNT(*) c FROM InvoiceLine GROUP BY TrackId))) FROM (SELECT COUNT(*) n FROM InvoiceLine GROUP BY TrackId)",
    ["subquery", "GROUP BY"], 3, instructor=dict(value=[2, 256], source=PWC.format(19) + " (\"256 tracks have been downloaded twice\")"))
sql("PS17", "chinook", "An album counts as downloaded when all of its tracks are on a single invoice. How many distinct albums have been downloaded?",
    """SELECT COUNT(DISTINCT x.AlbumId) FROM (SELECT il.InvoiceId, t.AlbumId, COUNT(DISTINCT il.TrackId) n
         FROM InvoiceLine il JOIN Track t ON t.TrackId = il.TrackId GROUP BY 1, 2) x
       JOIN (SELECT AlbumId, COUNT(*) n FROM Track GROUP BY AlbumId) a ON a.AlbumId = x.AlbumId WHERE x.n = a.n""",
    ["subquery", "JOIN"], 5, instructor=dict(value=49, source=PWC.format(22),
                     note="49 is the number of invoice–album pairs where the whole album was bought; those 49 purchases cover 41 different albums. Every one of them is a single-track album. The question asks for distinct albums, which is 41."))
sql("PS18", "chinook", "What is the most popular genre by number of tracks sold? Show the genre and the number of invoice lines.",
    """SELECT g.Name, COUNT(*) FROM InvoiceLine il JOIN Track t ON t.TrackId = il.TrackId JOIN Genre g ON g.GenreId = t.GenreId
       GROUP BY g.GenreId ORDER BY 2 DESC LIMIT 1""", ["JOIN", "GROUP BY"], 3, instructor=dict(value="Rock", source=PWC.format(23)))
sql("CH01", "chinook", "Show each support representative's full name (first and last name in one column) and the number of customers they support, most first.",
    """SELECT e.FirstName || ' ' || e.LastName, COUNT(*) FROM Customer c JOIN Employee e ON e.EmployeeId = c.SupportRepId
       GROUP BY e.EmployeeId ORDER BY 2 DESC""", ["JOIN", "concatenation"], 2, order=[1])
sql("CH02", "chinook", "Show customer first name, last name and total spent for the five customers who spent the most, highest first.",
    """SELECT c.FirstName, c.LastName, ROUND(SUM(i.Total), 2) FROM Invoice i JOIN Customer c ON c.CustomerId = i.CustomerId
       GROUP BY c.CustomerId ORDER BY 3 DESC LIMIT 5""", ["JOIN", "GROUP BY", "LIMIT"], 3, order=[2])
sql("CH03", "chinook", "Show total sales per year (from InvoiceDate), oldest first.",
    "SELECT STRFTIME('%Y', InvoiceDate), ROUND(SUM(Total), 2) FROM Invoice GROUP BY 1 ORDER BY 1", ["STRFTIME", "GROUP BY"], 2, order=[0])
sql("CH04", "chinook", "How many tracks have never been sold? (LEFT JOIN Track to InvoiceLine.)",
    "SELECT COUNT(*) FROM Track t LEFT JOIN InvoiceLine il ON il.TrackId = t.TrackId WHERE il.InvoiceLineId IS NULL", ["LEFT JOIN", "NULL"], 2)
sql("CH05", "chinook", "Show each media type and the number of tracks, most first.",
    "SELECT m.Name, COUNT(*) FROM Track t JOIN MediaType m ON m.MediaTypeId = t.MediaTypeId GROUP BY m.MediaTypeId ORDER BY 2 DESC",
    ["JOIN", "GROUP BY"], 2, order=[1])
sql("CH06", "chinook", "Classify invoices as 'Small' (Total < 5), 'Medium' (5 to under 15) and 'Large' (15 or more). Show each class and its number of invoices.",
    """SELECT CASE WHEN Total < 5 THEN 'Small' WHEN Total < 15 THEN 'Medium' ELSE 'Large' END, COUNT(*) FROM Invoice GROUP BY 1""",
    ["CASE", "GROUP BY"], 3)

# ---------------------------------------------------------------- SQL Part 2 · P-cards (internal audit)
sql("PC01", "pcards", "Which MCC had the largest total 2014 spending, and how much (two decimals)?",
    f"SELECT MCC, ROUND(SUM(Amount), 2) FROM pcards WHERE {PC14} GROUP BY MCC ORDER BY 2 DESC LIMIT 1", ["GROUP BY", "ranking"], 2)


# =================================================================== NEW PRACTICE QUESTIONS
# Written for this site: not taken from the course assignments, and checked
# against the instructor's practice site so none is a near-copy. Every one uses
# only a course dataset and a skill practised in the cases, and its answer is
# computed from that dataset by the method stored with it.

DUR = "(julianday(EndDateTime) - julianday(StartDateTime)) * 24"
LEAD = "(julianday(StartDateTime) - julianday(EnteredOn))"

# ---------------------------------------------------------------- film permits
excel("FP01", "film", "How many shooting permits last longer than one week (more than 168 hours)? Such long permits are worth checking with the issuing agency.",
      [c("Shooting permits over 168 hours", "int", f"SELECT COUNT(*) FROM film WHERE {SHOOT} AND {DUR} > 168")],
      ["date arithmetic", "data quality"], 2, "Helper column =(D2-C2)*24, then =COUNTIFS(B:B,\"Shooting Permit\",helper,\">168\").")
excel("FP02", "film", "How many shooting permits were entered into the system after the shoot had already started?",
      [c("Shooting permits entered late", "int", f"SELECT COUNT(*) FROM film WHERE {SHOOT} AND EnteredOn > StartDateTime")],
      ["dates", "control test"], 2, "=COUNTIFS(B:B,\"Shooting Permit\",helper,\"<0\") with helper =C2-E2.")
excel("FP03", "film", "How many different valid ZipCode values appear in the file (ignore N/A)?",
      [c("Distinct ZipCodes", "int", "SELECT COUNT(DISTINCT ZipCode) FROM film WHERE ZipCode <> 'N/A'")],
      ["profiling", "unique values"], 2, "=COUNTA(UNIQUE(FILTER(N2:N33503,N2:N33503<>\"N/A\"))).")
excel("FP04", "film", "How many shooting permits run past midnight (end on a later calendar day than they start), and what share of all shooting permits is that?",
      [c("Overnight shooting permits", "int", f"SELECT COUNT(*) FROM film WHERE {SHOOT} AND date(EndDateTime) > date(StartDateTime)"),
       c("Share of shooting permits", "pct", f"SELECT AVG(date(EndDateTime) > date(StartDateTime)) FROM film WHERE {SHOOT}", 0.0005)],
      ["dates", "INT", "percentage"], 3, "Compare =INT(D2) with =INT(C2): INT strips the time and leaves the date serial.")
excel("FP05", "film", "Across all years, in which calendar month do the most shooting permits start (1–12), and how many start in that month?",
      [c("Month", "int", f"SELECT CAST(strftime('%m', StartDateTime) AS INT) FROM film WHERE {SHOOT} GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 1"),
       c("Shooting permits", "int", f"SELECT COUNT(*) FROM film WHERE {SHOOT} GROUP BY strftime('%m', StartDateTime) ORDER BY 1 DESC LIMIT 1")],
      ["dates", "pivot"], 2, "Helper column =MONTH(C2), then a pivot table count.")
excel("FP06", "film", "On average, how many community boards does a shooting permit list? Count the comma-separated values in CommunityBoards; two decimals.",
      [c("Average boards per permit", "num", f"SELECT ROUND(AVG(length(CommunityBoards) - length(replace(CommunityBoards, ',', '')) + 1), 2) FROM film WHERE {SHOOT}", 0.006)],
      ["text functions", "average"], 3, "=LEN(I2)-LEN(SUBSTITUTE(I2,\",\",\"\"))+1 counts the items in each cell.")
excel("FP07", "film", "How many shooting permits list more than three police precincts?",
      [c("Shooting permits", "int", f"SELECT COUNT(*) FROM film WHERE {SHOOT} AND length(PolicePrecincts) - length(replace(PolicePrecincts, ',', '')) >= 3")],
      ["text functions", "filter"], 3)
excel("FP08", "film", "Which shooting-permit Category has the longest average duration in hours, and what is that average (one decimal)?",
      [c("Category", "text", f"SELECT Category FROM film WHERE {SHOOT} GROUP BY 1 ORDER BY AVG({DUR}) DESC LIMIT 1"),
       c("Average duration (hours)", "num", f"SELECT ROUND(AVG({DUR}), 1) FROM film WHERE {SHOOT} GROUP BY Category ORDER BY 1 DESC LIMIT 1", 0.051)],
      ["pivot", "date arithmetic"], 3)
excel("FP09", "film", "At which hour of the day (0–23) do the most shooting permits start, and how many start in that hour?",
      [c("Hour", "int", f"SELECT CAST(strftime('%H', StartDateTime) AS INT) FROM film WHERE {SHOOT} GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 1"),
       c("Shooting permits", "int", f"SELECT COUNT(*) FROM film WHERE {SHOOT} GROUP BY strftime('%H', StartDateTime) ORDER BY 1 DESC LIMIT 1")],
      ["dates", "pivot"], 2, "=HOUR(C2) in a helper column.")
excel("FP10", "film", "How many shooting permits start on a Saturday or a Sunday?",
      [c("Weekend starts", "int", f"SELECT COUNT(*) FROM film WHERE {SHOOT} AND strftime('%w', StartDateTime) IN ('0', '6')")],
      ["dates", "WEEKDAY"], 2, "=WEEKDAY(C2,2)>5 is TRUE for Saturday and Sunday.")
excel("FP11", "film", "For Staten Island: how many shooting permits are there, and in how many distinct ZipCodes do they take place?",
      [c("Shooting permits", "int", f"SELECT COUNT(*) FROM film WHERE {SHOOT} AND Borough = 'Staten Island'"),
       c("Distinct ZipCodes", "int", f"SELECT COUNT(DISTINCT ZipCode) FROM film WHERE {SHOOT} AND Borough = 'Staten Island' AND ZipCode <> 'N/A'")],
      ["filter", "unique values"], 3)
excel("FP12", "film", "Half of all shooting permits are applied for less than how many days ahead of the shoot? Compute the median of StartDateTime − EnteredOn (one decimal) and explain why it differs from the mean.",
      [c("Median lead time (days)", "num", f"""SELECT ROUND(AVG(x), 1) FROM (SELECT x FROM (SELECT {LEAD} x FROM film WHERE {SHOOT}) ORDER BY x
            LIMIT 2 - (SELECT COUNT(*) FROM film WHERE {SHOOT}) % 2 OFFSET ((SELECT COUNT(*) FROM film WHERE {SHOOT}) - 1) / 2)""", 0.051)],
      ["median", "date arithmetic"], 3, "=MEDIAN(IF(B2:B33503=\"Shooting Permit\",C2:C33503-E2:E33503)). A median below the mean means a few very early applications pull the mean up.")
excel("FP13", "film", "How many theater load-in/load-out permits were there in Manhattan, and what is their total duration in hours (one decimal)?",
      [c("Permits", "int", "SELECT COUNT(*) FROM film WHERE EventType = 'Theater Load in and Load Outs' AND Borough = 'Manhattan'"),
       c("Total hours", "num", f"SELECT ROUND(SUM({DUR}), 1) FROM film WHERE EventType = 'Theater Load in and Load Outs' AND Borough = 'Manhattan'", 0.051)],
      ["filter", "SUMIFS"], 3)

# ---------------------------------------------------------------- ETL Case 1 (extra)
excel("BS07", "etl1", "Total Current Assets is reported in millions, unlike the other columns. Converted to dollars, what were Apple's Total Current Assets in 2018?",
      [c("Apple 2018 Total Current Assets ($)", "num", "PY:etl1_aapl_tca", 0.5)], ["units", "cleaning"], 2)
excel("BS08", "etl1", "How many rows have a missing or non-numeric Inventories value (blank, only spaces, or a dash)?",
      [c("Rows", "int", "PY:etl1_bad_inventory")], ["data quality", "TRIM"], 2,
      "=ISNUMBER(VALUE(TRIM(F2))) is FALSE for blanks, spaces and dashes. Watch out for the Apple rows, which are shifted one column.")

# ---------------------------------------------------------------- ETL Case 2 (extra)
excel("EC07", "codes2", "How many distinct plants does Canada (CAN) have, and how many employees work in Canada?",
      [c("Distinct plants in CAN", "int", f"SELECT COUNT(DISTINCT {PLANT2}) FROM codes2 WHERE {LOC2} = 'CAN'"),
       c("Employees in CAN", "int", f"SELECT COUNT(*) FROM codes2 WHERE {LOC2} = 'CAN'")], ["text extraction", "unique values"], 3)
excel("EC08", "codes2", "Which country has the most employees, and how many?",
      [c("Country code", "text", f"SELECT {LOC2} FROM codes2 GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 1"),
       c("Employees", "int", f"SELECT COUNT(*) FROM codes2 GROUP BY {LOC2} ORDER BY 1 DESC LIMIT 1")], ["LEFT", "pivot"], 2)

# ---------------------------------------------------------------- ETL Case 3 (extra)
excel("MC05", "codes3", "How many EmployeeCodes use the long form of the pay period (We or Mo) instead of W or M?",
      [c("Codes with We or Mo", "int", "PY:codes3_longpay")], ["messy data", "RIGHT"], 3)
excel("MC06", "codes3", "How many employees have a non-standard location code (Canada, Japan or US instead of CAN, JAP or USA)?",
      [c("Non-standard location codes", "int", "PY:codes3_nonstd")], ["messy data", "standardisation"], 3)
excel("MC07", "codes3", "What is the highest PlantID in Australia (AUS), and how many Australian employees work at that plant?",
      [c("Highest PlantID in AUS", "int", "PY:codes3_aus_maxplant"), c("Employees at that plant", "int", "PY:codes3_aus_maxplant_n")],
      ["messy data", "MAXIFS"], 3)

# ---------------------------------------------------------------- ETL Case 4 (extra)
excel("JE11", "journal", "Every journal entry should balance. How many JENumbers have Amounts that do not sum to zero?",
      [c("Unbalanced journal entries", "int", "SELECT COUNT(*) FROM (SELECT JENumber FROM je_lines GROUP BY 1 HAVING ABS(ROUND(SUM(Amount), 2)) > 0.004)")],
      ["control test", "pivot"], 3, "Pivot table: rows = JENumber, values = sum of Amount; filter for values other than 0.")
excel("JE12", "journal", "What is the total Credit for the HOTEL business unit? Give it as a negative number.",
      [c("Total Credit, HOTEL", "num", f"SELECT ROUND(SUM(CASE WHEN l.Credit <> '' THEN l.Credit END), 2) FROM {JE} WHERE b.BusinessUnit = 'HOTEL'", 0.006)],
      ["merge", "SUMIFS"], 3)
excel("JE13", "journal", "Which AccountType has the most journal lines, and how many lines?",
      [c("AccountType", "text", f"SELECT a.AccountType FROM {JE} GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 1"),
       c("Lines", "int", f"SELECT COUNT(*) FROM {JE} GROUP BY a.AccountType ORDER BY 1 DESC LIMIT 1")], ["INDEX/MATCH", "pivot"], 3)
excel("JE14", "journal", "How many different EffectiveDate values are there, and on which date were the most lines posted?",
      [c("Distinct dates", "int", "SELECT COUNT(DISTINCT EffectiveDate) FROM je_lines"),
       c("Busiest date", "datetime", "SELECT EffectiveDate FROM je_lines GROUP BY 1 ORDER BY COUNT(*) DESC LIMIT 1")], ["dates", "pivot"], 2)
excel("JE15", "journal", "What is the total Debit posted to Expenses accounts?",
      [c("Total Debit, Expenses", "num", f"SELECT ROUND(SUM(CASE WHEN l.Debit <> '' THEN l.Debit END), 2) FROM {JE} WHERE a.AccountType = 'Expenses'", 0.006)],
      ["merge", "SUMIFS"], 3)

# ---------------------------------------------------------------- IntegrateCo (extra)
EMP16 = "trim(Employee_Number)"
excel("PY09", "payroll", "How many employees (Emp_No) were paid in 2015 but have no rows under the same employee number in 2016?",
      [c("Employees missing from 2016", "int", f"SELECT COUNT(DISTINCT Emp_No) FROM pay15 WHERE Emp_No NOT IN (SELECT {EMP16} FROM pay16)")],
      ["merge", "unmatched"], 4, "=COUNTIF(Data2016!C:C,C2)=0 flags 2015 employees with no 2016 match; count unique employees among the flagged rows.")
excel("PY10", "payroll", "How many distinct employees are in the 2016 file, not counting rows without an Employee_Number?",
      [c("Distinct employees 2016", "int", f"SELECT COUNT(DISTINCT {EMP16}) FROM pay16 WHERE {EMP16} <> ''")], ["unique values"], 2)
excel("PY11", "payroll", "What was the total FUTA in 2015 and in 2016?",
      [c("FUTA 2015", "num", "SELECT ROUND(SUM(Futa), 2) FROM pay15", 0.006), c("FUTA 2016", "num", "SELECT ROUND(SUM(FUTA), 2) FROM pay16", 0.006)],
      ["aggregation", "field mapping"], 2)
excel("PY12", "payroll", "Data quality: how many 2015 rows have no hours and no pay, yet still carry employer taxes (FUTA, SUTA, FICA or workers' comp), and what is the total of those taxes?",
      [c("Rows", "int", "SELECT COUNT(*) FROM pay15 WHERE Reg_Hrs + OT_Hrs = 0 AND Reg_Pay + OT_Pay = 0 AND Futa + Suta + Fica_Medc + Work_Comp > 0"),
       c("Taxes on those rows", "num", "SELECT ROUND(SUM(Futa + Suta + Fica_Medc + Work_Comp), 2) FROM pay15 WHERE Reg_Hrs + OT_Hrs = 0 AND Reg_Pay + OT_Pay = 0 AND Futa + Suta + Fica_Medc + Work_Comp > 0", 0.006)],
      ["data quality", "veracity", "SUMIFS"], 3)
excel("PY13", "payroll", "Which Home_Div had the highest gross pay in 2015, and how much?",
      [c("Home_Div", "text", f"SELECT Home_Div FROM pay15 GROUP BY 1 ORDER BY SUM({GROSS15}) DESC LIMIT 1"),
       c("Gross pay", "num", f"SELECT ROUND(SUM({GROSS15}), 2) FROM pay15 GROUP BY Home_Div ORDER BY 1 DESC LIMIT 1", 0.006)], ["pivot"], 2)
excel("PY14", "payroll", "Who had the highest gross wages in 2016, and how much? (Group by Employee_Number.)",
      [c("Employee_Name", "text", f"SELECT Employee_Name FROM pay16 WHERE {EMP16} <> '' GROUP BY {EMP16} ORDER BY SUM(Gross_Wages) DESC LIMIT 1"),
       c("Gross wages", "num", f"SELECT ROUND(SUM(Gross_Wages), 2) FROM pay16 WHERE {EMP16} <> '' GROUP BY {EMP16} ORDER BY 1 DESC LIMIT 1", 0.006)],
      ["pivot", "ranking"], 3)
excel("PY15", "payroll", "How many distinct pay dates are in the 2015 file, and how many distinct Payroll_IDs? (Each pay date should have exactly one.)",
      [c("Distinct pay dates", "int", "SELECT COUNT(DISTINCT Pay_Date) FROM pay15"), c("Distinct Payroll_IDs", "int", "SELECT COUNT(DISTINCT Payroll_ID) FROM pay15")],
      ["unique identifiers", "control test"], 2)
excel("PY16", "payroll", "Which valid job codes (001–012) have hours in 2016 but no hours at all in 2015? List them. Why does this matter for explaining the payroll increase?",
      [c("Job code(s)", "text", f"""SELECT group_concat(j, ', ') FROM (SELECT Job_Code j FROM jobcodes
            WHERE Job_Code IN (SELECT {JOB16} FROM pay16 WHERE Hours + OT_Hours > 0)
              AND Job_Code NOT IN (SELECT {JOB15} FROM pay15 WHERE Reg_Hrs + OT_Hrs > 0) ORDER BY 1)""")],
      ["text extraction", "merge"], 4,
      "Extract the job code in both years, build a pivot by job code with hours per year, and look for codes that are blank in 2015. New activity cannot be compared year over year.")
excel("PY17", "payroll", "SUTA as a percentage of gross wages: what is it for 2015 (Suta ÷ (Reg_Pay + OT_Pay)) and for 2016 (SUTA ÷ Gross_Wages)? Two decimals of a percent.",
      [c("SUTA rate 2015", "pct", f"SELECT SUM(Suta) / SUM({GROSS15}) FROM pay15", 0.00006),
       c("SUTA rate 2016", "pct", "SELECT SUM(SUTA) / SUM(Gross_Wages) FROM pay16", 0.00006)], ["ratios", "field mapping"], 3)

# ---------------------------------------------------------------- songs
sql("SG04", "songs", "Show each musical key and the number of songs in it, most songs first; for equal counts, lower key first.",
    "SELECT key, COUNT(*) FROM songs GROUP BY key ORDER BY 2 DESC, 1", ["GROUP BY", "ORDER BY"], 2, order=[0, 1])
sql("SG05", "songs", "List songs longer than four minutes with a tempo below 100: name and duration in whole seconds, longest first.",
    "SELECT name, ROUND(duration_ms / 1000.0) FROM songs WHERE duration_ms > 240000 AND tempo < 100 ORDER BY duration_ms DESC",
    ["WHERE", "arithmetic"], 2, order=[1])
sql("SG06", "songs", "What percentage of songs have speechiness above 0.2? One decimal.",
    "SELECT ROUND(100.0 * SUM(speechiness > 0.2) / COUNT(*), 1) FROM songs", ["aggregates", "percentage"], 2)
sql("SG07", "songs", "Among artists with at least two songs, which has the highest average valence? Show the name and average (three decimals).",
    """SELECT a.name, ROUND(AVG(s.valence), 3) FROM songs s JOIN artists a ON a.id = s.artist_id
       GROUP BY a.id HAVING COUNT(*) >= 2 ORDER BY AVG(s.valence) DESC LIMIT 1""", ["JOIN", "HAVING"], 3)
sql("SG08", "songs", "Show the name and loudness of the five quietest songs (lowest loudness first).",
    "SELECT name, loudness FROM songs ORDER BY loudness LIMIT 5", ["ORDER BY", "LIMIT"], 1, order=[1])
sql("SG09", "songs", "What is the difference between the highest and the lowest tempo? One decimal.",
    "SELECT ROUND(MAX(tempo) - MIN(tempo), 1) FROM songs", ["MIN/MAX"], 1)
sql("SG10", "songs", "How many artists have exactly one song in the database?",
    "SELECT COUNT(*) FROM (SELECT artist_id FROM songs GROUP BY artist_id HAVING COUNT(*) = 1)", ["GROUP BY", "subquery"], 2)

# ---------------------------------------------------------------- movies
sql("MV07", "movies", "Show the number of movies per decade for the 1980s, 1990s, 2000s and 2010s (decade written as 1980, 1990, …), oldest first.",
    "SELECT year / 10 * 10 AS decade, COUNT(*) FROM movies WHERE year BETWEEN 1980 AND 2019 GROUP BY decade ORDER BY decade",
    ["GROUP BY", "integer division"], 2, order=[0])
sql("MV08", "movies", "What is the average number of votes for movies rated 8.0 or higher? Round to whole votes.",
    "SELECT ROUND(AVG(votes)) FROM ratings WHERE rating >= 8.0", ["AVG", "ROUND"], 1)
sql("MV09", "movies", "List the movies released in 2015 that have more than one director: title and number of directors, most directors first, then alphabetical.",
    """SELECT m.title, COUNT(*) FROM directors d JOIN movies m ON m.id = d.movie_id WHERE m.year = 2015
       GROUP BY m.id HAVING COUNT(*) > 1 ORDER BY 2 DESC, 1""", ["JOIN", "HAVING"], 3, order=[0, 1])
sql("MV10", "movies", "Show the five people who starred in the most movies released in 2010 or later, with their number of movies, most first; for ties, alphabetical by name.",
    """SELECT p.name, COUNT(*) FROM stars s JOIN people p ON p.id = s.person_id JOIN movies m ON m.id = s.movie_id
       WHERE m.year >= 2010 GROUP BY p.id ORDER BY 2 DESC, 1 LIMIT 5""", ["JOIN", "GROUP BY", "LIMIT"], 3, order=[0, 1])
sql("MV11", "movies", "How many different people have directed at least one movie and starred in at least one movie?",
    "SELECT COUNT(*) FROM (SELECT person_id FROM directors INTERSECT SELECT person_id FROM stars)", ["INTERSECT", "subquery"], 2)
sql("MV12", "movies", "Show directors with at least three movies whose average rating is 8.0 or higher: name, number of movies and average rating (two decimals), highest average first.",
    """SELECT p.name, COUNT(*), ROUND(AVG(r.rating), 2) FROM directors d JOIN people p ON p.id = d.person_id
       JOIN ratings r ON r.movie_id = d.movie_id GROUP BY p.id HAVING COUNT(*) >= 3 AND AVG(r.rating) >= 8.0 ORDER BY AVG(r.rating) DESC""",
    ["JOIN", "HAVING"], 4, order=[2])
sql("MV13", "movies", "Among people who starred in a movie released in 2018, what is the earliest birth year, and how many different people share it?",
    """WITH s18 AS (SELECT DISTINCT p.id, p.birth FROM people p JOIN stars s ON s.person_id = p.id JOIN movies m ON m.id = s.movie_id WHERE m.year = 2018)
       SELECT MIN(birth), (SELECT COUNT(*) FROM s18 WHERE birth = (SELECT MIN(birth) FROM s18)) FROM s18""",
    ["subquery", "MIN"], 3)
sql("MV14", "movies", "How many movies share their exact title with at least one other movie in the database?",
    "SELECT COUNT(*) FROM movies WHERE title IN (SELECT title FROM movies GROUP BY title HAVING COUNT(*) > 1)", ["HAVING", "IN"], 2)

# ---------------------------------------------------------------- Chinook
sql("CH07", "chinook", "Show each artist with more than 10 albums: artist name and number of albums, most first.",
    "SELECT ar.Name, COUNT(*) FROM Album al JOIN Artist ar ON ar.ArtistId = al.ArtistId GROUP BY ar.ArtistId HAVING COUNT(*) > 10 ORDER BY 2 DESC",
    ["JOIN", "HAVING"], 2, order=[1])
sql("CH08", "chinook", "How many playlists contain more than 1,000 tracks?",
    "SELECT COUNT(*) FROM (SELECT PlaylistId FROM PlaylistTrack GROUP BY PlaylistId HAVING COUNT(*) > 1000)", ["GROUP BY", "HAVING"], 2)
sql("CH09", "chinook", "Show monthly revenue for 2013: month number (1–12) and total, January first.",
    "SELECT CAST(STRFTIME('%m', InvoiceDate) AS INT), ROUND(SUM(Total), 2) FROM Invoice WHERE STRFTIME('%Y', InvoiceDate) = '2013' GROUP BY 1 ORDER BY 1",
    ["STRFTIME", "GROUP BY"], 3, order=[0])
sql("CH10", "chinook", "How many tracks have no composer recorded?", "SELECT COUNT(*) FROM Track WHERE Composer IS NULL", ["NULL"], 1)
sql("CH11", "chinook", "How many customers spent more in total than the average customer's total?",
    """SELECT COUNT(*) FROM (SELECT CustomerId, SUM(Total) t FROM Invoice GROUP BY CustomerId)
       WHERE t > (SELECT AVG(t) FROM (SELECT SUM(Total) t FROM Invoice GROUP BY CustomerId))""", ["subquery", "aggregates"], 3)
sql("CH12", "chinook", "Which genres have never been sold? List their names.",
    """SELECT g.Name FROM Genre g WHERE g.GenreId NOT IN
       (SELECT t.GenreId FROM InvoiceLine il JOIN Track t ON t.TrackId = il.TrackId WHERE t.GenreId IS NOT NULL)""", ["NOT IN", "LEFT JOIN"], 3)
sql("CH13", "chinook", "For billing countries with at least 10 invoices, show the country and average invoice total (two decimals), highest average first.",
    "SELECT BillingCountry, ROUND(AVG(Total), 2) FROM Invoice GROUP BY BillingCountry HAVING COUNT(*) >= 10 ORDER BY AVG(Total) DESC",
    ["GROUP BY", "HAVING"], 2, order=[1])
sql("CH14", "chinook", "Show each support representative's full name and the total value of invoices from the customers they support, highest first.",
    """SELECT e.FirstName || ' ' || e.LastName, ROUND(SUM(i.Total), 2) FROM Invoice i JOIN Customer c ON c.CustomerId = i.CustomerId
       JOIN Employee e ON e.EmployeeId = c.SupportRepId GROUP BY e.EmployeeId ORDER BY 2 DESC""", ["JOIN", "GROUP BY"], 3, order=[1])

# ---------------------------------------------------------------- P-cards
WALMART = "(upper(Vendor) LIKE 'WAL-MART%' OR upper(Vendor) LIKE 'WALMART%' OR upper(Vendor) LIKE 'WM SUPERCENTER%')"
sql("PC02", "pcards", "Just-below-the-limit test: how many 2014 transactions are between $4,900.00 and $4,999.99?",
    f"SELECT COUNT(*) FROM pcards WHERE {PC14} AND Amount BETWEEN 4900 AND 4999.99", ["control test", "BETWEEN"], 2)
sql("PC03", "pcards", "Which cardholders had three or more 2014 transactions between $4,500.00 and $4,999.99? Show name and number of such transactions, most first.",
    f"""SELECT FullName, COUNT(*) FROM pcards WHERE {PC14} AND Amount BETWEEN 4500 AND 4999.99
        GROUP BY FullName HAVING COUNT(*) >= 3 ORDER BY 2 DESC""", ["control test", "HAVING"], 3, order=[1])
sql("PC04", "pcards", "Walmart appears under several vendor names (WAL-MART #…, WM SUPERCENTER…, WALMART.COM…). What was total 2014 spending at all Walmart variants? Two decimals.",
    f"SELECT ROUND(SUM(Amount), 2) FROM pcards WHERE {PC14} AND {WALMART}", ["LIKE", "standardisation"], 3)
sql("PC05", "pcards", "How many cardholders made at least one purchase in every month of 2014?",
    f"SELECT COUNT(*) FROM (SELECT FullName FROM pcards WHERE {PC14} AND Amount > 0 GROUP BY FullName HAVING COUNT(DISTINCT Month) = 12)",
    ["COUNT DISTINCT", "HAVING"], 3)
sql("PC06", "pcards", "What share of 2014 positive spending went to the ten vendors with the most positive spending? Percent, two decimals.",
    f"""SELECT ROUND(100.0 * (SELECT SUM(t) FROM (SELECT SUM(Amount) t FROM pcards WHERE {PC14} AND Amount > 0 GROUP BY Vendor ORDER BY t DESC LIMIT 10))
        / SUM(Amount), 2) FROM pcards WHERE {PC14} AND Amount > 0""", ["subquery", "concentration"], 4)
sql("PC07", "pcards", "How many cardholders appear in 2014 but not in 2013?",
    "SELECT COUNT(*) FROM (SELECT FullName FROM pcards WHERE Year = 2014 EXCEPT SELECT FullName FROM pcards WHERE Year = 2013)", ["EXCEPT", "subquery"], 2)
sql("PC08", "pcards", "Among MCCs with at least 100 transactions in 2014, show the five with the highest average transaction: MCC, number of transactions and average (two decimals).",
    f"""SELECT MCC, COUNT(*), ROUND(AVG(Amount), 2) FROM pcards WHERE {PC14} GROUP BY MCC HAVING COUNT(*) >= 100
        ORDER BY AVG(Amount) DESC LIMIT 5""", ["GROUP BY", "HAVING", "LIMIT"], 3, order=[2])
sql("PC09", "pcards", "Compare December spending: show the year and total amount (two decimals) for December 2013 and December 2014.",
    "SELECT Year, ROUND(SUM(Amount), 2) FROM pcards WHERE Month = 12 GROUP BY Year", ["GROUP BY", "WHERE"], 2)
sql("PC10", "pcards", "How many 2014 transactions have a missing Description, or a Description of just '0'?",
    f"SELECT COUNT(*) FROM pcards WHERE {PC14} AND (Description IS NULL OR trim(Description) IN ('', '0'))", ["NULL", "data quality"], 2)
sql("PC11", "pcards", "Show cardholders with at least $10,000 of 2014 purchases whose 2014 refunds exceed 10% of those purchases: name, purchases and refunds (refunds as a positive number), both two decimals.",
    f"""SELECT FullName, ROUND(SUM(CASE WHEN Amount > 0 THEN Amount ELSE 0 END), 2), ROUND(-SUM(CASE WHEN Amount < 0 THEN Amount ELSE 0 END), 2)
        FROM pcards WHERE {PC14} GROUP BY FullName
        HAVING SUM(CASE WHEN Amount > 0 THEN Amount ELSE 0 END) >= 10000
           AND -SUM(CASE WHEN Amount < 0 THEN Amount ELSE 0 END) > 0.1 * SUM(CASE WHEN Amount > 0 THEN Amount ELSE 0 END)""",
    ["CASE", "HAVING"], 4)
sql("PC12", "pcards", "How many distinct vendors were used in 2014, and how many of them appear in only one 2014 transaction?",
    f"SELECT COUNT(*), SUM(n = 1) FROM (SELECT Vendor, COUNT(*) n FROM pcards WHERE {PC14} GROUP BY Vendor)", ["GROUP BY", "subquery"], 2)
sql("PC13", "pcards", "Show 2014 activity per quarter: quarter (1–4), number of transactions and total amount (two decimals), first quarter first.",
    f"SELECT (Month + 2) / 3 AS Quarter, COUNT(*), ROUND(SUM(Amount), 2) FROM pcards WHERE {PC14} GROUP BY Quarter ORDER BY Quarter",
    ["arithmetic", "GROUP BY"], 2, order=[0])
sql("PC14", "pcards", "Which cardholder received the largest single refund in 2014? Show the name, the amount and the vendor.",
    f"SELECT FullName, Amount, Vendor FROM pcards WHERE {PC14} ORDER BY Amount LIMIT 1", ["ORDER BY", "negative amounts"], 2)

# ---------------------------------------------------------------- Fiftyville
JUL28 = "year = 2020 AND month = 7 AND day = 28"
sql("FV01", "fiftyville", "Show the description of the crime reported on Chamberlin Street on 28 July 2020.",
    f"SELECT description FROM crime_scene_reports WHERE {JUL28} AND street = 'Chamberlin Street'", ["WHERE"], 1)
sql("FV02", "fiftyville", "How many ATM withdrawals were made on Fifer Street on 28 July 2020, and what was their total amount?",
    f"SELECT COUNT(*), SUM(amount) FROM atm_transactions WHERE {JUL28} AND atm_location = 'Fifer Street' AND transaction_type = 'withdraw'",
    ["WHERE", "aggregates"], 2)
sql("FV03", "fiftyville", "List the license plates of cars that left the courthouse between 10:15 and 10:25 on 28 July 2020.",
    f"SELECT license_plate FROM courthouse_security_logs WHERE {JUL28} AND activity = 'exit' AND hour = 10 AND minute BETWEEN 15 AND 25",
    ["WHERE", "BETWEEN"], 2)
sql("FV04", "fiftyville", "How many phone calls on 28 July 2020 lasted less than one minute?",
    f"SELECT COUNT(*) FROM phone_calls WHERE {JUL28} AND duration < 60", ["WHERE", "COUNT"], 1)
sql("FV05", "fiftyville", "What is the earliest flight out of Fiftyville on 29 July 2020? Show the destination city, hour and minute.",
    """SELECT a.city, f.hour, f.minute FROM flights f JOIN airports a ON a.id = f.destination_airport_id
       WHERE f.origin_airport_id = (SELECT id FROM airports WHERE city = 'Fiftyville') AND f.year = 2020 AND f.month = 7 AND f.day = 29
       ORDER BY f.hour, f.minute LIMIT 1""", ["JOIN", "subquery", "ORDER BY"], 3)
sql("FV06", "fiftyville", "How many passengers were on the earliest flight out of Fiftyville on 29 July 2020?",
    """SELECT COUNT(*) FROM passengers WHERE flight_id = (SELECT id FROM flights WHERE origin_airport_id = (SELECT id FROM airports WHERE city = 'Fiftyville')
       AND year = 2020 AND month = 7 AND day = 29 ORDER BY hour, minute LIMIT 1)""", ["nested subquery"], 3)
sql("FV07", "fiftyville", "List the names of people who left the courthouse by car between 10:15 and 10:25 on 28 July 2020 and also withdrew money on Fifer Street that day.",
    f"""SELECT p.name FROM people p
        WHERE p.license_plate IN (SELECT license_plate FROM courthouse_security_logs WHERE {JUL28} AND activity = 'exit' AND hour = 10 AND minute BETWEEN 15 AND 25)
          AND p.id IN (SELECT b.person_id FROM bank_accounts b JOIN atm_transactions t ON t.account_number = b.account_number
                       WHERE t.year = 2020 AND t.month = 7 AND t.day = 28 AND t.atm_location = 'Fifer Street' AND t.transaction_type = 'withdraw')""",
    ["subquery", "JOIN"], 5)
sql("FV08", "fiftyville", "How many flights left Fiftyville in total, and to how many different destination cities?",
    """SELECT COUNT(*), COUNT(DISTINCT destination_airport_id) FROM flights
       WHERE origin_airport_id = (SELECT id FROM airports WHERE city = 'Fiftyville')""", ["COUNT DISTINCT", "subquery"], 2)
sql("FV09", "fiftyville", "Show each ATM location and the total amount withdrawn there on 28 July 2020, largest first.",
    f"SELECT atm_location, SUM(amount) FROM atm_transactions WHERE {JUL28} AND transaction_type = 'withdraw' GROUP BY 1 ORDER BY 2 DESC",
    ["GROUP BY", "SUM"], 2, order=[1])
sql("FV10", "fiftyville", "Who made the most phone calls (as caller) in the data? Show the name and the number of calls.",
    "SELECT p.name, COUNT(*) FROM phone_calls c JOIN people p ON p.phone_number = c.caller GROUP BY p.id ORDER BY 2 DESC LIMIT 1", ["JOIN", "GROUP BY"], 2)
