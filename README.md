# Data Analytics in Accounting: Practice Exams

A practice website for the oral exam in Data Analytics in Accounting (SDU), built only from the course materials.

- **Data manipulation exams**: two Excel questions and three SQL questions per exam, with an eight-minute timer. SQL runs in the browser (sql.js) on the course databases.
- **Question bank**: 148 questions covering NYC film permits, the four ETL cases, IntegrateCo payroll, CS50 songs and movies, Chinook and the OSU P-card audit.
- **Data interpretation exams**: regression models fitted on the merged Restaurant case data, which anyone can re-run in Excel.

## The improvement: an answer key a person stands behind

Students can check their own answers: they enter the reconciling figures for an Excel task, or run their query and compare result sets for SQL. Every answer shows who vouches for it:

| Badge | Meaning |
|---|---|
| Matches instructor key | Computed answer equals the published course answer (solutions.xlsx, PWC answers.docx) |
| Reviewed by *name* | A person re-performed the method and signed off; the sign-off is a commit to `content/reviews.json` |
| Instructor key disagrees | The published key and the data differ; both are shown, and a person decides |
| Changed since review | The question, method or data changed after sign-off, so the approval no longer applies |
| Not yet reviewed | Computed by a documented method, but no person has checked it yet |

The checker accepts Danish number and date formats and Excel date serials, so a correct answer is not marked wrong because of how it was typed.

## Repository layout

```
index.html, assets/      the static site (no build step)
content/                 generated answer key + human reviews (reviews.json)
data/                    downloadable datasets (movies and P-cards are practice subsets)
tools/make_subsets.py    builds data/ from the course materials
tools/build_content.py   computes every answer and writes content/
tools/questions.py       the questions and the method behind each answer
tools/regression.py      merges the Restaurant data and fits the regression models
```

Rebuild the answer key (Python 3 with `openpyxl`):

```
python tools/make_subsets.py   # needs ../materials
python tools/build_content.py
```

## Reviewing answers

Open `#/review` on the site, enter your name, re-perform and approve or withdraw answers, then download `reviews.json` and commit it to `content/reviews.json`. Each review is tied to a fingerprint of the question, method and answer.
