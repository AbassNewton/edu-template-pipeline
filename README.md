# edu-template-pipeline

Build, validate and consolidate **multi-year Excel data-collection templates** for
education statistics in Python, with an R cross-check and a Streamlit front end.

> Independent portfolio project on **synthetic data** (fictional countries `XAA`–`XAJ`).
> It is not affiliated with, and does not reproduce, any official UNESCO/UIS template.
> The indicator definitions are simplified versions of standard ones (GER, GPI, PTR,
> expenditure shares).

## What it does

| Step | Command | Output |
|---|---|---|
| Generate blank templates in EN, FR, ES, PT and RU from one spec | `python -m edutpl build` | `out/templates/*.xlsx` |
| Re-issue a filled template for a new cycle (new years), in any language | `python -m edutpl extend old.xlsx new.xlsx --last-year 2027` | pre-filled workbook |
| Create 10 synthetic country submissions with known injected errors | `python -m edutpl synth` | `out/submissions/` + `ground_truth.json` |
| Read, validate and consolidate every submission | `python -m edutpl process` | SQLite + CSV, QC log, one review report per country |
| All of the above | `python -m edutpl all` | |

```bash
pip install -r requirements.txt
set PYTHONPATH=src            # or: pip install -e .
python -m edutpl all
python -m pytest              # 23 tests
streamlit run app/streamlit_app.py
```

## Design choices

- **One spec, generated workbooks.** Years, levels, ages and thresholds live in
  `config/spec.yaml`; labels for five languages in `config/labels.yaml`. Adding a
  year means editing the spec and regenerating, then pre-filling from the data held.
  There is no manual cell editing, and every change can be reviewed in Git.
- **Codes, not labels.** Each data row carries a stable code in hidden column A
  (`ENR|L1|F|`, `POP||M|10`, ...). The reader keys on codes and the header years. A
  Russian submission, a renamed sheet or a reordered row is therefore read the same way.
- **Live checks inside the workbook.** The template computes indicators and
  consistency checks with Excel formulas. Countries see the problems before they send.
  Input cells are unlocked and validated, while calculated cells are locked.
- **Three independent implementations of the indicators.** These are the Excel
  formulas, `edutpl/indicators.py` and `r/crosscheck_indicators.R`. They agree on all
  1,360 indicator cells of the demo (`scripts/verify_formulas_excel.py` recalculates
  the workbooks in Excel through COM). This check caught a real bug: a single `n/a`
  cell at any age blanked the GER of every level through `SUMPRODUCT`. The formula now
  uses `SUMIFS` with an `ISNUMBER` count of official-age cells.
- **Provenance everywhere.** Every value keeps its file hash, sheet and cell, and
  every finding points to the cell to fix.

## Validation rules

18 rules with stable ids and severities (`python -m edutpl rules`). They cover
missing parameters, gaps in the latest year, text in numeric cells, negatives,
total ≠ female + male, sum by age ≠ total, enrolment above population at a given
age, **population age-cohort breaks** (pop(a+1, t+1) / pop(a, t)), year-on-year
jumps, implausible GER/GPI/PTR, expenditure inconsistencies, missing sources and
outdated template versions. Thresholds are in the spec, not in code.

The test suite injects 13 kinds of respondent error into synthetic submissions
(five random seeds). It checks that **every injected error is detected at the right
code and year**, and that **a clean submission produces zero findings**.

## Layout

```
config/     spec.yaml (structure, thresholds), labels.yaml (EN/FR/ES/PT/RU)
src/edutpl/ build.py (templates + formulas), reader.py, validate.py, indicators.py,
            consolidate.py (SQLite/CSV), report.py (HTML review, QC log), synth.py, cli.py
app/        Streamlit app: drop workbooks, see findings, download QC log and reports
r/          crosscheck_indicators.R (base R, no package needed)
scripts/    verify_formulas_excel.py (Windows + Excel)
tests/      pytest suite
```

Author: Abass Ndiaye. Licence: MIT.
