"""Workbook layout derived from the spec.

Every data row carries a language-independent code in (hidden) column A:
``SERIES|LEVEL|SEX|AGE`` (empty parts allowed). Year columns start at column D,
headers on row 4, data from row 5. The same layout object is used to *write*
templates, to *read* submissions and to point validation findings at a cell,
so the three can never drift apart.
"""
from __future__ import annotations

from dataclasses import dataclass

from openpyxl.utils import get_column_letter

from .config import Spec

HEADER_ROW = 4
FIRST_DATA_ROW = 5
FIRST_YEAR_COL = 4  # column D

DATA_SHEETS = ("enrolment", "enrolment_age", "population", "teachers", "expenditure")
SHEET_ORDER = ("instructions", "parameters", *DATA_SHEETS, "indicators", "checks", "metadata")
INPUT_SERIES = ("ENR", "ENRA", "POP", "TCH", "GOV_TOTAL", "GOV_EDU", "GDP")


def make_code(series: str, level: str = "", sex: str = "", age: int | str = "") -> str:
    return f"{series}|{level}|{sex}|{age}"


def parse_code(code: str) -> dict | None:
    parts = str(code).split("|")
    if len(parts) != 4:
        return None
    series, level, sex, age = parts
    return {"series": series, "level": level or None, "sex": sex or None,
            "age": int(age) if age.isdigit() else None}


@dataclass(frozen=True)
class Row:
    code: str
    series: str
    level: str | None
    sex: str | None
    age: int | None


class Layout:
    def __init__(self, spec: Spec):
        self.spec = spec
        self.rows: dict[str, list[Row]] = {k: [] for k in (*DATA_SHEETS, "indicators", "checks")}
        sp = spec
        for lvl in sp.levels:
            for sex in (*sp.sexes, "T"):
                self._add("enrolment", "ENR", lvl, sex)
        for lvl in sp.age_levels:
            for sex in sp.sexes:
                for age in sp.ages:
                    self._add("enrolment_age", "ENRA", lvl, sex, age)
        for sex in sp.sexes:
            for age in sp.ages:
                self._add("population", "POP", None, sex, age)
        for lvl in sp.levels:
            for sex in (*sp.sexes, "T"):
                self._add("teachers", "TCH", lvl, sex)
        for item in sp.expenditure_items:
            self._add("expenditure", item)
        for lvl in sp.levels:
            for sex in (*sp.sexes, "T"):
                self._add("indicators", "GER", lvl, sex)
            self._add("indicators", "GPI", lvl)
            self._add("indicators", "PTR", lvl, "T")
        self._add("indicators", "XGOV")
        self._add("indicators", "XGDP")
        for lvl in sp.levels:
            self._add("checks", "total_vs_sum", lvl)
        for lvl in sp.age_levels:
            self._add("checks", "age_vs_total", lvl)
        self._add("checks", "negatives")

        self._row_index: dict[str, tuple[str, int]] = {}
        for sheet, rows in self.rows.items():
            for i, r in enumerate(rows):
                self._row_index[r.code] = (sheet, FIRST_DATA_ROW + i)

    def _add(self, sheet, series, level=None, sex=None, age=None):
        self.rows[sheet].append(Row(make_code(series, level or "", sex or "", "" if age is None else age),
                                    series, level, sex, age))

    # ---- addressing -------------------------------------------------------
    def year_col(self, year: int) -> int:
        return FIRST_YEAR_COL + self.spec.years.index(year)

    def year_letter(self, year: int) -> str:
        return get_column_letter(self.year_col(year))

    @property
    def last_col(self) -> int:
        return FIRST_YEAR_COL + len(self.spec.years) - 1

    def locate(self, code: str) -> tuple[str, int]:
        """(sheet key, row number) of a code."""
        return self._row_index[code]

    def cell(self, code: str, year: int) -> tuple[str, str]:
        sheet, row = self.locate(code)
        return sheet, f"{self.year_letter(year)}{row}"

    def block(self, sheet: str, predicate) -> tuple[int, int]:
        """First and last row of a contiguous block of rows matching predicate."""
        rows = [FIRST_DATA_ROW + i for i, r in enumerate(self.rows[sheet]) if predicate(r)]
        if not rows or rows[-1] - rows[0] + 1 != len(rows):
            raise ValueError(f"block in {sheet} is empty or not contiguous")
        return rows[0], rows[-1]

    # Parameters sheet: fixed rows
    def param_rows(self) -> list[tuple[str, str, str | None]]:
        """(code, label key, level) for every parameter row, in order."""
        out = [(f"PAR|country_code||", "param.country_code", None),
               (f"PAR|country_name||", "param.country_name", None),
               (f"PAR|currency||", "param.currency", None),
               (f"PAR|respondent||", "param.respondent", None)]
        for lvl in self.spec.levels:
            out.append((f"PAR|entry_age|{lvl}|", "param.entry_age", lvl))
            out.append((f"PAR|duration|{lvl}|", "param.duration", lvl))
        return out

    def param_row(self, code: str) -> int:
        for i, (c, _, _) in enumerate(self.param_rows()):
            if c == code:
                return FIRST_DATA_ROW + i
        raise KeyError(code)
