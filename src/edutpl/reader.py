"""Read a submitted template into tidy data, whatever its language.

Rows are recognised by their code in column A and years by the header row, so a
country that renames a sheet, hides columns or submits the Russian version is
read the same way. Every value keeps its provenance (sheet, cell) and any
reading problem (text in a numeric cell, formula typed by the respondent) is
recorded as a finding for the validation step.
"""
from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook

from .layout import FIRST_DATA_ROW, FIRST_YEAR_COL, HEADER_ROW, INPUT_SERIES, parse_code

_NUMBER = re.compile(r"^-?[\d\s  .,']+$")


@dataclass
class Submission:
    path: Path
    sha256: str
    template_version: str | None
    language: str | None
    observations: pd.DataFrame          # code, series, level, sex, age, year, value, raw, sheet, cell
    params: dict = field(default_factory=dict)
    metadata: dict = field(default_factory=dict)
    read_issues: list[dict] = field(default_factory=list)

    @property
    def country(self) -> str:
        return str(self.params.get("PAR|country_code||") or self.path.stem)


def parse_number(raw) -> tuple[float, str | None]:
    """Return (value, issue). Accepts '1 234', '1,234.5', '1.234,5'; flags them."""
    if raw is None or (isinstance(raw, str) and raw.strip() == ""):
        return np.nan, None
    if isinstance(raw, bool):
        return np.nan, "non_numeric"
    if isinstance(raw, (int, float)):
        return float(raw), None
    s = str(raw).strip()
    if not _NUMBER.match(s):
        return np.nan, "non_numeric"
    s = s.replace(" ", "").replace(" ", "").replace(" ", "").replace("'", "")
    if "," in s and "." in s:
        s = s.replace(",", "") if s.rfind(".") > s.rfind(",") else s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", "") if re.match(r"^-?\d{1,3}(,\d{3})+$", s) else s.replace(",", ".")
    try:
        return float(s), "number_as_text"
    except ValueError:
        return np.nan, "non_numeric"


def file_hash(path: Path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read_workbook(path: Path | str) -> Submission:
    path = Path(path)
    wb_f = load_workbook(path, data_only=False)   # to see formulas typed in input cells
    wb_v = load_workbook(path, data_only=True)    # cached values, if Excel saved them
    meta = {}
    if "_meta" in wb_v.sheetnames:
        for k, v in wb_v["_meta"].iter_rows(min_col=1, max_col=2, values_only=True):
            if k:
                meta[str(k)] = v

    records, issues, params, metadata = [], [], {}, {}
    for ws_v in wb_v.worksheets:
        if ws_v.title == "_meta":
            continue
        ws_f = wb_f[ws_v.title]
        years = {}
        for c in ws_v[HEADER_ROW][FIRST_YEAR_COL - 1:]:
            if isinstance(c.value, (int, float)) and 1900 < c.value < 2100:
                years[c.column] = int(c.value)
        for row in ws_v.iter_rows(min_row=FIRST_DATA_ROW):
            code = row[0].value
            if not isinstance(code, str):
                continue
            if code.startswith("PAR|"):
                params[code] = row[FIRST_YEAR_COL - 1].value
                continue
            if code.startswith("META|"):
                s = code.split("|")[1]
                metadata[s] = (row[FIRST_YEAR_COL - 1].value, row[FIRST_YEAR_COL].value)
                continue
            parts = parse_code(code)
            if not parts or parts["series"] not in INPUT_SERIES:
                continue  # calculated sheets (indicators, checks) are recomputed, not read
            for col, year in years.items():
                cell_v = row[col - 1]
                raw = cell_v.value
                f_raw = ws_f.cell(cell_v.row, col).value
                issue = None
                if isinstance(f_raw, str) and f_raw.startswith("="):
                    issue = "formula_in_input"
                value, parse_issue = parse_number(raw)
                issue = issue or parse_issue
                records.append({**parts, "code": code, "year": year, "value": value,
                                "raw": None if raw is None else str(raw),
                                "sheet": ws_v.title, "cell": cell_v.coordinate})
                if issue:
                    issues.append({"rule": issue, "sheet": ws_v.title, "cell": cell_v.coordinate,
                                   "code": code, "year": year, "raw": None if raw is None else str(raw)})
    obs = pd.DataFrame.from_records(records, columns=[
        "series", "level", "sex", "age", "code", "year", "value", "raw", "sheet", "cell"])
    obs["age"] = obs["age"].astype("Int64")
    for k in [k for k in params if "entry_age" in k or "duration" in k]:
        v, _ = parse_number(params[k])
        params[k] = None if np.isnan(v) else int(v)
    return Submission(path=path, sha256=file_hash(path), template_version=meta.get("template_version"),
                      language=meta.get("language"), observations=obs, params=params,
                      metadata=metadata, read_issues=issues)
