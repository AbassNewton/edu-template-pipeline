"""Generate a multi-year template (any language) from the spec, optionally pre-filled.

The workbook is rebuilt from the spec rather than patched: adding a year, a level
or a check means editing ``config/spec.yaml`` and regenerating, then pre-filling
with the data already held. This keeps every language version structurally
identical and makes the change reviewable in Git.
"""
from __future__ import annotations

from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from .config import Labels, Spec, load_labels
from .layout import (DATA_SHEETS, FIRST_DATA_ROW, FIRST_YEAR_COL, HEADER_ROW, INPUT_SERIES,
                     Layout, Row)

HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
CALC_FILL = PatternFill("solid", fgColor="E7E6E6")
LABEL_FILL = PatternFill("solid", fgColor="F2F2F2")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WHITE_BOLD = Font(bold=True, color="FFFFFF")


def q(sheet_name: str) -> str:
    """Quoted sheet reference for formulas (names are translated and may contain spaces)."""
    return "'" + sheet_name.replace("'", "''") + "'!"


def row_label(lab: Labels, r: Row) -> tuple[str, str | int]:
    """(series label, breakdown) shown in columns B and C."""
    series = lab(f"series.{r.series}")
    parts = []
    if r.level:
        parts.append(lab(f"level.{r.level}"))
    if r.sex:
        parts.append(lab(f"sex.{r.sex}"))
    label = series + (" - " + ", ".join(parts) if parts else "")
    return label, (r.age if r.age is not None else "")


class TemplateBuilder:
    def __init__(self, spec: Spec, lang: str = "en"):
        self.spec = spec
        self.lang = lang
        self.lab = load_labels(lang)
        self.layout = Layout(spec)
        self.names = {k: self.lab(f"sheet.{k}") for k in
                      ("instructions", "parameters", *DATA_SHEETS, "indicators", "checks", "metadata")}

    # ------------------------------------------------------------------ public
    def build(self, path: Path | str, values: dict | None = None, params: dict | None = None,
              metadata: dict | None = None) -> Path:
        """values: {(code, year): number}; params: {param code: value}; metadata: {series: (source, note)}."""
        values, params, metadata = values or {}, params or {}, metadata or {}
        wb = Workbook()
        wb.remove(wb.active)
        self._instructions(wb.create_sheet(self.names["instructions"]))
        self._parameters(wb.create_sheet(self.names["parameters"]), params)
        for key in DATA_SHEETS:
            self._data_sheet(wb.create_sheet(self.names[key]), key, values)
        self._indicators(wb.create_sheet(self.names["indicators"]))
        self._checks(wb.create_sheet(self.names["checks"]))
        self._metadata(wb.create_sheet(self.names["metadata"]), metadata)
        self._meta(wb.create_sheet("_meta"))
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(path)
        return path

    # ----------------------------------------------------------------- helpers
    def _header(self, ws, title: str, with_years: bool = True, extra: list[str] | None = None):
        ws["B1"] = title
        ws["B1"].font = Font(bold=True, size=13, color="1F4E79")
        ws["B2"] = self.lab("text.line1")
        ws["B2"].font = Font(italic=True, color="595959")
        heads = [self.lab("header.code"), self.lab("header.series"), self.lab("header.dimension")]
        if with_years:
            heads += [str(y) for y in self.spec.years]
        heads += extra or []
        for j, h in enumerate(heads, start=1):
            c = ws.cell(HEADER_ROW, j, int(h) if h.isdigit() else h)
            c.fill, c.font, c.border = HEADER_FILL, WHITE_BOLD, BORDER
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions["A"].hidden = True
        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 48
        ws.column_dimensions["C"].width = 12
        for j in range(FIRST_YEAR_COL, FIRST_YEAR_COL + len(heads) - 3):
            ws.column_dimensions[get_column_letter(j)].width = 12
        ws.freeze_panes = ws.cell(FIRST_DATA_ROW, FIRST_YEAR_COL)

    def _label_cells(self, ws, row: int, code: str, label: str, breakdown):
        ws.cell(row, 1, code)
        for j, v in ((2, label), (3, breakdown)):
            c = ws.cell(row, j, v)
            c.fill, c.border = LABEL_FILL, BORDER

    @staticmethod
    def _input(cell, value=None, fmt="#,##0"):
        if value is not None:
            cell.value = value
        cell.border = BORDER
        cell.number_format = fmt
        cell.protection = Protection(locked=False)

    @staticmethod
    def _calc(cell, formula, fmt="0.0"):
        cell.value = formula
        cell.fill, cell.border, cell.number_format = CALC_FILL, BORDER, fmt

    def _protect(self, ws):
        ws.protection.sheet = True
        ws.protection.formatColumns = False
        ws.protection.formatRows = False

    # ------------------------------------------------------------------ sheets
    def _instructions(self, ws):
        ws["B2"] = self.lab("text.title")
        ws["B2"].font = Font(bold=True, size=15, color="1F4E79")
        for i, key in enumerate(("line1", "line2", "line3"), start=4):
            ws.cell(i, 2, f"{i - 3}. " + self.lab(f"text.{key}"))
        ws["B8"] = f"{self.lab('text.version')}: {self.spec.template_version}"
        ws["B9"] = f"{self.lab('text.language')}: {self.lang.upper()}"
        ws["B10"] = f"{self.spec.first_year}-{self.spec.last_year}"
        ws.column_dimensions["A"].width = 3
        ws.column_dimensions["B"].width = 110

    def _parameters(self, ws, params):
        self._header(ws, self.names["parameters"], with_years=False, extra=[self.lab("header.value")])
        for i, (code, key, lvl) in enumerate(self.layout.param_rows()):
            r = FIRST_DATA_ROW + i
            self._label_cells(ws, r, code, self.lab(key), self.lab(f"level.{lvl}") if lvl else "")
            self._input(ws.cell(r, FIRST_YEAR_COL), params.get(code), fmt="General")
        ws.column_dimensions["C"].width = 30
        ws.column_dimensions["D"].width = 30
        self._protect(ws)

    def _data_sheet(self, ws, key, values):
        self._header(ws, self.names[key])
        dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0",
                            allow_blank=True, showErrorMessage=True, errorStyle="warning",
                            errorTitle=self.lab("check.review"), error=self.lab("text.line2"))
        ws.add_data_validation(dv)
        for i, r in enumerate(self.layout.rows[key]):
            row = FIRST_DATA_ROW + i
            label, breakdown = row_label(self.lab, r)
            self._label_cells(ws, row, r.code, label, breakdown)
            for y in self.spec.years:
                self._input(ws.cell(row, self.layout.year_col(y)), values.get((r.code, y)))
        last = FIRST_DATA_ROW + len(self.layout.rows[key]) - 1
        dv.add(f"{get_column_letter(FIRST_YEAR_COL)}{FIRST_DATA_ROW}:"
               f"{get_column_letter(self.layout.last_col)}{last}")
        self._protect(ws)

    # ---- formulas ---------------------------------------------------------
    def _ref(self, code: str, col: str, absolute_col=False) -> str:
        sheet, row = self.layout.locate(code)
        c = f"${col}" if absolute_col else col
        return f"{q(self.names[sheet])}{c}{row}"

    def _param_ref(self, code: str) -> str:
        return f"{q(self.names['parameters'])}$D${self.layout.param_row(code)}"

    def _pop_terms(self, sex: str, col: str, e: str, d: str) -> tuple[str, str]:
        r0, r1 = self.layout.block("population", lambda r: r.sex == sex)
        p = q(self.names["population"])
        ages, vals = f"{p}$C${r0}:$C${r1}", f"{p}{col}${r0}:{col}${r1}"
        # SUMIFS ignores text cells outside the official age range (a SUMPRODUCT over
        # the whole column would turn one 'n/a' at any age into #VALUE!); the second
        # term counts official-age cells that hold a number, so that a missing or
        # non-numeric official-age cell makes the indicator missing, as in Python.
        total = f'SUMIFS({vals},{ages},">="&{e},{ages},"<"&({e}+{d}))'
        n_num = f"SUMPRODUCT(({ages}>={e})*({ages}<{e}+{d})*ISNUMBER({vals}))"
        return total, n_num

    def ger_formula(self, lvl: str, sex: str, col: str) -> str:
        num = self._ref(f"ENR|{lvl}|{sex}|", col)
        e, d = self._param_ref(f"PAR|entry_age|{lvl}|"), self._param_ref(f"PAR|duration|{lvl}|")
        sexes = self.spec.sexes if sex == "T" else [sex]
        terms = [self._pop_terms(s, col, e, d) for s in sexes]
        den = "+".join(t[0] for t in terms)
        n_num = "+".join(t[1] for t in terms)
        return (f'=IF(OR(ISBLANK({num}),ISBLANK({e}),ISBLANK({d})),"",'
                f'IFERROR(IF(({n_num})<{len(sexes)}*{d},"",{num}/({den})*100),""))')

    @staticmethod
    def ratio_formula(a: str, b: str, scale: int = 1) -> str:
        mult = f"*{scale}" if scale != 1 else ""
        return f'=IF(OR(ISBLANK({a}),ISBLANK({b})),"",IFERROR({a}/{b}{mult},""))'

    def _indicators(self, ws):
        self._header(ws, self.names["indicators"])
        lay = self.layout
        for i, r in enumerate(lay.rows["indicators"]):
            row = FIRST_DATA_ROW + i
            label, _ = row_label(self.lab, r)
            self._label_cells(ws, row, r.code, label, "")
            for y in self.spec.years:
                col = lay.year_letter(y)
                if r.series == "GER":
                    f = self.ger_formula(r.level, r.sex, col)
                    fmt = "0.0"
                elif r.series == "GPI":
                    f_row, m_row = (lay.locate(f"GER|{r.level}|{s}|")[1] for s in ("F", "M"))
                    f = f'=IFERROR({col}{f_row}/{col}{m_row},"")'
                    fmt = "0.00"
                elif r.series == "PTR":
                    f = self.ratio_formula(self._ref(f"ENR|{r.level}|T|", col),
                                           self._ref(f"TCH|{r.level}|T|", col))
                    fmt = "0.0"
                elif r.series == "XGOV":
                    f = self.ratio_formula(self._ref("GOV_EDU|||", col), self._ref("GOV_TOTAL|||", col), 100)
                    fmt = "0.0"
                else:  # XGDP
                    f = self.ratio_formula(self._ref("GOV_EDU|||", col), self._ref("GDP|||", col), 100)
                    fmt = "0.00"
                self._calc(ws.cell(row, lay.year_col(y)), f, fmt)
        self._protect(ws)

    def _checks(self, ws):
        status_col = self.layout.last_col + 1
        self._header(ws, self.names["checks"], extra=[self.lab("header.status")])
        lay = self.layout
        first_l, last_l = get_column_letter(FIRST_YEAR_COL), get_column_letter(lay.last_col)
        for i, r in enumerate(lay.rows["checks"]):
            row = FIRST_DATA_ROW + i
            lvl_label = self.lab(f"level.{r.level}") if r.level else ""
            self._label_cells(ws, row, r.code, self.lab(f"check.{r.series}"), lvl_label)
            for y in self.spec.years:
                col = lay.year_letter(y)
                if r.series == "total_vs_sum":
                    fv, mv, tv = (self._ref(f"ENR|{r.level}|{s}|", col) for s in ("F", "M", "T"))
                    f = f'=IF(COUNT({fv},{mv},{tv})<3,"",{tv}-({fv}+{mv}))'
                elif r.series == "age_vs_total":
                    a = q(self.names["enrolment_age"])
                    rngs = []
                    for s in self.spec.sexes:
                        r0, r1 = lay.block("enrolment_age", lambda x, s=s: x.level == r.level and x.sex == s)
                        rngs.append(f"{a}{col}{r0}:{col}{r1}")
                    tv = self._ref(f"ENR|{r.level}|T|", col)
                    f = (f'=IF(OR(COUNT({",".join(rngs)})=0,COUNT({tv})=0),"",'
                         f'SUM({",".join(rngs)})-{tv})')
                else:  # negatives
                    parts = []
                    for key in DATA_SHEETS:
                        n = len(lay.rows[key])
                        parts.append(f'COUNTIF({q(self.names[key])}{col}{FIRST_DATA_ROW}:'
                                     f'{col}{FIRST_DATA_ROW + n - 1},"<0")')
                    f = "=" + "+".join(parts)
                self._calc(ws.cell(row, lay.year_col(y)), f, "#,##0")
            rng = f"{first_l}{row}:{last_l}{row}"
            self._calc(ws.cell(row, status_col),
                       f'=IF(SUMPRODUCT(({rng}<>"")*({rng}<>0))>0,"{self.lab("check.review")}",'
                       f'"{self.lab("check.ok")}")', "General")
        self._protect(ws)

    def _metadata(self, ws, metadata):
        self._header(ws, self.names["metadata"], with_years=False,
                     extra=[self.lab("header.source"), self.lab("header.note")])
        for i, s in enumerate(INPUT_SERIES):
            row = FIRST_DATA_ROW + i
            self._label_cells(ws, row, f"META|{s}||", self.lab(f"series.{s}"), "")
            src, note = metadata.get(s, (None, None))
            self._input(ws.cell(row, FIRST_YEAR_COL), src, fmt="@")
            self._input(ws.cell(row, FIRST_YEAR_COL + 1), note, fmt="@")
        ws.column_dimensions["D"].width = 40
        ws.column_dimensions["E"].width = 60
        self._protect(ws)

    def _meta(self, ws):
        rows = [("template_version", self.spec.template_version), ("language", self.lang),
                ("first_year", self.spec.first_year), ("last_year", self.spec.last_year),
                ("generated", date.today().isoformat()), ("generator", "edutpl")]
        for i, (k, v) in enumerate(rows, start=1):
            ws.cell(i, 1, k)
            ws.cell(i, 2, v)
        ws.sheet_state = "hidden"


def build_template(path, spec: Spec, lang: str = "en", values=None, params=None, metadata=None) -> Path:
    return TemplateBuilder(spec, lang).build(path, values, params, metadata)


def extend_template(src: Path | str, dst: Path | str, spec: Spec, lang: str | None = None) -> Path:
    """Re-issue an existing (possibly filled) workbook under a new spec, e.g. with new years.

    Data, parameters and metadata are carried over by code, so the operation works
    across languages (read a French submission, issue a Spanish template).
    """
    from .reader import read_workbook

    sub = read_workbook(src)
    values = {(o.code, o.year): o.value for o in sub.observations.itertuples()
              if o.year in spec.years and o.value == o.value}  # drop NaN
    return build_template(dst, spec, lang or sub.language, values, sub.params, sub.metadata)
