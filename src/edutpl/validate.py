"""Validation rules applied to a submission.

Each rule has a stable id, a severity and a one-line description; the registry
is exported as the quality-control checklist so that reviewers, countries and
code all refer to the same list. Every finding points to a sheet and a cell.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass

import numpy as np
import pandas as pd

from .config import Spec
from .indicators import compute_indicators
from .reader import Submission

RULES = {
    "R01": ("error", "Required parameter missing (official entry age, duration, country code)"),
    "R02": ("warning", "Series reported in the previous year but missing in the latest year"),
    "R03": ("error", "Non-numeric value in a numeric cell"),
    "R04": ("warning", "Number stored as text (read, but should be entered as a number)"),
    "R05": ("info", "Formula typed in an input cell (cached value used)"),
    "R06": ("error", "Negative value"),
    "R07": ("error", "Total differs from female + male"),
    "R08": ("error", "Sum of enrolment by age differs from enrolment total (by sex)"),
    "R09": ("error", "Enrolment at a given age exceeds population of that age"),
    "R10": ("warning", "Population age-cohort break: pop(age+1, year+1) / pop(age, year) outside range"),
    "R11": ("warning", "Year-on-year change above threshold"),
    "R12": ("warning", "Gross enrolment ratio above maximum plausible value"),
    "R13": ("warning", "Gender parity index outside plausible range"),
    "R14": ("warning", "Pupil-teacher ratio outside plausible range"),
    "R15": ("error", "Expenditure inconsistency (education > total government, or government > GDP)"),
    "R16": ("warning", "Official entry ages are not consistent across levels"),
    "R17": ("warning", "Data reported without a source in Metadata"),
    "R18": ("warning", "Template version differs from the current version"),
}
READ_RULE = {"non_numeric": "R03", "number_as_text": "R04", "formula_in_input": "R05"}


@dataclass
class Finding:
    country: str
    rule: str
    severity: str
    sheet: str | None
    cell: str | None
    year: int | None
    code: str | None
    message: str


def rules_table() -> pd.DataFrame:
    return pd.DataFrame([(k, s, d) for k, (s, d) in RULES.items()],
                        columns=["rule", "severity", "description"])


class Validator:
    def __init__(self, sub: Submission, spec: Spec):
        self.sub, self.spec, self.th = sub, spec, spec.thresholds
        self.obs = sub.observations
        self.years = sorted(self.obs.year.unique().tolist()) if len(self.obs) else []
        self.val = {(r.code, r.year): r.value for r in self.obs.itertuples()}
        self.loc = {(r.code, r.year): (r.sheet, r.cell) for r in self.obs.itertuples()}
        self.findings: list[Finding] = []

    def add(self, rule, msg, code=None, year=None, sheet=None, cell=None):
        if code is not None and year is not None and sheet is None:
            sheet, cell = self.loc.get((code, year), (None, None))
        self.findings.append(Finding(self.sub.country, rule, RULES[rule][0], sheet, cell,
                                     None if year is None else int(year), code, msg))

    def v(self, code, year):
        return self.val.get((code, year), np.nan)

    # ------------------------------------------------------------------ rules
    def run(self) -> list[Finding]:
        for fn in (self.r_params, self.r_reading, self.r_completeness, self.r_negative,
                   self.r_totals, self.r_age_sums, self.r_enrol_vs_pop, self.r_cohort,
                   self.r_yoy, self.r_expenditure, self.r_indicators, self.r_metadata,
                   self.r_version):
            fn()
        return self.findings

    def r_params(self):
        p = self.sub.params
        if not p.get("PAR|country_code||"):
            self.add("R01", "Country code is empty", code="PAR|country_code||")
        for lvl in self.spec.levels:
            for k in ("entry_age", "duration"):
                if p.get(f"PAR|{k}|{lvl}|") is None:
                    self.add("R01", f"{k} missing for {lvl}", code=f"PAR|{k}|{lvl}|")
        lv = self.spec.levels
        for a, b in zip(lv, lv[1:]):
            ea, da, eb = (p.get(f"PAR|entry_age|{a}|"), p.get(f"PAR|duration|{a}|"),
                          p.get(f"PAR|entry_age|{b}|"))
            if None not in (ea, da, eb) and ea + da != eb:
                self.add("R16", f"{a} entry {ea} + duration {da} != {b} entry {eb}",
                         code=f"PAR|entry_age|{b}|")

    def r_reading(self):
        for i in self.sub.read_issues:
            self.add(READ_RULE[i["rule"]], f"Cell content: {i['raw']!r}", code=i["code"],
                     year=i["year"], sheet=i["sheet"], cell=i["cell"])

    def r_completeness(self):
        if len(self.years) < 2:
            return
        last, prev = self.years[-1], self.years[-2]
        for code in self.obs.code.unique():
            if not np.isnan(self.v(code, prev)) and np.isnan(self.v(code, last)):
                self.add("R02", f"Reported for {prev}, empty for {last}", code=code, year=last)

    def r_negative(self):
        for r in self.obs[self.obs.value < 0].itertuples():
            self.add("R06", f"Value {r.value:,.0f}", r.code, r.year, r.sheet, r.cell)

    def r_totals(self):
        tol = self.th.get("total_tolerance", 0)
        for series in ("ENR", "TCH"):
            for lvl in self.spec.levels:
                for y in self.years:
                    f, m, t = (self.v(f"{series}|{lvl}|{s}|", y) for s in ("F", "M", "T"))
                    if not np.isnan([f, m, t]).any() and abs(t - (f + m)) > tol:
                        self.add("R07", f"{series} {lvl}: total {t:,.0f} vs F+M {f + m:,.0f}",
                                 f"{series}|{lvl}|T|", y)

    def r_age_sums(self):
        enra = self.obs[self.obs.series == "ENRA"]
        for lvl in self.spec.age_levels:
            for sex in self.spec.sexes:
                sub = enra[(enra.level == lvl) & (enra.sex == sex)]
                for y, g in sub.groupby("year"):
                    if g.value.notna().sum() == 0:
                        continue
                    tot = self.v(f"ENR|{lvl}|{sex}|", y)
                    s = g.value.sum()
                    if not np.isnan(tot) and abs(s - tot) > self.th.get("total_tolerance", 0):
                        self.add("R08", f"{lvl} {sex}: sum by age {s:,.0f} vs total {tot:,.0f}",
                                 f"ENR|{lvl}|{sex}|", y)

    def r_enrol_vs_pop(self):
        enra = self.obs[self.obs.series == "ENRA"].groupby(["sex", "age", "year"]).value.sum(min_count=1)
        for (sex, age, y), e in enra.items():
            p = self.v(f"POP||{sex}|{age}", y)
            if not np.isnan(e) and not np.isnan(p) and e > p:
                self.add("R09", f"Age {age} {sex}: enrolment {e:,.0f} > population {p:,.0f}",
                         f"POP||{sex}|{age}", y)

    def r_cohort(self):
        lo, hi = self.th.get("cohort_ratio", [0.85, 1.15])
        for sex in self.spec.sexes:
            for age in self.spec.ages[:-1]:
                for y in self.years[:-1]:
                    a, b = self.v(f"POP||{sex}|{age}", y), self.v(f"POP||{sex}|{age + 1}", y + 1)
                    if a and b and not np.isnan(a) and not np.isnan(b) and a > 0:
                        r = b / a
                        if not lo <= r <= hi:
                            self.add("R10", f"{sex} age {age}->{age + 1}, {y}->{y + 1}: ratio {r:.2f}",
                                     f"POP||{sex}|{age + 1}", y + 1)

    def r_yoy(self):
        th = self.th.get("yoy_change_warning", 0.3)
        codes = [f"ENR|{l}|T|" for l in self.spec.levels] + [f"TCH|{l}|T|" for l in self.spec.levels] + ["GOV_EDU|||"]
        for code in codes:
            for y0, y1 in zip(self.years, self.years[1:]):
                a, b = self.v(code, y0), self.v(code, y1)
                if not np.isnan(a) and not np.isnan(b) and a > 0 and abs(b / a - 1) > th:
                    self.add("R11", f"{y0}->{y1}: {a:,.0f} -> {b:,.0f} ({b / a - 1:+.0%})", code, y1)

    def r_expenditure(self):
        for y in self.years:
            edu, tot, gdp = self.v("GOV_EDU|||", y), self.v("GOV_TOTAL|||", y), self.v("GDP|||", y)
            if not np.isnan(edu) and not np.isnan(tot) and edu > tot:
                self.add("R15", f"Education {edu:,.0f} > total government {tot:,.0f}", "GOV_EDU|||", y)
            if not np.isnan(tot) and not np.isnan(gdp) and tot > gdp:
                self.add("R15", f"Total government {tot:,.0f} > GDP {gdp:,.0f}", "GOV_TOTAL|||", y)

    def r_indicators(self):
        ind = compute_indicators(self.sub, self.spec)
        lo_g, hi_g = self.th.get("gpi_range", [0.5, 1.5])
        lo_p, hi_p = self.th.get("ptr_range", [5, 100])
        for r in ind.dropna(subset=["value"]).itertuples():
            if r.indicator == "GER" and r.value > self.th.get("ger_max", 150):
                self.add("R12", f"GER {r.level} {r.sex} = {r.value:.1f}", f"ENR|{r.level}|{r.sex}|", r.year)
            elif r.indicator == "GPI" and not lo_g <= r.value <= hi_g:
                self.add("R13", f"GPI {r.level} = {r.value:.2f}", f"ENR|{r.level}|F|", r.year)
            elif r.indicator == "PTR" and not lo_p <= r.value <= hi_p:
                self.add("R14", f"PTR {r.level} = {r.value:.1f}", f"TCH|{r.level}|T|", r.year)

    def r_metadata(self):
        reported = set(self.obs.loc[self.obs.value.notna(), "series"])
        for s in sorted(reported):
            src = (self.sub.metadata.get(s) or (None, None))[0]
            if not src:
                self.add("R17", f"No source given for {s}", code=f"META|{s}||")

    def r_version(self):
        if self.sub.template_version != self.spec.template_version:
            self.add("R18", f"Submitted {self.sub.template_version}, current {self.spec.template_version}")


def validate(sub: Submission, spec: Spec) -> pd.DataFrame:
    rows = [asdict(f) for f in Validator(sub, spec).run()]
    cols = list(Finding.__dataclass_fields__)
    return pd.DataFrame(rows, columns=cols)


def checklist(findings: pd.DataFrame, country: str) -> pd.DataFrame:
    """Quality-control checklist: one line per rule, with pass/fail and count."""
    t = rules_table()
    counts = findings.groupby("rule").size() if len(findings) else pd.Series(dtype=int)
    t["findings"] = t.rule.map(counts).fillna(0).astype(int)
    t["status"] = np.where(t.findings == 0, "pass",
                           np.where(t.severity == "error", "fail", "review"))
    t.insert(0, "country", country)
    return t
