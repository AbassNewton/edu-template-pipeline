"""Indicator calculations, written to match the template formulas exactly.

Keeping an independent implementation in code (and a third one in R, see
``r/crosscheck_indicators.R``) is what lets us verify that the Excel formulas
follow the methodology: the three must agree on every submission.

Definitions (simplified for the demo):
  GER  = enrolment at level (all ages) / population of official age for the level * 100
         -> missing if any official-age population cell is missing
  GPI  = GER female / GER male
  PTR  = enrolment total / teachers total
  XGOV = government expenditure on education / total government expenditure * 100
  XGDP = government expenditure on education / GDP * 100
"""
from __future__ import annotations

import numpy as np
import pandas as pd

from .config import Spec
from .reader import Submission


def _div(a, b, scale=1.0):
    if a is None or b is None or np.isnan(a) or np.isnan(b) or b == 0:
        return np.nan
    return a / b * scale


def compute_indicators(sub: Submission, spec: Spec) -> pd.DataFrame:
    obs = sub.observations
    val = {(r.code, r.year): r.value for r in obs.itertuples()}
    pop = obs[obs.series == "POP"]
    out = []
    for year in spec.years:
        ger = {}
        for lvl in spec.levels:
            e = sub.params.get(f"PAR|entry_age|{lvl}|")
            d = sub.params.get(f"PAR|duration|{lvl}|")
            for sex in (*spec.sexes, "T"):
                num = val.get((f"ENR|{lvl}|{sex}|", year), np.nan)
                g = np.nan
                if e is not None and d is not None and not np.isnan(num):
                    sexes = spec.sexes if sex == "T" else [sex]
                    p = pop[(pop.year == year) & pop.sex.isin(sexes)
                            & (pop.age >= e) & (pop.age < e + d)]
                    expected = len(sexes) * d
                    if len(p) == expected and p.value.notna().all():
                        g = _div(num, p.value.sum(), 100)
                ger[(lvl, sex)] = g
                out.append(("GER", lvl, sex, year, g))
            out.append(("GPI", lvl, None, year, _div(ger[(lvl, "F")], ger[(lvl, "M")])))
            out.append(("PTR", lvl, "T", year, _div(val.get((f"ENR|{lvl}|T|", year), np.nan),
                                                   val.get((f"TCH|{lvl}|T|", year), np.nan))))
        edu = val.get(("GOV_EDU|||", year), np.nan)
        out.append(("XGOV", None, None, year, _div(edu, val.get(("GOV_TOTAL|||", year), np.nan), 100)))
        out.append(("XGDP", None, None, year, _div(edu, val.get(("GDP|||", year), np.nan), 100)))
    df = pd.DataFrame(out, columns=["indicator", "level", "sex", "year", "value"])
    df.insert(0, "country", sub.country)
    return df
