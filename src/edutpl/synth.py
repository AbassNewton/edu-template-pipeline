"""Synthetic country submissions with known, injected errors.

The data are invented (fictional countries XAA, XAB, ... using ISO user-assigned
codes) but internally consistent, so that a clean submission produces no
finding at all. Errors are then injected the way respondents make them, and the
list of injections is saved as ground truth: the test suite checks that the
validator catches every one of them.
"""
from __future__ import annotations

import json
from pathlib import Path

import numpy as np
from openpyxl import load_workbook

from .build import build_template
from .config import Spec, load_labels
from .layout import FIRST_DATA_ROW, FIRST_YEAR_COL, INPUT_SERIES, Layout

INJECTIONS = {   # kind -> rule expected to catch it
    "total_mismatch": "R07", "negative": "R06", "text_value": "R03", "number_as_text": "R04",
    "age_sum": "R08", "enrol_gt_pop": "R09", "cohort_break": "R10", "yoy_jump": "R11",
    "missing_last_year": "R02", "missing_param": "R01", "exp_inconsistent": "R15",
    "no_source": "R17", "old_version": "R18",
}
LANG_CYCLE = ("en", "fr", "es", "pt", "ru")
PARAM_SETS = [((6, 6), (12, 3), (15, 3)), ((7, 6), (13, 3), (16, 3)), ((6, 5), (11, 4), (15, 3))]
LEVEL_RATE = {"L1": 0.90, "L2": 0.65, "L3": 0.42}
PTR = {"L1": 38, "L2": 27, "L3": 21}


def country_data(spec: Spec, rng: np.random.Generator, code: str, name: str):
    """Return (values, params, metadata) for one consistent synthetic country."""
    params_set = PARAM_SETS[rng.integers(len(PARAM_SETS))]
    params = {"PAR|country_code||": code, "PAR|country_name||": name,
              "PAR|currency||": "Demo unit", "PAR|respondent||": "Statistics unit, Ministry of Education"}
    for lvl, (e, d) in zip(spec.levels, params_set):
        params[f"PAR|entry_age|{lvl}|"] = e
        params[f"PAR|duration|{lvl}|"] = d

    years, ages = spec.years, spec.ages
    base = rng.uniform(80_000, 600_000)
    pop = {}
    for sex, share in (("F", 0.49), ("M", 0.51)):
        for a in ages:
            pop[(sex, a, years[0])] = base * share * 0.98 ** (a - ages[0])
        for t in years[1:]:
            pop[(sex, ages[0], t)] = pop[(sex, ages[0], t - 1)] * 1.025
            for a in ages[1:]:
                pop[(sex, a, t)] = pop[(sex, a - 1, t - 1)] * 0.995
    pop = {k: round(v) for k, v in pop.items()}

    values = {(f"POP||{s}|{a}", t): v for (s, a, t), v in pop.items()}
    gpi = {lvl: rng.uniform(0.92, 1.03) for lvl in spec.levels}
    for li, lvl in enumerate(spec.levels):
        e, d = params_set[li]
        for t_i, t in enumerate(years):
            trend = 1 + 0.006 * t_i
            for sex in spec.sexes:
                sex_f = gpi[lvl] if sex == "F" else 1.0
                rates = {}
                for a in ages:
                    if e <= a < e + d:
                        rates[a] = LEVEL_RATE[lvl] * trend * sex_f
                    elif a == e - 1:
                        rates[a] = 0.02  # a few under-age pupils
                    elif e + d <= a < e + d + 3:
                        rates[a] = LEVEL_RATE[lvl] * 0.25 / (a - e - d + 1)
                by_age = {a: round(pop[(sex, a, t)] * min(r, 0.98)) for a, r in rates.items()}
                total = sum(by_age.values())
                if lvl in spec.age_levels:
                    for a, v in by_age.items():
                        values[(f"ENRA|{lvl}|{sex}|{a}", t)] = v
                values[(f"ENR|{lvl}|{sex}|", t)] = total
            values[(f"ENR|{lvl}|T|", t)] = sum(values[(f"ENR|{lvl}|{s}|", t)] for s in spec.sexes)
            tch_t = round(values[(f"ENR|{lvl}|T|", t)] / PTR[lvl])
            tch_f = round(tch_t * 0.55)
            values[(f"TCH|{lvl}|F|", t)], values[(f"TCH|{lvl}|M|", t)] = tch_f, tch_t - tch_f
            values[(f"TCH|{lvl}|T|", t)] = tch_t
    gdp0 = base * rng.uniform(4e4, 9e4)
    edu_share = rng.uniform(0.15, 0.2)
    for t_i, t in enumerate(years):
        gdp = gdp0 * 1.05 ** t_i
        values[("GDP|||", t)] = round(gdp)
        values[("GOV_TOTAL|||", t)] = round(gdp * 0.22)
        values[("GOV_EDU|||", t)] = round(gdp * 0.22 * edu_share * rng.uniform(0.97, 1.03))
    metadata = {s: ("Ministry of Education, administrative census" if s in ("ENR", "ENRA", "TCH")
                    else "National statistics office" if s == "POP" else "Ministry of Finance", None)
                for s in INPUT_SERIES}
    return values, params, metadata


def inject(path: Path, spec: Spec, lang: str, kind: str, values: dict, rng) -> dict:
    """Apply one error to a built workbook; return the ground-truth record."""
    lay, lab = Layout(spec), load_labels(lang)
    wb = load_workbook(path)
    years = spec.years
    y = int(rng.choice(years[1:-1]))
    rec = {"kind": kind, "rule": INJECTIONS[kind], "code": None, "year": None}

    def put(code, year, value):
        sheet, addr = lay.cell(code, year)
        wb[lab(f"sheet.{sheet}")][addr] = value
        rec["code"], rec["year"] = code, year

    if kind == "total_mismatch":
        put("ENR|L2|T|", y, values[("ENR|L2|T|", y)] + 137)
    elif kind == "negative":
        put("TCH|L3|F|", y, -values[("TCH|L3|F|", y)])
    elif kind == "text_value":
        put("POP||M|10", y, "n/a")
    elif kind == "number_as_text":
        put("ENR|L1|M|", y, f"{values[('ENR|L1|M|', y)]:,}".replace(",", " "))
    elif kind == "age_sum":
        put("ENRA|L1|F|8", y, values[("ENRA|L1|F|8", y)] + 500)
        rec["code"] = "ENR|L1|F|"
    elif kind == "enrol_gt_pop":
        put("ENRA|L1|M|8", y, round(values[("POP||M|8", y)] * 1.2))
        rec["code"] = "POP||M|8"
    elif kind == "cohort_break":
        put("POP||F|14", y, round(values[("POP||F|14", y)] * 1.5))
    elif kind == "yoy_jump":
        f, m = (round(values[(f"ENR|L3|{s}|", y)] * 1.6) for s in ("F", "M"))
        put("ENR|L3|F|", y, f)
        put("ENR|L3|M|", y, m)
        put("ENR|L3|T|", y, f + m)
    elif kind == "missing_last_year":
        put("GOV_EDU|||", years[-1], None)
    elif kind == "missing_param":
        wb[lab("sheet.parameters")][f"D{lay.param_row('PAR|entry_age|L3|')}"] = None
        rec["code"] = "PAR|entry_age|L3|"
    elif kind == "exp_inconsistent":
        put("GOV_EDU|||", y, round(values[("GOV_TOTAL|||", y)] * 1.1))
    elif kind == "no_source":
        i = INPUT_SERIES.index("TCH")
        wb[lab("sheet.metadata")].cell(FIRST_DATA_ROW + i, FIRST_YEAR_COL).value = None
        rec["code"] = "META|TCH||"
    elif kind == "old_version":
        wb["_meta"]["B1"] = "2026.1"
    wb.save(path)
    return rec


def generate(out_dir: Path | str, spec: Spec, n: int = 10, seed: int = 2027) -> list[dict]:
    """Write n submissions to out_dir; country 1 is clean, the others carry 1-4 errors."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    rng = np.random.default_rng(seed)
    kinds = list(INJECTIONS)
    truth = []
    for i in range(n):
        code = "X" + chr(65 + i // 26) + chr(65 + i % 26)
        lang = LANG_CYCLE[i % len(LANG_CYCLE)]
        values, params, metadata = country_data(spec, rng, code, f"Demo country {code}")
        path = build_template(out_dir / f"{code}_{lang}.xlsx", spec, lang, values, params, metadata)
        chosen = [] if i == 0 else rng.choice(kinds, size=int(rng.integers(1, 5)), replace=False)
        for kind in chosen:
            rec = inject(path, spec, lang, str(kind), values, rng)
            truth.append({"country": code, "file": path.name, **rec})
    (out_dir / "ground_truth.json").write_text(json.dumps(truth, indent=2), encoding="utf-8")
    return truth
