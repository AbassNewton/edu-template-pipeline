import json

import pandas as pd
import pytest

from edutpl.indicators import compute_indicators
from edutpl.reader import read_workbook
from edutpl.synth import generate
from edutpl.validate import RULES, validate


@pytest.mark.parametrize("seed", [1, 2, 3, 4, 5])
def test_every_injected_error_is_caught_and_clean_file_is_clean(tmp_path, spec, seed):
    truth = pd.DataFrame(generate(tmp_path, spec, n=8, seed=seed))
    found = []
    for p in sorted(tmp_path.glob("*.xlsx")):
        sub = read_workbook(p)
        found.append(validate(sub, spec))
    f = pd.concat([x for x in found if len(x)])
    # clean country (first file): no finding at all, i.e. no false positive
    assert f[f.country == "XAA"].empty
    for r in truth.itertuples():
        hit = f[(f.country == r.country) & (f.rule == r.rule)]
        assert not hit.empty, f"{r.kind} not detected in {r.country}"
        if isinstance(r.code, str) and r.code and not pd.isna(r.year):
            assert ((hit.code == r.code) & (hit.year == r.year)).any(), (r.kind, hit)


def test_findings_point_to_a_cell(tmp_path, spec):
    generate(tmp_path, spec, n=6, seed=9)
    f = pd.concat([x for x in (validate(read_workbook(p), spec) for p in tmp_path.glob("*.xlsx")) if len(x)])
    cell_rules = f[f.rule.isin(["R03", "R04", "R06", "R07", "R08", "R09", "R10", "R11", "R15"])]
    assert cell_rules.cell.notna().all()
    assert set(f.rule) <= set(RULES)


def test_ger_matches_hand_calculation(tmp_path, spec):
    generate(tmp_path, spec, n=1, seed=3)
    sub = read_workbook(next(tmp_path.glob("*.xlsx")))
    ind = compute_indicators(sub, spec)
    y = spec.last_year
    e, d = sub.params["PAR|entry_age|L1|"], sub.params["PAR|duration|L1|"]
    o = sub.observations
    pop = o[(o.series == "POP") & (o.year == y) & (o.age >= e) & (o.age < e + d)].value.sum()
    enr = o[(o.code == "ENR|L1|T|") & (o.year == y)].value.item()
    got = ind.query("indicator == 'GER' and level == 'L1' and sex == 'T' and year == @y").value.item()
    assert got == pytest.approx(enr / pop * 100)
