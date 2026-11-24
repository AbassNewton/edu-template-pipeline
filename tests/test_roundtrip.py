import numpy as np
import pytest

from edutpl.build import build_template, extend_template
from edutpl.config import LANGUAGES
from edutpl.reader import parse_number, read_workbook
from edutpl.synth import country_data


@pytest.fixture(scope="module")
def data(spec):
    return country_data(spec, np.random.default_rng(1), "XZZ", "Test")


@pytest.mark.parametrize("lang", LANGUAGES)
def test_build_then_read_returns_same_values(tmp_path, spec, data, lang):
    values, params, meta = data
    sub = read_workbook(build_template(tmp_path / f"t_{lang}.xlsx", spec, lang, values, params, meta))
    got = {(r.code, r.year): r.value for r in sub.observations.itertuples() if not np.isnan(r.value)}
    assert got == {k: float(v) for k, v in values.items()}
    assert sub.language == lang and sub.country == "XZZ"
    assert sub.params["PAR|entry_age|L1|"] == params["PAR|entry_age|L1|"]
    assert sub.read_issues == []


def test_extend_adds_year_and_keeps_data(tmp_path, spec, data):
    values, params, meta = data
    src = build_template(tmp_path / "old.xlsx", spec, "fr", values, params, meta)
    new = spec.with_years(last_year=spec.last_year + 1, version="2028.1")
    sub = read_workbook(extend_template(src, tmp_path / "new.xlsx", new, lang="es"))
    assert sub.template_version == "2028.1" and sub.language == "es"
    assert spec.last_year + 1 in set(sub.observations.year)
    assert sub.observations.query("year == @spec.last_year + 1").value.isna().all()
    old = {(r.code, r.year): r.value for r in sub.observations.itertuples()
           if r.year <= spec.last_year and not np.isnan(r.value)}
    assert old == {k: float(v) for k, v in values.items()}


@pytest.mark.parametrize("raw,expected,issue", [
    (1234, 1234.0, None), ("1 234", 1234.0, "number_as_text"), ("1,234", 1234.0, "number_as_text"),
    ("1.234,5", 1234.5, "number_as_text"), ("12,5", 12.5, "number_as_text"),
    ("n/a", np.nan, "non_numeric"), ("-", np.nan, "non_numeric"), (None, np.nan, None)])
def test_parse_number(raw, expected, issue):
    v, i = parse_number(raw)
    assert i == issue
    assert (np.isnan(v) and np.isnan(expected)) or v == expected
