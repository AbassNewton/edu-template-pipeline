"""Streamlit front end for non-programmers: drop workbooks, get findings and the QC log.

    streamlit run app/streamlit_app.py
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from edutpl.build import build_template  # noqa: E402
from edutpl.config import LANGUAGES, load_spec  # noqa: E402
from edutpl.consolidate import process_folder  # noqa: E402
from edutpl.synth import generate  # noqa: E402
from edutpl.validate import rules_table  # noqa: E402

st.set_page_config(page_title="Education template checker", layout="wide")
spec = load_spec()

st.title("Education template checker")
st.caption(f"Template version {spec.template_version} · years {spec.first_year}-{spec.last_year} · "
           "demo on synthetic data, not an official tool")

tab_check, tab_tpl, tab_rules = st.tabs(["Check submissions", "Blank templates", "Validation rules"])


@st.cache_data(show_spinner="Reading and validating workbooks...")
def run(files: tuple[tuple[str, bytes], ...]):
    with tempfile.TemporaryDirectory() as tmp:
        inp, out = Path(tmp) / "in", Path(tmp) / "out"
        inp.mkdir()
        for name, data in files:
            (inp / name).write_bytes(data)
        t = process_folder(inp, out, spec, reports=True)
        reports = {p.stem: p.read_text(encoding="utf-8") for p in (out / "reports").glob("*.html")}
        return t, (out / "qc_log.xlsx").read_bytes(), reports


@st.cache_data
def demo_files():
    with tempfile.TemporaryDirectory() as tmp:
        generate(tmp, spec, n=10, seed=2027)
        return tuple((p.name, p.read_bytes()) for p in sorted(Path(tmp).glob("*.xlsx")))


with tab_check:
    up = st.file_uploader("Country submissions (.xlsx)", type="xlsx", accept_multiple_files=True)
    use_demo = st.toggle("Use the 10 synthetic demo submissions", value=not up)
    files = demo_files() if use_demo and not up else tuple((f.name, f.getvalue()) for f in up or [])
    if files:
        t, qc_bytes, reports = run(files)
        subs, f, ind = t["submissions"], t["findings"], t["indicators"]
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Submissions", len(subs))
        c2.metric("Accepted", int((subs.status == "accepted").sum()))
        c3.metric("Errors", int((f.severity == "error").sum()) if len(f) else 0)
        c4.metric("Warnings", int((f.severity == "warning").sum()) if len(f) else 0)
        st.subheader("Submissions")
        st.dataframe(subs.drop(columns=["sha256", "processed_utc"], errors="ignore"),
                     hide_index=True, use_container_width=True)
        st.subheader("Findings")
        if len(f):
            sel = st.multiselect("Severity", ["error", "warning", "info"], default=["error", "warning"])
            ctry = st.multiselect("Country", sorted(f.country.unique()))
            view = f[f.severity.isin(sel) & (f.country.isin(ctry) if ctry else True)]
            st.dataframe(view, hide_index=True, use_container_width=True)
        else:
            st.success("No finding.")
        st.subheader("Gross enrolment ratio, both sexes")
        g = ind[(ind.indicator == "GER") & (ind.sex == "T")].dropna(subset=["value"])
        lvl = st.radio("Level", spec.levels, horizontal=True)
        chart = g[g.level == lvl].pivot_table(index="year", columns="country", values="value")
        chart.index = chart.index.astype(str)  # avoid "2,019" axis labels
        st.line_chart(chart)
        st.download_button("Download QC log (xlsx)", qc_bytes, "qc_log.xlsx")
        st.download_button("Download consolidated observations (csv)",
                           t["observations"].to_csv(index=False).encode("utf-8"), "observations.csv")
        if reports:
            pick = st.selectbox("Country review report", sorted(reports))
            st.download_button("Download this report (html)", reports[pick], f"{pick}.html")
            st.components.v1.html(reports[pick], height=600, scrolling=True)

with tab_tpl:
    lang = st.selectbox("Language", LANGUAGES, format_func=str.upper)
    with tempfile.TemporaryDirectory() as tmp:
        p = build_template(Path(tmp) / f"template_{spec.template_version}_{lang}.xlsx", spec, lang)
        st.download_button(f"Download blank template ({lang.upper()})", p.read_bytes(), p.name)

with tab_rules:
    st.dataframe(rules_table(), hide_index=True, use_container_width=True)
    st.caption("Thresholds are set in config/spec.yaml.")
