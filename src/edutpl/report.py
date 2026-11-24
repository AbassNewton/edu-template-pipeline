"""Country indicator review report (self-contained HTML) and the QC log workbook."""
from __future__ import annotations

import base64
import html
import io
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402

from .reader import Submission  # noqa: E402

CSS = """
body{font-family:Segoe UI,Arial,sans-serif;margin:32px;color:#1a1a1a;max-width:1100px}
h1{color:#1F4E79;margin-bottom:4px} h2{color:#1F4E79;border-bottom:2px solid #1F4E79;padding-bottom:4px;margin-top:32px}
.meta{color:#595959;font-size:13px} table{border-collapse:collapse;font-size:12.5px;margin:8px 0;width:100%}
th{background:#1F4E79;color:#fff;padding:5px 8px;text-align:left} td{border-bottom:1px solid #ddd;padding:4px 8px}
td.num{text-align:right;font-variant-numeric:tabular-nums}
.pill{display:inline-block;padding:2px 10px;border-radius:10px;font-weight:600;margin-right:8px}
.error{background:#F8D7DA;color:#842029}.warning{background:#FFF3CD;color:#664D03}
.info{background:#CFE2FF;color:#084298}.pass{background:#D1E7DD;color:#0F5132}
.fail{background:#F8D7DA;color:#842029}.review{background:#FFF3CD;color:#664D03}
"""


def _png(fig) -> str:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=110, bbox_inches="tight")
    plt.close(fig)
    return base64.b64encode(buf.getvalue()).decode()


def _ger_chart(ind: pd.DataFrame) -> str | None:
    g = ind[(ind.indicator == "GER") & (ind.sex == "T")].dropna(subset=["value"])
    if g.empty:
        return None
    fig, ax = plt.subplots(figsize=(7, 3.2))
    for lvl, d in g.groupby("level"):
        ax.plot(d.year, d.value, marker="o", lw=2, label=lvl)
    ax.set_ylabel("GER (%)")
    ax.set_title("Gross enrolment ratio, both sexes")
    ax.grid(alpha=0.3)
    ax.legend(frameon=False)
    ax.spines[["top", "right"]].set_visible(False)
    return _png(fig)


def _table(df: pd.DataFrame, num_cols=(), classes: dict | None = None) -> str:
    head = "".join(f"<th>{html.escape(str(c))}</th>" for c in df.columns)
    body = []
    for _, r in df.iterrows():
        cells = []
        for c in df.columns:
            v = r[c]
            txt = "" if pd.isna(v) else (f"{v:,.2f}" if isinstance(v, float) else str(v))
            if classes and c in classes and txt:
                txt = f'<span class="pill {html.escape(txt)}">{html.escape(txt)}</span>'
            else:
                txt = html.escape(txt)
            cells.append(f'<td class="num">{txt}</td>' if c in num_cols else f"<td>{txt}</td>")
        body.append("<tr>" + "".join(cells) + "</tr>")
    return f"<table><thead><tr>{head}</tr></thead><tbody>{''.join(body)}</tbody></table>"


def country_report(sub: Submission, findings: pd.DataFrame, checklist: pd.DataFrame,
                   ind: pd.DataFrame, path: Path | str) -> Path:
    sev = findings.severity.value_counts() if len(findings) else pd.Series(dtype=int)
    pills = "".join(f'<span class="pill {s}">{s}: {int(sev.get(s, 0))}</span>'
                    for s in ("error", "warning", "info"))
    piv = ind.assign(series=lambda d: d.indicator + " " + d.level.fillna("") + " " + d.sex.fillna(""))
    piv = piv.pivot_table(index="series", columns="year", values="value", sort=False).reset_index()
    piv.columns = [str(c) for c in piv.columns]
    chart = _ger_chart(ind)
    f = findings.copy()
    if len(f):
        f["where"] = f.sheet.fillna("") + "!" + f.cell.fillna("")
        f = f[["rule", "severity", "where", "year", "message"]]
    html_doc = f"""<!doctype html><html><head><meta charset="utf-8">
<title>Indicator review - {html.escape(sub.country)}</title><style>{CSS}</style></head><body>
<h1>Indicator review: {html.escape(str(sub.params.get('PAR|country_name||') or sub.country))}</h1>
<div class="meta">File {html.escape(sub.path.name)} &middot; language {html.escape(str(sub.language))}
&middot; template {html.escape(str(sub.template_version))} &middot; sha256 {sub.sha256[:12]}
&middot; synthetic demo data</div>
<h2>Summary</h2><p>{pills}</p>
<h2>Quality-control checklist</h2>{_table(checklist.drop(columns=['country']), ('findings',), {'status': 1})}
<h2>Findings to resolve with the country</h2>
{_table(f, ('year',), {'severity': 1}) if len(f) else '<p>No finding.</p>'}
<h2>Indicators (recomputed)</h2>
{f'<img alt="GER chart" src="data:image/png;base64,{chart}">' if chart else ''}
{_table(piv, tuple(c for c in piv.columns if c != 'series'))}
</body></html>"""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(html_doc, encoding="utf-8")
    return path


def qc_log(checklists: pd.DataFrame, findings: pd.DataFrame, submissions: pd.DataFrame,
           path: Path | str) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    summary = checklists.pivot_table(index="country", columns="rule", values="status", aggfunc="first")
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        submissions.to_excel(xw, sheet_name="Submissions", index=False)
        summary.to_excel(xw, sheet_name="Checklist summary")
        checklists.to_excel(xw, sheet_name="Checklists", index=False)
        findings.to_excel(xw, sheet_name="Findings", index=False)
    return path
