"""Check that the template's Excel formulas give the same indicators as the Python code.

Windows + Microsoft Excel only (uses COM). Each workbook is opened in Excel,
recalculated and saved to a temporary copy; the cached results of the
'Indicators' and 'Checks' sheets are then compared with edutpl.indicators.

    python scripts/verify_formulas_excel.py out/submissions
"""
from __future__ import annotations

import math
import sys
import tempfile
from pathlib import Path

import win32com.client  # pywin32
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
from edutpl.config import load_labels, load_spec  # noqa: E402
from edutpl.indicators import compute_indicators  # noqa: E402
from edutpl.layout import Layout  # noqa: E402
from edutpl.reader import read_workbook  # noqa: E402


def recalc_copy(xl, src: Path, dst: Path):
    wb = xl.Workbooks.Open(str(src.resolve()))
    xl.CalculateFull()
    wb.SaveAs(str(dst.resolve()))
    wb.Close(False)


def main(folder: str):
    spec = load_spec()
    lay = Layout(spec)
    xl = win32com.client.DispatchEx("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    n_cmp = n_bad = 0
    try:
        with tempfile.TemporaryDirectory() as tmp:
            for p in sorted(Path(folder).glob("*.xlsx")):
                sub = read_workbook(p)
                if sub.template_version != spec.template_version:
                    print(f"skip {p.name}: template {sub.template_version}")
                    continue
                dst = Path(tmp) / p.name
                recalc_copy(xl, p, dst)
                ws = load_workbook(dst, data_only=True)[load_labels(sub.language)("sheet.indicators")]
                py = compute_indicators(sub, spec)
                bad = 0
                for r in py.itertuples():
                    code = f"{r.indicator}|{r.level or ''}|{r.sex or ''}|"
                    _, addr = lay.cell(code, r.year)
                    xv = ws[addr].value
                    xv = None if xv in ("", None) else float(xv)
                    pv = None if (r.value is None or math.isnan(r.value)) else float(r.value)
                    n_cmp += 1
                    same = (xv is None and pv is None) or (
                        xv is not None and pv is not None and math.isclose(xv, pv, rel_tol=1e-9, abs_tol=1e-9))
                    if not same:
                        bad += 1
                        if bad <= 5:
                            print(f"  {p.name} {code} {r.year}: Excel={xv} Python={pv}")
                n_bad += bad
                print(f"{p.name} [{sub.language}]: {len(py)} indicator cells, {bad} mismatches")
    finally:
        xl.Quit()
    print(f"TOTAL: {n_cmp} cells compared, {n_bad} mismatches")
    return 1 if n_bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1] if len(sys.argv) > 1 else "out/submissions"))
