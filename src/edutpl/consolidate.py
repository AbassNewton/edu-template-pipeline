"""Process a folder of submissions into one validated, consolidated database."""
from __future__ import annotations

import logging
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

from .config import Spec
from .indicators import compute_indicators
from .reader import read_workbook
from .report import country_report, qc_log
from .validate import checklist, validate

log = logging.getLogger("edutpl")


def process_folder(in_dir: Path | str, out_dir: Path | str, spec: Spec, reports: bool = True) -> dict:
    in_dir, out_dir = Path(in_dir), Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    obs_all, ind_all, find_all, chk_all, subs, par_all = [], [], [], [], [], []
    files = sorted(p for p in in_dir.glob("*.xlsx") if not p.name.startswith("~$"))
    for p in files:
        try:
            sub = read_workbook(p)
        except Exception as exc:  # a corrupt file must not stop the batch
            log.error("cannot read %s: %s", p.name, exc)
            subs.append({"file": p.name, "country": None, "status": f"unreadable: {exc}"})
            continue
        f = validate(sub, spec)
        ind = compute_indicators(sub, spec)
        chk = checklist(f, sub.country)
        n_err = int((f.severity == "error").sum()) if len(f) else 0
        subs.append({"file": p.name, "country": sub.country, "language": sub.language,
                     "template_version": sub.template_version, "sha256": sub.sha256,
                     "errors": n_err, "warnings": int((f.severity == "warning").sum()) if len(f) else 0,
                     "status": "accepted" if n_err == 0 else "returned to country",
                     "processed_utc": datetime.now(timezone.utc).isoformat(timespec="seconds")})
        o = sub.observations.assign(country=sub.country, file=p.name)
        obs_all.append(o)
        par_all += [{"country": sub.country, "code": k, "value": None if v is None else str(v)}
                    for k, v in sub.params.items()]
        ind_all.append(ind)
        find_all.append(f)
        chk_all.append(chk)
        if reports:
            country_report(sub, f, chk, ind, out_dir / "reports" / f"{sub.country}_review.html")
        log.info("%s: %d errors, %d findings", p.name, n_err, len(f))

    tables = {
        "submissions": pd.DataFrame(subs),
        "observations": pd.concat(obs_all, ignore_index=True) if obs_all else pd.DataFrame(),
        "indicators": pd.concat(ind_all, ignore_index=True) if ind_all else pd.DataFrame(),
        "findings": pd.concat([x for x in find_all if len(x)], ignore_index=True)
        if any(len(x) for x in find_all) else pd.DataFrame(),
        "checklists": pd.concat(chk_all, ignore_index=True) if chk_all else pd.DataFrame(),
        "parameters": pd.DataFrame(par_all, columns=["country", "code", "value"]),
    }
    db = out_dir / "education.sqlite"
    db.unlink(missing_ok=True)
    con = sqlite3.connect(db)  # closed explicitly: "with" only commits, and Windows locks open files
    try:
        for name, df in tables.items():
            df.astype({"age": "float"} if "age" in df else {}).to_sql(name, con, index=False)
        con.commit()
    finally:
        con.close()
    for name, df in tables.items():
        df.to_csv(out_dir / f"{name}.csv", index=False, encoding="utf-8")
    qc_log(tables["checklists"], tables["findings"], tables["submissions"], out_dir / "qc_log.xlsx")
    return tables
