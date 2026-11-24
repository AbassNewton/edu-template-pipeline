"""Command line: python -m edutpl <command>."""
from __future__ import annotations

import argparse
import logging
from pathlib import Path

from .build import build_template, extend_template
from .config import LANGUAGES, ROOT, load_spec
from .consolidate import process_folder
from .synth import generate
from .validate import rules_table


def main(argv=None):
    ap = argparse.ArgumentParser(prog="edutpl", description="Multi-year education template pipeline")
    ap.add_argument("--spec", default=None, help="path to spec.yaml")
    sp = ap.add_subparsers(dest="cmd", required=True)

    b = sp.add_parser("build", help="blank templates in every language")
    b.add_argument("--out", default=str(ROOT / "out" / "templates"))
    b.add_argument("--langs", nargs="*", default=list(LANGUAGES))

    e = sp.add_parser("extend", help="re-issue a filled template under a new year range")
    e.add_argument("src")
    e.add_argument("dst")
    e.add_argument("--last-year", type=int, required=True)
    e.add_argument("--first-year", type=int)
    e.add_argument("--version")
    e.add_argument("--lang")

    s = sp.add_parser("synth", help="synthetic submissions with injected errors")
    s.add_argument("--out", default=str(ROOT / "out" / "submissions"))
    s.add_argument("-n", type=int, default=10)
    s.add_argument("--seed", type=int, default=2027)

    p = sp.add_parser("process", help="validate + consolidate a folder of submissions")
    p.add_argument("--in", dest="inp", default=str(ROOT / "out" / "submissions"))
    p.add_argument("--out", default=str(ROOT / "out" / "processed"))

    sp.add_parser("rules", help="print the validation rule registry")
    sp.add_parser("all", help="build + synth + process")

    a = ap.parse_args(argv)
    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(message)s")
    spec = load_spec(a.spec)

    if a.cmd in ("build", "all"):
        out = Path(getattr(a, "out", ROOT / "out" / "templates")) if a.cmd == "build" else ROOT / "out" / "templates"
        for lang in getattr(a, "langs", LANGUAGES):
            print(build_template(out / f"template_{spec.template_version}_{lang}.xlsx", spec, lang))
    if a.cmd in ("synth", "all"):
        out = Path(a.out) if a.cmd == "synth" else ROOT / "out" / "submissions"
        truth = generate(out, spec, getattr(a, "n", 10), getattr(a, "seed", 2027))
        print(f"{len(truth)} errors injected into submissions in {out}")
    if a.cmd in ("process", "all"):
        inp = Path(a.inp) if a.cmd == "process" else ROOT / "out" / "submissions"
        out = Path(a.out) if a.cmd == "process" else ROOT / "out" / "processed"
        t = process_folder(inp, out, spec)
        print(t["submissions"][["file", "country", "errors", "warnings", "status"]].to_string(index=False))
        print(f"outputs in {out}")
    if a.cmd == "extend":
        new = spec.with_years(a.first_year, a.last_year, a.version)
        print(extend_template(a.src, a.dst, new, a.lang))
    if a.cmd == "rules":
        print(rules_table().to_string(index=False))


if __name__ == "__main__":
    main()
