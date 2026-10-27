#!/usr/bin/env python3
"""DOCGEN xlsx recalc: compute every formula with LibreOffice and write the
results back into the ORIGINAL package as cached values (<v>). Nothing else
in the workbook is rewritten (tables, charts, styles, links stay as they are).

Run it after changing values or formulas (docgen.xlsx.set_value/set_formula)
so the file never carries stale results.

Usage:
    xlsx_recalc.py <file.xlsx> [-o out.xlsx] [--expect "Sheet!B4=20000"]... [--json]

Exit codes: 0 = ok, 1 = new formula errors / failed expectations, 2 = usage
"""
import argparse
import json
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "lib"))
from lxml import etree  # noqa: E402

from docgen import DocgenError  # noqa: E402
from docgen import office  # noqa: E402
from docgen import xlsx as X  # noqa: E402
from docgen.package import Package  # noqa: E402

S = X.s


def computed_values(path):
    """{(sheet, ref): (t, v)} from a LibreOffice-saved workbook."""
    pkg = Package(path)
    sst = X.shared_strings(pkg)
    out = {}
    for name, part, _ in X.sheets(pkg):
        for c in pkg.xml(part).iter(S("c")):
            t, v = c.get("t"), c.findtext(S("v"))
            if t == "s" and v is not None:
                t, v = "str", sst[int(v)]
            elif t == "inlineStr":
                t, v = "str", "".join(x.text or "" for x in c.iter(S("t")))
            out[(name, c.get("r"))] = (t, v)
    return out


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return v


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("file")
    ap.add_argument("-o", "--out", help="write here instead of updating the file in place")
    ap.add_argument("--expect", action="append", default=[], help='"Sheet!B4=20000" (repeatable)')
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args()
    src = Path(a.file)
    if not src.exists() or src.suffix.lower() not in (".xlsx", ".xlsm"):
        print("usage error: pass an existing .xlsx", file=sys.stderr)
        return 2
    try:
        lo_out = office.convert(src, "xlsx", tempfile.mkdtemp(prefix="docgen-recalc-"))
    except DocgenError as e:
        print(f"[xlsx_recalc] FAIL  {e}")
        return 1
    calc = computed_values(lo_out)
    pkg = Package(src)
    updated, errors, new_errors, changed = 0, [], [], []
    for name, part, _ in X.sheets(pkg):
        for c in pkg.xml(part).iter(S("c")):
            if c.find(S("f")) is None:
                continue
            ref = c.get("r")
            old_t, old_v = c.get("t"), c.findtext(S("v"))
            t, v = calc.get((name, ref), ("str", ""))
            for el in c.findall(S("v")):
                c.remove(el)
            c.attrib.pop("t", None)
            if t in ("str", "e", "b"):
                c.set("t", t)
            etree.SubElement(c, S("v")).text = "" if v is None else v
            updated += 1
            if t == "e":
                errors.append(f"{name}!{ref}={v}")
                if old_t != "e":
                    new_errors.append(f"{name}!{ref}={v}")
            if _num(old_v) != _num(v):
                changed.append(f"{name}!{ref}: {old_v} -> {v}")
    dst = a.out or str(src)
    pkg.save(dst)
    failed = []
    check = Package(dst)
    for spec in a.expect:
        target, _, want = spec.partition("=")
        got = X.get(check, target.strip())[0]
        if _num(str(got)) != _num(want.strip()) and str(got) != want.strip():
            failed.append(f"{target.strip()} is {got!r}, expected {want.strip()!r}")
    rep = {"file": dst, "formulas": updated, "changed_results": changed, "errors": errors,
           "new_errors": new_errors, "failed_expectations": failed,
           "ok": not new_errors and not failed}
    if a.json:
        print(json.dumps(rep, ensure_ascii=False, indent=2))
    else:
        print(f"[xlsx_recalc] {'PASS' if rep['ok'] else 'FAIL'}  {dst}  ({updated} formulas)")
        for c in changed[:30]:
            print(f"  result changed  {c}")
        if len(changed) > 30:
            print(f"  ... {len(changed) - 30} more")
        for e in new_errors:
            print(f"  NEW ERROR  {e}")
        for f in failed:
            print(f"  EXPECTATION FAILED  {f}")
    return 0 if rep["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
