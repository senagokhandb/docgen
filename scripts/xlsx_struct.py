#!/usr/bin/env python3
"""DOCGEN xlsx structure: insert/delete rows or columns and sort ranges IN
PLACE. Every reference to the sheet is shifted the way Excel does it
(formulas on all sheets, names/print titles, conditional formats,
validations, tables, merged cells, hyperlinks, comments, drawing anchors,
chart series); styles, theme and all other parts stay untouched. Cached
formula results are then refreshed with LibreOffice (xlsx_recalc).

Usage:
    xlsx_struct.py <in.xlsx> <out.xlsx> insert-rows --sheet S --at 10 [--count 1]
    xlsx_struct.py <in.xlsx> <out.xlsx> delete-rows --sheet S --at 10 [--count 1] [--onay]
    xlsx_struct.py <in.xlsx> <out.xlsx> insert-cols --sheet S --at C [--count 1]
    xlsx_struct.py <in.xlsx> <out.xlsx> delete-cols --sheet S --at C [--count 1] [--onay]
    xlsx_struct.py <in.xlsx> <out.xlsx> sort --sheet S --range A1:H41 --by G [--desc] [--no-header] [--onay]

--onay: only when the user explicitly confirmed — deleting cells that other
formulas/names/rules read (they become #REF!), or sorting rows whose
formulas read other rows. Exit codes: 0 ok, 1 refused/failed, 2 usage.
"""
import argparse
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "lib"))
from docgen import DocgenError, consented  # noqa: E402
from docgen import xlsx as X  # noqa: E402
from docgen import xlsx_struct as XS  # noqa: E402
from docgen.package import Package  # noqa: E402


def errors(pkg):
    out = set()
    for name, part, _ in X.sheets(pkg):
        for c in pkg.xml(part).iter(X.s("c")):
            f = c.find(X.s("f"))
            if c.get("t") == "e" or (f is not None and f.text and "#REF!" in f.text):
                out.add((name, c.get("r")))
    return out


def lost_features(before, after):
    lost = []
    for k in ("sheets", "names"):
        missing = set(before[k]) - set(after[k])
        if missing:
            lost.append(f"{k}: {sorted(missing)}")
    if after["external_links"] < before["external_links"]:
        lost.append("external links")
    for sheet in before["sheets"]:
        for k, v in before.get(sheet, {}).items():
            if after.get(sheet, {}).get(k, 0) < v:
                lost.append(f"{sheet}: {k} {v} -> {after.get(sheet, {}).get(k, 0)}")
    return lost


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("src")
    ap.add_argument("dst")
    ap.add_argument("op", choices=["insert-rows", "delete-rows", "insert-cols", "delete-cols", "sort"])
    ap.add_argument("--sheet", required=True)
    ap.add_argument("--at")
    ap.add_argument("--count", type=int, default=1)
    ap.add_argument("--range")
    ap.add_argument("--by", help="sort key column letter")
    ap.add_argument("--desc", action="store_true")
    ap.add_argument("--no-header", action="store_true")
    ap.add_argument("--onay", action="store_true")
    ap.add_argument("--no-recalc", action="store_true", help="skip xlsx_recalc (not recommended)")
    a = ap.parse_args()

    pkg = Package(a.src)
    before_f, before_e = X.features(pkg), errors(pkg)
    try:
        if a.op == "insert-rows":
            XS.insert_rows(pkg, a.sheet, int(a.at), a.count)
        elif a.op == "delete-rows":
            XS.delete_rows(pkg, a.sheet, int(a.at), a.count, force=a.onay)
        elif a.op == "insert-cols":
            XS.insert_cols(pkg, a.sheet, a.at if not str(a.at).isdigit() else int(a.at), a.count)
        elif a.op == "delete-cols":
            XS.delete_cols(pkg, a.sheet, a.at if not str(a.at).isdigit() else int(a.at), a.count, force=a.onay)
        else:
            if not (a.range and a.by):
                print("usage error: sort needs --range and --by", file=sys.stderr)
                return 2
            XS.sort_range(pkg, a.sheet, a.range, a.by, descending=a.desc, header=not a.no_header, force=a.onay)
    except DocgenError as e:
        print(f"[xlsx_struct] REFUSED  {type(e).__name__}: {e}")
        return 1
    lost = lost_features(before_f, X.features(pkg))
    new_err = errors(pkg) - before_e
    if (lost or new_err) and not (a.onay and consented()):
        print("[xlsx_struct] REFUSED  the operation would damage the workbook"
              + (" (--onay is accepted only when the user's request contains it)" if a.onay else "") + ":")
        for l in lost:
            print(f"  lost {l}")
        for sh, ref in sorted(new_err):
            print(f"  #REF! in {sh}!{ref}")
        return 1
    pkg.save(a.dst)
    print(f"[xlsx_struct] OK  {a.op} on {a.sheet} -> {a.dst}; references shifted, all features kept")
    if not a.no_recalc:
        recalc = Path(__file__).resolve().parent / "xlsx_recalc.py"
        r = subprocess.run([sys.executable, str(recalc), a.dst], capture_output=True, text=True)
        print("  " + (r.stdout.strip().splitlines() or ["(xlsx_recalc produced no output)"])[0])
        if r.returncode != 0:
            print(r.stdout)
            return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
