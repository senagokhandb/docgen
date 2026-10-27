#!/usr/bin/env python3
"""DOCGEN pptx repair — fixes known pptxgenjs XML quirks before validation.

Usage:
    python /opt/docgen/scripts/repair_pptx.py <deck.pptx> [--out other.pptx]

Fixes (idempotent, in-place by default):
  1. presentation.xml: reorders top-level children into the ECMA-376
     canonical order (pptxgenjs writes notesMasterIdLst after sldIdLst,
     which the XSD rejects).
  2. slide XML: splits paragraphs that contain a second <a:pPr> mid-
     paragraph (produced when pptxgenjs merges multiple text items into
     one <a:p>). Each pPr boundary becomes its own paragraph; the
     trailing endParaRPr stays with the last paragraph.

Exit: 0 (nothing to fix or fixed OK), 1 (file problem), 2 (usage).
"""
import argparse
import shutil
import sys
import zipfile
from pathlib import Path

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"

# canonical order for p:presentation children (pml-2ed.xsd sequence)
PRESENTATION_ORDER = [
    "sldMasterIdLst", "notesMasterIdLst", "handoutMasterIdLst",
    "sldIdLst", "sldSz", "notesSz", "defaultTextStyle", "textStyles",
    "clrMapOvr", "transition", "timing", "hwPr", "custDataLst",
]


def localname(el):
    return el.tag.rsplit("}", 1)[-1] if isinstance(el.tag, str) else ""


def reorder_presentation(root):
    order = {name: i for i, name in enumerate(PRESENTATION_ORDER)}
    children = list(root)
    indexed = sorted(
        enumerate(children),
        key=lambda ic: (order.get(localname(ic[1]), len(PRESENTATION_ORDER) + 1), ic[0]),
    )
    if [i for i, _ in indexed] == list(range(len(children))):
        return 0  # already in canonical order
    for c in children:
        root.remove(c)
    for i, c in indexed:
        root.append(c)
    return 1


def split_paragraphs(root):
    fixed = 0
    for p in root.iter(f"{{{A}}}p"):
        pprs = [c for c in p if localname(c) == "pPr"]
        if len(pprs) < 2:
            continue
        # groups: each pPr after the first starts a new paragraph
        groups = []
        current = []
        for c in list(p):
            if localname(c) == "pPr" and current:
                groups.append(current)
                current = []
            current.append(c)
        if current:
            groups.append(current)

        parent = p.getparent()
        insert_at = list(parent).index(p)
        new_paras = []
        for gi, group in enumerate(groups):
            from lxml import etree
            np = etree.Element(f"{{{A}}}p")
            for c in group:
                np.append(c)
            new_paras.append(np)
        for np in new_paras:
            insert_at += 1
            parent.insert(insert_at, np)
        parent.remove(p)
        fixed += 1
    return fixed


def repair(pptx_path: Path, out_path: Path) -> int:
    from lxml import etree

    with zipfile.ZipFile(pptx_path) as zin:
        names = zin.namelist()
        contents = {n: zin.read(n) for n in names}

    changed = 0

    pres_name = "ppt/presentation.xml"
    if pres_name in contents:
        tree = etree.fromstring(contents[pres_name])
        if reorder_presentation(tree):
            changed += 1
        contents[pres_name] = etree.tostring(
            tree, xml_declaration=True, encoding="UTF-8", standalone=True
        )

    for name in names:
        if name.startswith("ppt/slides/slide") and name.endswith(".xml"):
            tree = etree.fromstring(contents[name])
            n = split_paragraphs(tree)
            if n:
                changed += n
            contents[name] = etree.tostring(
                tree, xml_declaration=True, encoding="UTF-8", standalone=True
            )

    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for n in names:
            zout.writestr(n, contents[n])
    return changed


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pptx")
    ap.add_argument("--out", default=None)
    args = ap.parse_args()

    path = Path(args.pptx).expanduser().resolve()
    if not path.is_file():
        print(f"repair_pptx: no such file: {path}", file=sys.stderr)
        return 2

    out = Path(args.out).expanduser().resolve() if args.out else path
    tmp = path.parent / (path.name + ".repair-tmp")
    try:
        changed = repair(path, tmp)
        if out != path:
            shutil.move(str(tmp), str(out))
        else:
            shutil.move(str(tmp), str(path))
    except Exception as e:
        tmp.unlink(missing_ok=True)
        print(f"repair_pptx: failed: {e}", file=sys.stderr)
        return 1

    print(f"repair_pptx: {path.name}: {'fixed ' + str(changed) + ' problem(s)' if changed else 'clean'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
