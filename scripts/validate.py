#!/usr/bin/env python3
"""DOCGEN validation: structural checks + optional ECMA-376 XSD validation.

Usage:
    validate.py <file> [--json]

Supported: .docx .xlsx .pptx (OOXML) and .pdf
Exit codes: 0 = pass, 1 = fail, 2 = usage error
"""
import argparse
import json
import os
import re
import sys
import zipfile
from pathlib import Path

try:
    import lxml.etree as etree
except ImportError:  # pragma: no cover
    etree = None

SCHEMAS_DIR = Path(os.environ.get("DOCGEN_SCHEMAS", "/opt/docgen/schemas"))

# part path -> XSD file (relative to SCHEMAS_DIR; the original directory
# layout is kept so inter-schema xs:import locations resolve natively)
PART_SCHEMA = [
    (re.compile(r"^word/(document|styles|numbering|settings|fontTable|header\d*|footer\d*)\.xml$"),
     "wml/wml__ROOT.xsd"),
    (re.compile(r"^xl/worksheets/sheet\d*\.xml$"), "sml/sml_root.xsd"),
    (re.compile(r"^ppt/(presentation|slides/slide\d*|slideLayouts/slideLayout\d*|slideMasters/slideMaster\d*)\.xml$"),
     "pml/pml-2ed.xsd"),
]

XSD_REQUIRED_PARTS = {
    "docx": ["word/document.xml"],
    "xlsx": None,  # any xl/worksheets/sheet*.xml found is validated
    "pptx": ["ppt/presentation.xml"],
}


def _schema_for(part: str):
    for pattern, xsd in PART_SCHEMA:
        if pattern.match(part):
            return xsd
    return None


def structural_check(path: Path):
    """Zip integrity, [Content_Types].xml, .rels, XML well-formedness."""
    errors, warnings = [], []
    if not path.exists():
        return [f"file not found: {path}"], warnings
    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile as e:
        return [f"not a valid zip: {e}"], warnings

    with zf:
        names = zf.namelist()
        if "[Content_Types].xml" not in names:
            errors.append("missing [Content_Types].xml")

        bad_member = zf.testzip()
        if bad_member:
            errors.append(f"corrupt zip member: {bad_member}")

        if etree is None:
            warnings.append("lxml not available — XML checks skipped")
            return errors, warnings

        for name in names:
            if name.endswith((".xml", ".rels")):
                try:
                    etree.fromstring(zf.read(name))
                except Exception as e:
                    errors.append(f"XML not well-formed: {name}: {e}")
    return errors, warnings


def xsd_check(path: Path, kind: str):
    """Validate mapped parts against ECMA-376 XSDs (when present)."""
    errors, warnings = [], []
    if etree is None:
        warnings.append("lxml not available — XSD validation skipped")
        return errors, warnings

    if not SCHEMAS_DIR.is_dir():
        warnings.append(f"schemas dir missing ({SCHEMAS_DIR}) — XSD validation skipped")
        return errors, warnings

    schemas = {}

    def _load(name: str):
        if name not in schemas:
            p = SCHEMAS_DIR / name
            if not p.exists():
                schemas[name] = None
                return None
            try:
                schemas[name] = etree.XMLSchema(etree.parse(str(p)))
            except Exception as e:
                warnings.append(f"schema load failed ({name}): {e}")
                schemas[name] = False
        return schemas[name] if schemas[name] is not False else None

    try:
        zf = zipfile.ZipFile(path)
    except zipfile.BadZipFile as e:
        errors.append(f"zip unreadable for XSD check: {e}")
        return errors, warnings
    with zf:
        names = zf.namelist()
        checked = 0
        for name in names:
            xsd_name = _schema_for(name)
            if not xsd_name:
                continue
            schema = _load(xsd_name)
            if schema is None:
                continue
            try:
                elem = etree.fromstring(zf.read(name))
            except Exception as e:
                errors.append(f"cannot parse {name} for XSD check: {e}")
                continue
            checked += 1
            if not schema.validate(elem):
                for err in schema.error_log[:20]:
                    errors.append(f"XSD {xsd_name} :: {name} :: line {err.line}: {err.message}")

        if checked == 0:
            warnings.append("no parts matched the XSD part map (or no schemas present)")
        else:
            warnings.append(f"XSD-validated {checked} part(s)")
    return errors, warnings


def pdf_check(path: Path):
    errors, warnings = [], []
    head = path.open("rb").read(1024)
    if not head.startswith(b"%PDF"):
        errors.append("does not start with %PDF header")
        return errors, warnings
    try:
        import pdfplumber
        with pdfplumber.open(str(path)) as pdf:
            pages = len(pdf.pages)
        if pages == 0:
            errors.append("PDF has 0 pages")
        warnings.append(f"pdf pages: {pages}")
    except Exception as e:
        errors.append(f"pdfplumber could not open PDF: {e}")
    return errors, warnings


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("file")
    ap.add_argument("--json", action="store_true", help="machine-readable output")
    args = ap.parse_args()

    path = Path(args.file)
    kind = path.suffix.lstrip(".").lower()
    if kind not in ("docx", "xlsx", "pptx", "pdf"):
        print(f"usage error: unsupported extension .{kind}", file=sys.stderr)
        return 2

    if kind == "pdf":
        errors, warnings = pdf_check(path)
    else:
        errors, warnings = structural_check(path)
        if not errors:
            e, w = xsd_check(path, kind)
            errors += e
            warnings += w
        else:
            warnings.append("structural errors — XSD validation skipped")

    ok = not errors
    if args.json:
        print(json.dumps({
            "file": str(path), "kind": kind, "ok": ok,
            "errors": errors, "warnings": warnings,
        }, ensure_ascii=False, indent=2))
    else:
        status = "PASS" if ok else "FAIL"
        print(f"[validate] {status}  {path}  ({kind})")
        for w in warnings:
            print(f"  warn: {w}")
        for e in errors:
            print(f"  ERROR: {e}")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
