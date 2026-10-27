#!/usr/bin/env python3
"""DOCGEN recalc QA: xlsx -> LibreOffice (via Gotenberg) -> PDF -> scan for
formula error tokens and optionally verify expected values.

Usage:
    recalc_check.py <file.xlsx> [--expect "label:12345.67"]... [--json]

Exit codes: 0 = clean, 1 = errors found / check failed, 2 = usage error
"""
import argparse
import json
import os
import re
import sys
import tempfile

import requests

ERROR_TOKENS = ["#REF!", "#NAME?", "#DIV/0!", "#VALUE!", "#N/A", "#NULL!", "#NUM!"]


def convert_to_pdf(xlsx_path: str, gotenberg_url: str, timeout: int = 120):
    url = f"{gotenberg_url}/forms/libreoffice/convert"
    with open(xlsx_path, "rb") as f:
        resp = requests.post(
            url,
            files={"files": (os.path.basename(xlsx_path), f,
                              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            timeout=timeout,
        )
    if resp.status_code != 200:
        raise RuntimeError(
            f"Gotenberg convert failed (HTTP {resp.status_code}): {resp.text[:500]}"
        )
    return resp.content


def extract_text(pdf_bytes: bytes) -> str:
    import io
    import pdfplumber
    out = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            out.append(page.extract_text() or "")
    return "\n".join(out)


def norm(s: str) -> str:
    """Normalize text for value matching (drop spaces/underscores, unify comma)."""
    return re.sub(r"[\s_]", "", s).replace(",", ".")


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("file")
    ap.add_argument("--expect", action="append", default=[],
                    help='expected rendered value, e.g. "Toplam:20000" (repeatable)')
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()

    from pathlib import Path
    path = Path(args.file)
    if not path.exists() or path.suffix.lower() != ".xlsx":
        print("usage error: pass an existing .xlsx file", file=sys.stderr)
        return 2

    gotenberg_url = os.environ.get("GOTENBERG_URL", "http://gotenberg:3000")
    report = {"file": str(path), "gotenberg": gotenberg_url,
              "errors_found": 0, "total_errors": 0, "missing_expectations": []}

    try:
        pdf_bytes = convert_to_pdf(str(path), gotenberg_url)
    except Exception as e:
        report["ok"] = False
        report["fatal"] = str(e)
        _emit(report, args.json)
        return 1

    text = extract_text(pdf_bytes)
    report["rendered_chars"] = len(text)

    # 1) error tokens
    for token in ERROR_TOKENS:
        n = text.count(token)
        if n:
            report["total_errors"] += n
            report["errors_found"] += 1
            report.setdefault("error_tokens", {})[token] = n

    # 2) expected values
    for spec in args.expect:
        if ":" not in spec:
            report["missing_expectations"].append(f"bad --expect (no ':') : {spec}")
            continue
        label, value = spec.split(":", 1)
        if norm(value) in norm(text):
            report.setdefault("expectations_ok", []).append(spec)
        else:
            report["missing_expectations"].append(spec)

    report["ok"] = report["total_errors"] == 0 and not report["missing_expectations"]
    _emit(report, args.json)
    return 0 if report["ok"] else 1


def _emit(report: dict, as_json: bool):
    if as_json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        status = "PASS" if report.get("ok") else "FAIL"
        print(f"[recalc_check] {status}  {report['file']}")
        if report.get("fatal"):
            print(f"  FATAL: {report['fatal']}")
        if report.get("error_tokens"):
            for tok, n in report["error_tokens"].items():
                print(f"  ERROR TOKEN: {tok} x{n}")
        for m in report.get("missing_expectations", []):
            print(f"  MISSING EXPECTED VALUE: {m}")
        if report.get("expectations_ok"):
            for e in report["expectations_ok"]:
                print(f"  expected ok: {e}")
        print(f"  rendered chars: {report.get('rendered_chars', 0)}")


if __name__ == "__main__":
    sys.exit(main())
