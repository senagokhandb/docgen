#!/usr/bin/env python3
"""DOCGEN visual QA plug-in (OFF in the normal flow).

Renders a final PDF to per-page PNGs so the output can be inspected
visually — by a person, a vision model, or an OCR cross-check.

Usage:
    python /opt/docgen/scripts/visual_qa.py <final.pdf> \
        [--pages-dir DIR] [--contact-sheet] [--docling] [--json]

Options:
    --pages-dir DIR   where to write page_001.png ... (default: <pdf>.pages/)
    --contact-sheet   also write contact_sheet.png (grid of all pages)
    --docling         if Docling is importable in this environment, run it on
                      each page image and report extracted text length;
                      pages with ~0 text are flagged (useful for scanned or
                      render-broken pages). Skipped silently when Docling is
                      not installed — the PNGs are still produced.

Exit: 0 = rendered fine (no blank pages detected), 1 = problem(s) found,
      2 = usage/render error.
"""
import argparse
import json
import sys
from pathlib import Path


def render_pages(pdf_path: Path, pages_dir: Path):
    import pypdfium2 as pdfium

    pages_dir.mkdir(parents=True, exist_ok=True)
    pdf = pdfium.PdfDocument(str(pdf_path))
    paths = []
    for i in range(len(pdf)):
        page = pdf[i]
        bitmap = page.render(scale=150 / 72)  # 150 dpi
        pil = bitmap.to_pil()
        out = pages_dir / f"page_{i + 1:03d}.png"
        pil.save(out)
        paths.append(out)
    pdf.close()
    return paths


def contact_sheet(page_paths, out_path: Path, cols: int = 4):
    from PIL import Image

    thumbs, w, h = [], 384, 243  # 16:10-ish thumb
    for p in page_paths:
        im = Image.open(p).convert("RGB")
        im.thumbnail((w, h))
        thumbs.append(im)
    rows = (len(thumbs) + cols - 1) // cols
    sheet = Image.new("RGB", (cols * w, rows * h), "white")
    for idx, im in enumerate(thumbs):
        r, c = divmod(idx, cols)
        sheet.paste(im, (c * w, r * h))
    sheet.save(out_path)
    return out_path


def docling_check(page_paths):
    """Optional OCR cross-check. Returns per-page stats; None if no Docling."""
    try:
        from docling.datamodel.base_models import InputFormat
        from docling.document_converter import DocumentConverter
    except Exception:
        return None

    converter = DocumentConverter(
        allowed_formats=[InputFormat.IMAGE]
    )
    stats = []
    for p in page_paths:
        res = converter.convert(str(p))
        text = res.document.export_to_markdown() or ""
        stats.append({
            "page": p.name,
            "chars": len(text.replace(" ", "")),
            "blank_suspect": len(text.replace(" ", "")) < 20,
        })
    return stats


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("pdf")
    ap.add_argument("--pages-dir", default=None)
    ap.add_argument("--contact-sheet", action="store_true")
    ap.add_argument("--docling", action="store_true")
    ap.add_argument("--json", action="store_true")
    args = ap.parse_args()

    pdf_path = Path(args.pdf).expanduser().resolve()
    if not pdf_path.is_file():
        print(f"visual_qa: no such file: {pdf_path}", file=sys.stderr)
        return 2

    try:
        pages_dir = Path(args.pages_dir).expanduser().resolve() if args.pages_dir \
            else pdf_path.parent / (pdf_path.stem + ".pages")
        page_paths = render_pages(pdf_path, pages_dir)
    except Exception as e:
        print(f"visual_qa: render failed: {e}", file=sys.stderr)
        return 2

    report = {
        "pdf": str(pdf_path),
        "pages": len(page_paths),
        "pages_dir": str(pages_dir),
        "page_files": [str(p) for p in page_paths],
        "docling": None,
        "blank_pages": [],
        "ok": True,
    }
    if args.contact_sheet:
        cs = contact_sheet(page_paths, pages_dir / "contact_sheet.png")
        report["contact_sheet"] = str(cs)

    if args.docling:
        stats = docling_check(page_paths)
        if stats is not None:
            report["docling"] = stats
            report["blank_pages"] = [s["page"] for s in stats if s["blank_suspect"]]

    if report["blank_pages"]:
        report["ok"] = False

    if args.json:
        print(json.dumps(report, indent=2))
    else:
        print(f"visual_qa: {report['pages']} page(s) -> {pages_dir}")
        if report.get("contact_sheet"):
            print(f"contact sheet: {report['contact_sheet']}")
        if report["docling"] is not None:
            for s in report["docling"]:
                flag = " BLANK?" if s["blank_suspect"] else ""
                print(f"  {s['page']}: {s['chars']} chars{flag}")
        elif args.docling:
            print("docling not importable in this environment — PNGs only")

    return 0 if report["ok"] else 1


if __name__ == "__main__":
    sys.exit(main())
