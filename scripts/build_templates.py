#!/usr/bin/env python3
"""DOCGEN template builder — run at Docker image build time.

Generates base templates with {{PLACEHOLDER}}s from the theme setup:
  /opt/docgen/themes/templates/docx/base.docx   (semantic skeleton)
  /opt/docgen/themes/templates/pptx/basic.pptx  (10 placeholder slides)

The agent fills the placeholders (or uses the templates as structural
models). Full themed decks are built at request time with pptxgenjs
(see skills/pptx/SKILL.md).

Idempotent: safe to re-run; overwrites existing templates.
"""
import sys
from pathlib import Path

TEMPLATES_DIR = Path(__file__).resolve().parent.parent / "themes" / "templates"

DOCX_MD = """# {{TITLE}}

{{SUBTITLE}}

## {{SECTION_1_TITLE}}

{{SECTION_1_BODY}}

## {{SECTION_2_TITLE}}

{{SECTION_2_BODY}}

## {{SECTION_3_TITLE}}

{{SECTION_3_BODY}}

| {{TABLE_1_HEADER_1}} | {{TABLE_1_HEADER_2}} |
| --- | --- |
| {{TABLE_1_ROW_1_C1}} | {{TABLE_1_ROW_1_C2}} |
| {{TABLE_1_ROW_2_C1}} | {{TABLE_1_ROW_2_C2}} |
"""

N_SLIDES = 10


def build_docx():
    from office_oxide import create_from_markdown
    out = TEMPLATES_DIR / "docx" / "base.docx"
    out.parent.mkdir(parents=True, exist_ok=True)
    create_from_markdown(DOCX_MD, "docx", str(out))
    print(f"template: {out}")


def build_pptx():
    from office_oxide import PptxWriter
    out = TEMPLATES_DIR / "pptx" / "basic.pptx"
    out.parent.mkdir(parents=True, exist_ok=True)

    pw = PptxWriter()
    pw.set_presentation_size(12192000, 6858000)  # 16:9
    for i in range(1, N_SLIDES + 1):
        s = pw.add_slide()
        pw.set_slide_title(s, f"{{{{SLIDE{i}_TITLE}}}}")
        pw.add_slide_text(s, f"{{{{SLIDE{i}_BODY}}}}")
    pw.save(str(out))
    print(f"template: {out} ({N_SLIDES} slides)")


def main():
    try:
        build_docx()
        build_pptx()
    except Exception as e:
        print(f"template build FAILED: {e}", file=sys.stderr)
        return 1
    print("templates built OK")
    return 0


if __name__ == "__main__":
    sys.exit(main())
