---
name: pdf
description: Create and process PDF files — generate from HTML, edit text or pages of an existing PDF in place, merge, split, watermark, extract text and tables. Use for any .pdf task.
---

# PDF skill

## Tooling

| Need | Tool |
|---|---|
| New PDF (primary path) | Self-contained HTML + theme CSS → Gotenberg Chromium |
| Read text / tables | `readback.py <file.pdf> [--tables]`; fonts/positions: `readback.py <file.pdf> --format` |
| Where does a text occur? | `find_target.py <file.pdf> "<text>" [--page N] [--band header\|footer]` |
| **Change text in an existing PDF** (sentence, table cell, header/footer text) | `python /opt/docgen/scripts/pdf_text.py IN OUT "<old>" "<new>" [--page N] [--near "row key"] [--band header]` |
| **Add new paragraphs** after a line of an existing page | `python /opt/docgen/scripts/pdf_insert.py IN OUT --after "<line text>" --text "<paragraph>" [--text ...] [--page N]` |
| Delete / keep / move / insert / rotate pages (+ renumber "n / N") | `python /opt/docgen/scripts/pdf_pages.py IN OUT delete 4 --renumber` |
| Merge PDFs | `pdf_pages.py IN OUT insert other.pdf [--at N]` or Gotenberg `/forms/pdfengines/merge` |
| Watermark / stamp / optimize | Gotenberg `/forms/pdfengines/watermark` · `stamp` · `optimize` (REFERENCE.md §watermark — Turkish text needs the PDF-source recipe) |
| Render office→PDF | Gotenberg `/forms/libreoffice/convert` |
| Verify an edit | `diff_check.py SRC OUT --allow-page N --expect-text "<new>"` |

## Creation workflow

1. **Design as print HTML.** Write a self-contained file
   `/workspace/scratch/page.html`:
   - Inline the theme preset: COPY `/opt/docgen/themes/css/<theme>.css` into
     a `<style>` block (or upload it next to the HTML).
   - A4 page setup, print-safe typography (see REFERENCE.md §print-CSS).
   - User-requested styling for specific content (color, font, size,
     position): put it on exactly that element — inline `style="…"` or a
     class used only there.
   - Real content only — no placeholder text.
2. **Convert:**
   ```bash
   curl -s --fail-with-body -F "files=@/workspace/scratch/page.html;filename=index.html" \
     "$GOTENBERG_URL/forms/chromium/convert/html" -o /workspace/outputs/<ad>.pdf
   ```
3. **QA:** `readback.py` (all content present, nothing cut off), page count.

## Editing workflow (uploaded PDF)

Edit the PDF itself; do not re-create it from extracted text (that changes
fonts, spacing and every page).

**Source available → change the source and re-render.** The source is the
HTML (or docx/pptx) the PDF was built from: uploaded by the user, or in
`/workspace/previous/<n>/` of the message that delivered this PDF (chat UI).
Use it for anything more than a text fix — color/font, new tables or
sections, longer text that must reflow:
1. Copy it to `/workspace/scratch/` and render it UNCHANGED with the
   creation command below (same endpoint, no extra form fields).
   `diff_check.py <pdf> <render>` with no scope must PASS; otherwise the PDF
   changed after it was built — do not use the source, use the in-place
   steps below.
2. Change only what was asked (requested styling inline on that element,
   e.g. `<h2 style="color:#C00000">`) and render to `/workspace/outputs/`.
3. Update what the reflow moved (contents page numbers), then
   `diff_check.py <pdf> <out> --allow-page N ... --allow-structure
   --expect-text "<new>"` — every page from the changed one on, plus the
   contents page — and `readback.py --format` to confirm the style.

No source → edit in place:

1. `find_target.py` → one match? Use `--page`, `--near` (another text on the
   same line, e.g. the product name of a table row) or `--band` to narrow.
   Still several → report the candidates, write no file.
2. Text: `pdf_text.py SRC OUT "<old>" "<new>" ...` — only that line or
   paragraph is rewritten, in the document's own font family. For a table
   cell give the cell text as `<old>` plus `--near "<row key>"`.
   New paragraphs: `pdf_insert.py SRC OUT --after "<last line before them>"
   --text "<p1>" --text "<p2>"` — drawn below that line in the page's body
   font/size/color and column; nothing else moves. One `--text` per paragraph.
   Pages: `pdf_pages.py SRC OUT delete 4 --renumber` (also `keep`, `move 5
   --before 4`, `insert other.pdf --at 3`, `rotate 2 --degrees 90`).
   Header text on every page: run `pdf_text.py ... --band header --page N`
   once per page (chain the outputs).
3. Verify: `diff_check.py SRC OUT --allow-page 4 --expect-text "<new>"`
   (`--allow-structure --allow-band` for page operations with renumbering).

**No source → refuse honestly (no file, explain why) when:**
- `pdf_text.py` says the text does not fit / the font family is missing /
  the line mixes formatting — offer to edit the source document instead;
  never re-create the PDF without the user's consent.
- The request is a color/bold/font/alignment change of existing PDF text
  — not supported in place; ask for the source file.
- Moving a section that does not start/end on a page boundary (moving
  whole pages would also move other content) — explain and ask.
- `pdf_insert.py` says the text does not fit on the page: there is no
  reflow. Do not shrink the font or split the text on your own. Ask the
  user where the rest should go. With the user's
  placement ("2. sayfadan sonra yeni sayfalara"), render the new pages as
  HTML styled like the document (same font family/size/margins, read them
  with `readback.py --format`) and insert them with `pdf_pages.py insert`.
- **Scanned / image-only PDFs:** no text layer → write no file. Do not OCR
  it and never paint over or re-draw text in the image (that alters a
  record); ask for the source file or an OCR'd PDF. Do not guess content.

## Visual QA (off by default)

`--visual-qa` is **disabled** unless the user explicitly includes
`--visual-qa` in the request. Do not run it proactively.
