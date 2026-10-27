---
name: pdf
description: Create and process PDF files — generate from HTML, merge, split, extract text and tables. Use for any .pdf task.
---

# PDF skill

## Tooling

| Need | Tool |
|---|---|
| New PDF (primary path) | Self-contained HTML + theme CSS → Gotenberg Chromium |
| Merge PDFs | Gotenberg `/forms/pdfengines/merge` |
| Crop / page ranges | Gotenberg `/forms/pdfengines/crop` (`pages=1-3` etc.) |
| Watermark / compress | Gotenberg pdfengines (`watermark`, `compress`) |
| Extract text + tables | `python /opt/docgen/scripts/readback.py <file.pdf>` (pdfplumber) |
| Read PDF as markdown | `office_oxide.Document.open(p).to_markdown()` |
| Render office→PDF (other formats) | Gotenberg `/forms/libreoffice/convert` |

## Creation workflow

1. **Design as print HTML.** Write a self-contained file
   `/workspace/scratch/page.html`:
   - Inline the theme preset: `<link>` is not allowed across requests, so
     COPY the contents of `/opt/docgen/themes/css/<theme>.css` into a
     `<style>` block, or reference it as a same-folder file when converting
     (Gotenberg serves relative paths next to the HTML).
   - A4 page setup, print-safe typography (see REFERENCE.md §print-CSS).
   - User-requested styling for specific content (color, font, size,
     position): put it on exactly that element — inline `style="…"` or a
     class used only there — so it overrides the theme preset without
     changing anything else.
   - Real content only — no placeholder text.
2. **Convert:**
   ```bash
   # Gotenberg 8 chromium motoru HTML dosyasını "index.html" olarak ister.
   curl -s -F "files=@/workspace/scratch/page.html;filename=index.html" \
     "$GOTENBERG_URL/forms/chromium/convert/html" \
     -o /workspace/outputs/<ad>.pdf
   ```
3. **QA:**
   - `readback.py` on the result: verify all expected content is present,
     no text cut off at page breaks.
   - Check page count matches expectation (pdfplumber `len(pdf.pages)`).
   - If a table splits badly across pages, add `page-break-inside: avoid`.

## Processing workflow

```bash
# merge
curl -s -F "files=@a.pdf" -F "files=@b.pdf" \
  "$GOTENBERG_URL/forms/pdfengines/merge" -o /workspace/outputs/merged.pdf

# crop (page ranges, 1-based, comma-separated)
curl -s -F "files=@deck.pdf" -F "pages=1-3,7" \
  "$GOTENBERG_URL/forms/pdfengines/crop" -o /workspace/outputs/kisa.pdf

# watermark (PNG file, %rotation etc. via form fields — see REFERENCE.md)
curl -s -F "files=@a.pdf" -F "watermark.png=@logo.png" -F "watermark.png%rotation=45" \
  "$GOTENBERG_URL/forms/pdfengines/watermark" -o /workspace/outputs/w.pdf
```
After any pdfengine operation: `readback.py` sanity check (page count +
first/last page text).

## Rules

- **"Edit this PDF":** first ask/look for the source file (docx/pptx/html)
  in `/workspace/uploads/`. If a source exists, edit the source and
  re-render. If only the PDF exists, explain the realistic options
  (overlay, merge, re-create from extracted text) and pick the safest.
- **Scanned / image-only PDFs:** text extraction returns little or
  nothing — in that case tell the user OCR is needed (their Docling
  pipeline, outside this container); do not guess content.
- **Never invent content** to fill a page.

## Visual QA (off by default)

`--visual-qa` (screenshot-based layout review using a vision model via the
user's Docling setup) is **disabled** unless the user explicitly includes
`--visual-qa` in the request. Do not run it proactively.
