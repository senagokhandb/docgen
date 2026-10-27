# PDF — technical reference

## Gotenberg v8 endpoints (base: $GOTENBERG_URL)

| Endpoint | Method | Purpose |
|---|---|---|
| `/health` | GET | liveness |
| `/forms/chromium/convert/html` | POST multipart | HTML → PDF |
| `/forms/chromium/convert/pdf` | POST multipart | PDF → PDF (re-render) |
| `/forms/chromium/convert/url` | POST multipart | URL → PDF (`url=` field) |
| `/forms/libreoffice/convert` | POST multipart | docx/pptx/xlsx → PDF |
| `/forms/pdfengines/merge` | POST multipart | merge N PDFs (order = `files=` order) |
| `/forms/pdfengines/split` | POST multipart | `splitMode=pages`, `splitSpan=1-3,5`, `splitUnify=true` (one PDF) |
| `/forms/pdfengines/watermark` | POST multipart | `watermarkSource=text\|image\|pdf`, `watermarkExpression`, optional `watermarkPages`; see §watermark below. A JSON object in `watermarkOptions` is rejected (HTTP 400) — leave it out. |
| `/forms/pdfengines/stamp` | POST multipart | same fields with the `stamp` prefix |
| `/forms/pdfengines/rotate` | POST multipart | `rotateAngle`, `rotatePages` |
| `/forms/pdfengines/optimize` | POST multipart | `imageQuality` |
| `/forms/pdfengines/flatten`, `/metadata/read`, `/metadata/write`, `/encrypt`, `/embed` | POST multipart | as named |

### watermark (tested on Gotenberg 8.37)

`watermarkSource=text` draws with a font that has no Turkish glyphs
("GİZLİ" comes out as "G ZL"). For any non-ASCII text render the watermark
with Chromium first and apply it as a PDF:
```bash
cat > /workspace/scratch/index.html <<'EOF'
<!doctype html><html><head><meta charset="utf-8"><style>@page{size:A4;margin:0} body{margin:0}
div{margin-top:300px;text-align:center;font:bold 110px sans-serif;color:#C00000;opacity:.3;transform:rotate(-35deg)}
</style></head><body><div>GİZLİ</div></body></html>
EOF
curl -s --fail-with-body -F "files=@/workspace/scratch/index.html" "$GOTENBERG_URL/forms/chromium/convert/html" -o /workspace/scratch/wm.pdf
curl -s --fail-with-body -F "files=@/workspace/uploads/rapor.pdf" -F "watermark=@/workspace/scratch/wm.pdf" \
     -F "watermarkSource=pdf" -F "watermarkExpression=wm.pdf" \
     "$GOTENBERG_URL/forms/pdfengines/watermark" -o /workspace/outputs/rapor-filigranli.pdf
```
The watermark file goes in the form field `watermark` (not `files`).

Verified against Gotenberg 8.37 (`pkg/modules/pdfengines/routes.go`): there
is **no `crop` and no `compress` route**, and no route edits existing text.
Page delete/keep/move/insert and text edits: `scripts/pdf_pages.py`,
`scripts/pdf_text.py`, `scripts/pdf_insert.py` (pdfium, below).

## docgen.pdf — page and text editing (pdfium)

```python
from docgen import pdf as F
F.find(path, "iki yeni depo", page=None, band=None, near=None)   # matches (page/line context)
F.replace_text(src, dst, old, new, page=4, near="Defter", occurrence=None,
               band=None|"header"|"footer", strategy="auto"|"inplace"|"redraw")   # -> report dict
F.insert_paragraphs(src, dst, after="<line text>", paragraphs=["p1", "p2"], page=None, near=None)
F.renumber(src, dst)                              # "n / N" page numbers in header/footer bands
F.delete_pages(src, dst, "4"); F.keep_pages(src, dst, "1-3")
F.move_pages(src, dst, "5", before=4); F.insert_pdf(src, dst, other, at=None, pages=None)
F.rotate_pages(src, dst, "2", 90); F.lines(path); F.page_texts(path)
```
`replace_text` strategies: *inplace* rewrites the existing text object with
its own embedded font (only when every new character already appears in
that font and the line still fits); *redraw* removes the paragraph's lines
and draws them again with the full font file of the same family (installed:
Caladea, Carlito, Liberation, DejaVu), wrapping inside the column. It
refuses mixed-format lines, whole table rows, missing font families and
text that does not fit before the next block — and verifies the page by
reading it back before writing.

`insert_paragraphs` draws new paragraphs right below the line containing
`after`, using that page's body font family, size, color, line pitch and
text column (a small gap between paragraphs). Nothing on the page moves;
if the paragraphs do not fit above the next block / footer it raises
`NotSupported` with how many lines are needed vs. available.

Multipart conventions:
- PDFs: field name `files`.
- HTML: field name `files` with the `.html` file — the file **must be named
  `index.html`** for the chromium engine (Gotenberg 8): use
  `-F "files=@page.html;filename=index.html"`. Any same-dir referenced
  files (css/img) must be uploaded as additional `files=` entries.
- On error: the response body is Gotenberg's error message. Always pass
  `--fail-with-body`: curl then exits non-zero on 4xx/5xx (without it the
  error text is saved as the output file and curl exits 0) and `out` still
  holds the message to read.

## print-CSS template (A4)

```css
@page { size: A4; margin: 18mm 16mm; }
* { box-sizing: border-box; }
html { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
body { font-family: Calibri, "Segoe UI", Arial, sans-serif;
       font-size: 10.5pt; line-height: 1.5; color: #1A1A1A; }
h1 { font-size: 20pt; margin: 0 0 12pt; }
h2 { font-size: 13pt; margin: 16pt 0 6pt; page-break-after: avoid; }
h3 { font-size: 11pt; margin: 12pt 0 4pt; page-break-after: avoid; }
p  { margin: 0 0 8pt; text-align: justify; }
table { width: 100%; border-collapse: collapse; margin: 8pt 0 12pt;
        page-break-inside: auto; }
tr  { page-break-inside: avoid; }
th, td { border: 1px solid #BFBFBF; padding: 4pt 6pt; font-size: 9.5pt;
         vertical-align: top; }
thead { display: table-header-group; }        /* repeat header each page */
img  { max-width: 100%; }
.page-break { page-break-before: always; }
.avoid-break { page-break-inside: avoid; }
```

Theme CSS presets live in `/opt/docgen/themes/css/*.css` (same structure,
theme colors/fonts). Copy the preset's contents into the HTML `<style>`
block — Gotenberg converts one request's files, so self-contained HTML is
the reliable path.

## pdfplumber — read-back

```python
import pdfplumber
with pdfplumber.open("out.pdf") as pdf:
    print(len(pdf.pages))
    for i, page in enumerate(pdf.pages, 1):
        t = page.extract_text() or ""
        tables = page.extract_tables()
        # check expected strings, table shapes...
```

## office_oxide PDF read

```python
from office_oxide import Document
doc = Document.open("out.pdf")
doc.to_markdown()       # text-layer content as markdown
```

## Error tokens to scan (recalc-style QA for PDFs)

`#REF!` `#NAME?` `#DIV/0!` `#VALUE!` `#N/A` `#NULL!` `#NUM!` — relevant
when the PDF was rendered from an xlsx via Gotenberg.

## Gotchas

- Chromium renders headless: no JS-heavy pages (server-render only).
  Keep HTML static (inline CSS, local images).
- Fonts: only fonts installed in the Gotenberg image are available;
  the CSS stack `Calibri, "Segoe UI", Arial, sans-serif` degrades safely.
- Very long documents: split the HTML into per-section files and merge
  the PDFs if memory/page limits hit.
- Merged PDF page count = sum of parts — verify after every merge.
