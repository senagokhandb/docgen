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
| `/forms/pdfengines/crop` | POST multipart | page ranges (`pages=1-2,4`) |
| `/forms/pdfengines/watermark` | POST multipart | watermark PNG (`watermark.png=@f.png`, `%rotation`, `%scale`, `%opacity`, `%position`) |
| `/forms/pdfengines/compress` | POST multipart | compress (`mode=lossless`/`lossy`, `compress.images`, `quality`, `dpi`) |

Multipart conventions:
- PDFs: field name `files`.
- HTML: field name `files` with the `.html` file — the file **must be named
  `index.html`** for the chromium engine (Gotenberg 8): use
  `-F "files=@page.html;filename=index.html"`. Any same-dir referenced
  files (css/img) must be uploaded as additional `files=` entries.
- On error: response body is a JSON error message — always capture it:
  `curl -s ... -o out -w '%{http_code}'` and inspect `out` on 4xx/5xx.

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
