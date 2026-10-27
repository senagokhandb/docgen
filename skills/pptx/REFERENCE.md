# PPTX — technical reference

## office_oxide Python API (verified signatures)

```python
from office_oxide import PptxWriter, Document

# quick text-only deck
pw = PptxWriter()
pw.set_presentation_size(12192000, 6858000)   # EMU — 16:9 (13.33 x 7.5 in)
s = pw.add_slide()                             # -> slide index
pw.set_slide_title(s, "Başlık")
pw.add_slide_text(s, "Gövde metni")
pw.add_slide_image(s, png_bytes, "png", x, y, cx, cy)   # EMU coords/sizes
pw.save("deck.pptx")

# read
doc = Document.open("deck.pptx")
doc.to_markdown()        # slide-by-slide text
doc.plain_text()
```

Presentation sizes (EMU; 914400 EMU = 1 inch):
- 4:3  → `9144000 x 6858000`
- 16:9 → `12192000 x 6858000`

## pptxgenjs — key API

```js
const PptxGenJS = require("pptxgenjs");
const deck = new PptxGenJS();
deck.layout = "LAYOUT_WIDE";          // 13.33 x 7.5 in; "LAYOUT_16x9" = 10 x 5.63
deck.author = "docgen"; deck.title = "...";

// slide master = consistent template
const m = deck.defineSlideMaster({
  title: "CONTENT",
  background: { color: "FFFFFF" },
  objects: [
    { placeholder: { type: "title", x: .6, y: .35, w: 12.1, h: .9 } },
    { placeholder: { type: "body",  x: .6, y: 1.6, w: 12.1, h: 4.9 } },
    { text: { text: "", options: { placeholder: "slnum" } } },
  ],
});

const s = deck.addSlide({ masterName: "CONTENT" });
s.addText("Başlık", { placeholder: "title", fontSize: 30, bold: true, color: "1F4E79" });
s.addText(["b1", "b2", "b3"],
  { placeholder: "body", fontSize: 18, bullet: true, color: "1A1A1A", lineSpacing: 28 });

// charts
s.addChart(deck.ChartType.bar, [{ name: "2025", labels: ["Q1","Q2","Q3","Q4"], values: [1,2,3,4] }],
  { x: .8, y: 2, w: 11.5, h: 4.5, barDir: "col",
    chartColors: ["1F4E79", "2E75B6", "BDD7EE"] });

// tables
s.addTable([["A","B"],["1","2"]], { x: .8, y: 2, w: 11, h: 3,
  fontFace: "Calibri", fontSize: 14, color: "1A1A1A" });

// images
s.addImage({ path: "/workspace/scratch/chart.png", x: 8, y: 2, w: 4.5, h: 3 });

deck.writeFile({ fileName: "/workspace/outputs/deck.pptx" });
```

Color codes: hex WITHOUT `#`. Fonts: plain family name.

## Theme token schema

`/opt/docgen/themes/<name>.json`:
```json
{
  "name": "professional-blue",
  "colors": { "primary", "secondary", "accent", "background", "text", "muted" },
  "fonts":  { "heading", "body" },
  "spacing": { ... },
  "rules":   { "max_bullets_per_slide": 6, "max_fonts": 2, "accent_lines_under_titles": false }
}
```

## OOXML anatomy (pptx = zip)

```
[Content_Types].xml
ppt/presentation.xml           slide size, slide list
ppt/slideMasters/slideMasterN.xml
ppt/slideLayouts/slideLayoutN.xml
ppt/slides/slideN.xml          one file per slide
ppt/media/
ppt/_rels/presentation.xml.rels
```

Slide text (edit path):
```xml
<p:sld>
  <p:cSld>
    <p:spTree>
      <p:sp> ... <a:t>Başlık</a:t> ... </p:sp>   <!-- a:t = text runs -->
```
Namespaces:
- `p` = `http://schemas.openxmlformats.org/presentationml/2006/main`
- `a` = `http://schemas.openxmlformats.org/drawingml/2006/main`
- `r` = `.../officeDocument/2006/relationships`

Rules for hand edits:
- Replace text inside existing `a:t` only; do not add new `a:r` runs
  without a matching `a:rPr` (copy the sibling's run properties).
- Keep `xml:space="preserve"` where whitespace matters.
- After any edit: rezip from inside the dir → `validate.py` → readback →
  Gotenberg render (page count must equal slide count).

`scripts/validate.py` maps `ppt/slides/slide*.xml` → `schemas/pml.xsd`.

## Render QA details

- Gotenberg pptx→PDF: each slide = one PDF page (16:9 decks may letterbox).
- `pdfplumber`: `len(doc.pages)` == slide count; `extract_text()` on each
  page to confirm titles.
- If a slide's text is missing from the PDF text layer, the render likely
  failed on that slide — inspect the Gotenberg response body.
