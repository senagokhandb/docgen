---
name: pptx
description: Create and edit PowerPoint (.pptx) presentations. Use for decks, pitch materials, slide-based output, or to modify an existing .pptx.
---

# PPTX skill

## Tooling

| Need | Tool |
|---|---|
| New deck, text + bullets, theme applied automatically | `node /opt/docgen/scripts/build_themed_deck.js <theme.json> <content.json> <out.pptx>` (QUICK — recommended for plain decks) |
| New deck, full control (charts, images, custom layouts) | Node `pptxgenjs` script in `/workspace/scratch/` (PRIMARY for anything custom) |
| New deck, quick/text-only | Python `office_oxide.PptxWriter` |
| Read any deck's text | Python `office_oxide.Document.open(p).to_markdown()` |
| Text-only edit on existing deck | unzip → edit `ppt/slides/slideN.xml` (`a:t` elements) with `lxml` → rezip |
| Style edit on existing deck (color, font, size, position) | unzip → edit `a:rPr` / `a:xfrm` in `ppt/slides/slideN.xml` with `lxml` → rezip |
| Render check (page count, content) | Gotenberg `/forms/libreoffice/convert` (pptx→PDF) + `pdfplumber` |
| Validation | `python /opt/docgen/scripts/validate.py <file>` |

## Theme workflow (required for new decks)

1. Pick the theme: `/opt/docgen/themes/{professional-blue,minimal,dark-corporate}.json`.
   User's choice wins; otherwise default to `professional-blue` and say so.
2. Build with `pptxgenjs`: apply the theme tokens to `defineSlideMaster`
   (colors, fonts, background) and to every text/chart element.
   Do not invent colors outside the theme tokens on your own. If the user
   explicitly asks for a color, font, size or position, apply exactly that
   at the requested scope (QWEN.md › Themes) — that overrides this rule.
3. Anti-AI design rules (hard requirements):
   - No accent lines or bars under titles.
   - Max **6 short bullets** per slide; one idea per slide.
   - Max **2 font families** (theme `fonts.heading` + `fonts.body`), unless
     the user explicitly names another font.
   - Consistent template: title zone top, content zone, footer with page
     number on every slide.
   - Whitespace over decoration; no clip-art, no random icons.
4. Verify: validate → readback (all slide titles/text present) → Gotenberg
   render (PDF page count == slide count).

## Quick themed deck (recommended for text-only decks)

1. Pick the theme JSON (user's choice, else `professional-blue`).
2. Write the content as JSON in `/workspace/scratch/deck_content.json`:
   ```json
   {
     "title": "Proje Sunumu",
     "subtitle": "2026 Q4",
     "slides": [
       { "title": "Özet", "bullets": ["kapsam", "taktik", "bütçe"] },
       { "title": "Sonuç", "bullets": ["öneri", "sonraki adım"] }
     ]
   }
   ```
3. Build:
   ```
   node /opt/docgen/scripts/build_themed_deck.js \
     /opt/docgen/themes/professional-blue.json \
     /workspace/scratch/deck_content.json \
     /workspace/outputs/sunum.pptx
   ```
   The builder applies all theme tokens and enforces the anti-AI rules
   (max 6 bullets/slide, no accent lines, 2 fonts, page numbers).
4. Repair + QA:
   ```
   python /opt/docgen/scripts/repair_pptx.py /workspace/outputs/sunum.pptx
   ```
   then validate → readback → Gotenberg render (page count == slide count).

Use a custom `pptxgenjs` script instead whenever the deck needs charts,
images, tables, or non-standard layouts, or the user asks for specific
colors, fonts, sizes or positions (the builder only applies theme tokens).
In pptxgenjs every `addText` takes `x`, `y`, `w`, `h` (inches), `color`
(hex, no `#`), `fontFace`, `fontSize` (pt), `bold`, `align`; to style part
of a line in a non-placeholder text box, pass a run array:
`s.addText([{ text: "Normal " }, { text: "kırmızı", options: { color: "FF0000" } }], { x: 1, y: 2, w: 8, h: 0.6 })`.

## Creation workflow (pptxgenjs)

1. Read the theme JSON, then write `/workspace/scratch/build_deck.js`:
   ```js
   const PptxGenJS = require("pptxgenjs");
   const theme = require("/opt/docgen/themes/professional-blue.json");
   const deck = new PptxGenJS();
   deck.layout = "LAYOUT_WIDE";            // 13.33 x 7.5 in
   deck.theme = { headFontFace: theme.fonts.heading, bodyFontFace: theme.fonts.body };
   const master = deck.defineSlideMaster({
     title: "MASTER",
     background: { color: theme.colors.background },
     objects: [
       { placeholder: { options: { name: "PH_TITLE", type: "title",
             x: 0.6, y: 0.35, w: 12.1, h: 1.0 } } },
       { placeholder: { options: { name: "PH_BODY", type: "body",
             x: 0.6, y: 1.7, w: 12.1, h: 4.8 } } },
     ],
     slideNumber: { x: 12.7, y: 7.15, w: 0.45, h: 0.25,
       color: theme.colors.muted, fontSize: 10 },
   });
   const s1 = deck.addSlide({ masterName: "MASTER" });
   s1.addText("Başlık", { placeholder: "PH_TITLE", color: theme.colors.primary, fontSize: 32, bold: true });
   ["Madde 1", "Madde 2", "Madde 3"].forEach((t) =>
     s1.addText(t, { placeholder: "PH_BODY", color: theme.colors.text,
                     fontSize: 18, bullet: true, lineSpacing: 28 }));
   deck.writeFile({ fileName: "/workspace/outputs/sunum.pptx" });
   ```
   (`require("pptxgenjs")` resolves via `NODE_PATH` — no local node_modules needed.
   Master placeholders need `options.name`; slides fill them with the same name.
   `slideNumber` on the master handles page numbers — do not draw your own.
   **Placeholders:** one `addText` per paragraph. Passing an array to a
   placeholder merges items into one `<a:p>` with repeated `<a:pPr>`
   (invalid OOXML). Non-placeholder text may use a `{ text }` array.)
2. Run: `node /workspace/scratch/build_deck.js`
3. Repair: `python /opt/docgen/scripts/repair_pptx.py <out.pptx>`
   (fixes known pptxgenjs XML quirks: element order in presentation.xml,
   split merged bullet paragraphs. Idempotent — "clean" when nothing to fix.)
4. QA loop: validate.py → readback.py → Gotenberg render + `pdfplumber`
   page-count check.

## Editing workflow (existing deck)

1. Read: `Document.open(path).to_markdown()` — map slides to content.
2. **Text-only changes:**
   - `unzip -o deck.pptx -d /workspace/scratch/deck_x`
   - edit `ppt/slides/slideN.xml` with `lxml`; text lives in `a:t`
     elements inside `a:p` paragraphs (namespace `a` =
     `http://schemas.openxmlformats.org/drawingml/2006/main`).
   - rezip: `cd /workspace/scratch/deck_x && zip -r ../new.pptx . -x ".*"`
3. **Targeted style changes** ("bu başlığı kırmızı yap", "3. slayttaki
   metin 20 punto olsun", "logoyu sağa al"): same unzip/rezip as step 2,
   but change only the matching shape/runs — do NOT rebuild the deck.
   - Find the `p:sp` whose `a:t` text matches; edit each `a:r/a:rPr`
     (create it as the run's first child if missing):
     - size: attribute `sz` in 1/100 pt (`sz="2000"` = 20 pt)
     - color: child `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>`
       (replace any existing fill child)
     - font: child `<a:latin typeface="Arial"/>`
     - bold/italic: attributes `b="1"` / `i="1"`
     - child order inside `a:rPr`: `a:ln`, fill, effects, `a:highlight`,
       then `a:latin`, `a:ea`, `a:cs` — fill always before `a:latin`.
   - Alignment: `algn="ctr"|"l"|"r"` on that paragraph's `a:pPr`.
   - Position/size: `p:spPr/a:xfrm` → `a:off x= y=` and `a:ext cx= cy=`
     in EMU (914400 EMU = 1 inch). If the shape has no `a:xfrm` it inherits
     from the layout — add one as the first child of `p:spPr`.
   - If the target is only part of a run, split the run (copy its `a:rPr`)
     and style only the target part.
   - An XSD order mistake shows up in `validate.py` as "This element is not
     expected" — fix the order and re-run.
4. **Layout/theme/structural changes:** regenerating is more reliable than
   patching — extract the deck's text content, then rebuild the deck with
   pptxgenjs using the theme. Do this when many slides change or the user
   asks for a restyle.
5. QA: validate → readback → render.

## Quality rules

- Title slide: deck title + subtitle, nothing else.
- Section dividers for decks longer than ~8 slides.
- Charts: use pptxgenjs `addChart` with theme colors (unless the user
  specifies colors); one chart per slide.
- Never deliver a deck where a slide exceeds 6 bullets.
