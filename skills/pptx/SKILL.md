---
name: pptx
description: Create and edit PowerPoint (.pptx) presentations. Use for decks, pitch materials, slide-based output, to modify an existing .pptx, or to build a deck on a user's template (.potx/.pptx).
---

# PPTX skill

## Tooling

| Need | Tool |
|---|---|
| Read a deck (text / formatting, positions) | `readback.py <file>` / `readback.py <file> --format` |
| Where does a text occur? | `find_target.py <file> "<text>" [--slide N]` |
| **Edit an uploaded deck** (text, formatting, position/size, fill, background, notes, tables, slide delete/move/duplicate/hide) | Python `docgen.pptx` + `docgen.package.Package` (REFERENCE.md §docgen.pptx) |
| Theme colors / fonts of a deck | `docgen.theme.set_color` / `set_fonts` |
| New deck, text + bullets, theme applied automatically | `node /opt/docgen/scripts/build_themed_deck.js <theme.json> <content.json> <out.pptx>` |
| New deck, full control (charts, images, custom layouts) | Node `pptxgenjs` script in `/workspace/scratch/` |
| New deck on a **user template** | `docgen.pptx.start_from_template` + python-pptx layouts |
| .ppt / .odp input | `convert_legacy.py` first, then edit the .pptx |
| Render check (pages == visible slides, content, no text drawn over text) | `python /opt/docgen/scripts/render_check.py <file> [--baseline <original>]` |
| Validation | `validate.py <file> [--baseline <original>]`, `diff_check.py <original> <result> ...` |

pptxgenjs and office_oxide cannot open an existing deck: never "edit" an
uploaded deck by rebuilding it with them.

## Theme workflow (required for new decks without a user template)

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
4. Verify: validate → readback (all slide titles/text present) →
   `render_check.py` (PDF page count == visible slides, no text drawn over
   other text).

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
   then validate → readback → `render_check.py` (page count == visible
   slides, no text drawn over other text).

Use a custom `pptxgenjs` script (new decks only) whenever the deck needs charts,
images, tables, or non-standard layouts, or the user asks for specific
colors, fonts, sizes or positions (the builder only applies theme tokens).
In pptxgenjs every `addText` takes `x`, `y`, `w`, `h` (inches), `color`
(hex, no `#`), `fontFace`, `fontSize` (pt), `bold`, `align`; to style part
of a line, pass a run array:
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
   s1.addText(["Madde 1", "Madde 2", "Madde 3"].map((t) =>
       ({ text: t, options: { bullet: true, lineSpacing: 28, breakLine: true } })),
     { placeholder: "PH_BODY", color: theme.colors.text, fontSize: 18 });
   deck.writeFile({ fileName: "/workspace/outputs/sunum.pptx" });
   ```
   (`require("pptxgenjs")` resolves via `NODE_PATH` — no local node_modules needed.
   Master placeholders need `options.name`; slides fill them with the same name.
   `slideNumber` on the master handles page numbers — do not draw your own.
   **One text box = one `addText`** with all of its paragraphs as an array.
   Every call adds a new shape: bullets added one call at a time into a
   placeholder all sit on its spot and are drawn over each other. Each
   paragraph is one item, `{ text, options: { bullet: true, breakLine: true } }`;
   put paragraph options (`lineSpacing`, `align`, `paraSpaceAfter`) on the
   item — on a placeholder pptxgenjs drops a `lineSpacing` given for the
   whole call. Part of a paragraph in another style: more runs before the
   `breakLine`, e.g. `{ text: "Kalın: ", options: { bullet: true, bold: true } },
   { text: "devamı", options: { breakLine: true } }`.)
2. Run: `node /workspace/scratch/build_deck.js`
3. Repair: `python /opt/docgen/scripts/repair_pptx.py <out.pptx>`
   (fixes known pptxgenjs XML quirks: element order in presentation.xml,
   the extra `<a:pPr>` pptxgenjs writes before each later run of a
   paragraph. Idempotent — "clean" when nothing to fix.)
4. QA loop: validate.py → readback.py → render_check.py (page count ==
   visible slides, no text drawn over other text).

## Editing workflow (uploaded deck)

Never rebuild an uploaded deck; change only what was asked. Charts,
SmartArt, images, hidden slides, links and notes must survive untouched.

1. `readback.py <file> --format` to map slides/shapes; `find_target.py` for
   the target. Use the user's context (slide number/title, title vs body,
   table row). Still several matches → report candidates, write no file.
2. Write `/workspace/scratch/edit.py` with the helpers:
   ```python
   from docgen.package import Package
   from docgen import pptx as P, one
   pkg = Package("/workspace/uploads/sunum.pptx")

   sl = P.slide(pkg, title="Sonuç")                       # or P.slide(pkg, 6)
   P.replace(one(P.find(pkg, "Büyüme sürüyor", slide=sl.index), "madde"), "Büyüme 2026'da da sürecek")

   sh = one(P.find_shapes(pkg, slide=2, text="Kaynak: İç analiz"), "kutu")
   P.set_run_props(P.runs_of(sh), size=10, italic=True)          # bold/underline/strike/color/font too
   m = one(P.find(pkg, "risk", slide=1, in_title=False, whole_word=True), "risk")
   P.set_run_props(P.split_runs(m), underline=True)              # only that word

   W, H = P.slide_size(pkg); x, y, w, h = sh.xfrm
   P.set_geometry(sh, x=W - w - P.cm(1), y=P.cm(1))              # move / resize (EMU; cm(), inch())
   pkg.save("/workspace/outputs/sunum-duzenlenmis.pptx")
   ```
   More: `set_paragraph_props(P.paragraphs_of(sh), align="right")`,
   `scale(pic, 0.5)`, `set_fill(sh, "FF8800")`, `set_background(sl, "EDEDED")`
   (one slide only), `set_notes(pkg, 3, text)`, tables (`tables`, `table_rows`,
   `set_cell_text`, `cell` + `set_run_props`/`set_paragraph_props`,
   `set_cell_fill`, `add_column`), slides (`delete_slide`, `move_slide`,
   `duplicate_slide`, `set_hidden`), theme (`docgen.theme.set_color(pkg, "accent1", "800020")`).
   - `P.replace` changes ONE occurrence; for "everywhere" loop over
     `P.find(...)` and report the count.
   - `delete_slide` refuses (`NeedsConfirmation`) when another slide links
     to it — needs `--onay`.
3. Verify:
   ```bash
   python /opt/docgen/scripts/validate.py OUT --baseline SRC
   python /opt/docgen/scripts/diff_check.py SRC OUT --allow-slide 2 --format-only \
          --record /workspace/outputs/manifest.json
   python /opt/docgen/scripts/readback.py OUT --format
   python /opt/docgen/scripts/render_check.py OUT --baseline SRC
   ```
   Scope flags: `--allow-slide N`, `--allow-notes N`, `--allow-text T`,
   `--allow-structure` (slide add/delete/move), `--allow-theme`;
   `--text-only` / `--format-only`.

## User template (.potx / .pptx given as template)

The template's masters, layouts, theme and fonts define the look; docgen
themes and `build_themed_deck.js` do NOT apply.

```python
from docgen.pptx import start_from_template
from pptx import Presentation                      # python-pptx
out = "/workspace/outputs/sunum.pptx"
start_from_template("/workspace/uploads/kurumsal.potx", out)   # example slides removed
prs = Presentation(out)
print([l.name for l in prs.slide_layouts])         # pick layouts by name
s = prs.slides.add_slide(next(l for l in prs.slide_layouts if l.name == "Title and Content"))
s.shapes.title.text = "Pazar görünümü"
s.placeholders[1].text_frame.text = "Talep büyüyor"
prs.save(out)
```
Keep the template's example slides instead (`keep_slides=True`) when the user
wants to fill them; then edit them with `docgen.pptx` as above. Verify with
`validate.py out --baseline <template>` and `render_check.py out`.

## Quality rules

- Title slide: deck title + subtitle, nothing else.
- Section dividers for decks longer than ~8 slides.
- Charts: use pptxgenjs `addChart` with theme colors (unless the user
  specifies colors); one chart per slide.
- Never deliver a deck where a slide exceeds 6 bullets.
