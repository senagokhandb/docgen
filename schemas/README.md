# ECMA-376 / ISO 29500 XSD schemas

`scripts/validate.py` validates office files against the ECMA-376
transitional schemas in this directory. The original directory layout is
kept so inter-schema `xs:import` locations (`../dml/...`, `../shared/...`)
resolve natively — do **not** flatten this tree.

## Entry schemas used by validate.py

| File type | Parts validated | Entry schema |
|---|---|---|
| .docx | `word/document.xml`, `word/styles.xml`, `word/numbering.xml`, headers/footers | `wml/wml__ROOT.xsd` |
| .xlsx | `xl/worksheets/sheetN.xml` | `sml/sml_root.xsd` |
| .pptx | `ppt/presentation.xml`, `ppt/slides/slideN.xml`, layouts, masters | `pml/pml-2ed.xsd` |

Each root schema transitively imports the extension schemas (w14/w15/w16
for WordML, 2010+ revisions for SpreadsheetML, dml/shared for all), so one
entry point validates the full transitional namespace set.

## Provenance

Fetched 2026-07-21 from the open-source **docx4j** project
(https://github.com/plutext/docx4j, `xsd/` directory, Apache-2.0), which
mirrors the ECMA-376 5th edition (ISO/IEC 29500-1) standard schemas.
178 files. FOP-related directories (`FopConfig`, `xslfo`) removed.

## Fallback behavior

If the schemas are missing from a container, `validate.py` automatically
falls back to **structural mode**:

- zip integrity (`zipfile.testzip`)
- `[Content_Types].xml` present
- every `.xml` / `.rels` part well-formed (lxml)
- PDF: `%PDF` header + page count

No code change is needed to toggle modes — presence of the entry `.xsd`
files activates XSD validation.
