# XLSX — technical reference

## office_oxide Python API (verified signatures)

```python
from office_oxide import XlsxWriter

wb = XlsxWriter()
s = wb.add_sheet("Bütçe")                          # -> sheet index (int)
wb.set_column_width(s, 0, 32.0)                    # (sheet, col, width)
wb.set_cell(s, 0, 0, "Kalem")                      # (sheet, row, col, value)
                                                  # value: str|bool|int|float|None
wb.set_cell_styled(s, 0, 0, "Kalem", bold=True,
                   bg_color="#1F4E79")             # (sheet,row,col,val,bold,bg)
wb.set_formula(s, 3, 1, "=SUM(B2:B3)")             # formula INCLUDES leading "="
wb.merge_cells(s, 0, 0, 1, 2)                      # (sheet, row, col, row_span, col_span)
wb.save("/workspace/outputs/x.xlsx")
wb.to_bytes()
```

**Row/col are 0-based.** A1 = (row 0, col 0); B3 = (row 2, col 1).
**Formula references inside the string stay A1-notation** (`=SUM(B2:B3)`).

```python
from office_oxide import EditableDocument   # editing an existing xlsx
d = EditableDocument.open("x.xlsx")
d.set_cell(0, "B4", 999)                    # (sheet_index, "A1"-ref string, value)
d.replace_text("eski hücre metni", "yeni")  # works on cell text
d.save("x.xlsx")
```

### Column width cheat (0-based col)
| col idx | ref |
|---|---|
| 0 | A |
| 1 | B |
| 26 | Z |
| 27 | AA |

`openpyxl.utils.get_column_letter(i+1)` converts programmatically.

## openpyxl — key patterns

```python
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment

wb = load_workbook("x.xlsx")            # data_only=False keeps formulas
ws = wb["Bütçe"]
ws["B4"] = "=SUM(B2:B3)"
ws["A1"].font = Font(bold=True, color="1F4E79")
ws["A1"].fill = PatternFill("solid", fgColor="BDD7EE")
ws.column_dimensions["A"].width = 32
wb.save("x.xlsx")
# read values after LibreOffice/Excel calc: load with data_only=True
```

## Formula compatibility table

| Banned | Use instead |
|---|---|
| `XLOOKUP` | `INDEX(rng, MATCH(key, keys, 0))` |
| `XMATCH` | `MATCH` |
| `SORT` | helper column + `RANK`, or pre-sort in Python |
| `FILTER` | `SUMIFS`/`COUNTIFS` + helper columns |
| `UNIQUE` | helper column + `COUNTIF` |
| `SEQUENCE` | literal list or `ROW(INDIRECT("1:10"))` |

`_xlfn.` prefix (write exactly):
`_xlfn.TEXTJOIN` · `_xlfn.CONCAT` · `_xlfn.IFS` · `_xlfn.SWITCH` ·
`_xlfn.MAXIFS` · `_xlfn.MINIFS`

## Color convention (financial models)

| Meaning | Font / fill |
|---|---|
| Input (manual) | font `#0000FF` (blue) |
| Formula | font `#000000` (black) |
| Internal cross-sheet link | font `#008000` (green) |
| External reference | font `#FF0000` (red) |
| Key assumption | fill `#FFFF00` (yellow) |

## QA: recalc_check.py flow

1. `POST $GOTENBERG_URL/forms/libreoffice/convert` (xlsx→PDF) — LibreOffice
   recalculates formulas on load.
2. `pdfplumber` extracts the rendered text.
3. Scan for error tokens: `#REF!`, `#NAME?`, `#DIV/0!`, `#VALUE!`, `#N/A`,
   `#NULL!`, `#NUM!`.
4. Optional `--expect "label:12345.67"` — verifies specific rendered values.
5. JSON report + exit code (0 = clean).

Note: formulas written by office_oxide/openpyxl have **no cached values**;
LibreOffice computes them on load — so the recalc QA is the source of truth.

## OOXML anatomy (xlsx = zip)

```
[Content_Types].xml
_rels/.rels
xl/workbook.xml            sheet list
xl/worksheets/sheetN.xml   rows/cells (c, r, f, v elements)
xl/sharedStrings.xml
xl/styles.xml
xl/_rels/workbook.xml.rels
```
`scripts/validate.py` maps `xl/worksheets/sheet*.xml` → `schemas/sml.xsd`.
