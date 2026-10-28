# XLSX — technical reference

## docgen.xlsx — in-place editing API

```python
from docgen.package import Package
from docgen import xlsx as X, one
pkg = Package(path)                         # .xlsx; templates: Package(p).as_document()

# --- read / locate -----------------------------------------------------
X.sheets(pkg)                               # [(name, part, state)]
X.get(pkg, "Satışlar!E2")                   # (value, "=formula" | None); value = cached result for formulas
X.cells(pkg, "Satışlar")                    # iterate (ref, value, formula)
X.find(pkg, "Başarısız", sheet="Durum", exact=False)      # ['Durum!B3', ...]
X.column_of(pkg, "Durum", "Not")            # column letter of a header
X.rows_where(pkg, "Durum", "Servis", "Servis B")           # row numbers by key column (letter or header)

# --- write -------------------------------------------------------------
X.set_value(pkg, "Satışlar!E2", 18, expect=16, allow_formula_overwrite=False)   # str/int/float/bool/None
X.set_formula(pkg, "Özet!B12", "=SUM(B4:B9)")
X.formula_to_value(pkg, "Özet!B4:B9", force=True)          # only with the user's --onay (the library checks it)
X.set_style(pkg, "Durum!A1:C1",
            font={"bold": True, "italic": False, "underline": False, "strike": False,
                  "color": "FFFFFF", "size": 14, "name": "Arial"},
            fill="1F3864",                  # or False = no fill
            number_format="#,##0.00",
            align={"horizontal": "center", "vertical": "center", "wrap": True})
X.add_conditional_format(pkg, "Satışlar!E2:E41", operator="greaterThan", formula="20", fill="C6EFCE")
X.add_conditional_format(pkg, "Durum!A2:C9", expression='=$B2="Başarısız"', fill="F8CBAD")
X.set_column_width(pkg, "Satışlar", "B", 24)
X.dependents_of_rows(pkg, "Parametreler", 2, 2)            # who reads these rows (before deleting)
X.features(pkg)                                             # tables/CF/DV/merged/charts/images/... per sheet
pkg.save(dst)
```
Every value/formula change sets `fullCalcOnLoad` and drops `calcChain.xml`;
run `xlsx_recalc.py` afterwards so the stored results are correct.
Errors: `NeedsConfirmation`, `docgen.xlsx.ValueMismatch`, `Ambiguous`, `NotFound`.

## Structural operations — `scripts/xlsx_struct.py` (docgen.xlsx_struct)

`insert-rows/delete-rows/insert-cols/delete-cols/sort` edit the sheet XML in
place and shift every reference to that sheet: formulas on all sheets
(shared/array included), defined names and print titles, conditional
formats and validations (ranges + formulas), tables, merged cells,
hyperlinks, auto-filters, comments, drawing anchors and chart series.
Refusals: deleting cells that are read elsewhere (needs `--onay`), sorting
rows whose formulas read other rows, merged cells inside a sort range,
pivot tables. Afterwards it runs `xlsx_recalc.py` itself.

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

## QA: xlsx_recalc.py flow (primary)

1. LibreOffice (headless, in the container, profile forces "always
   recalculate") converts a copy to xlsx — i.e. computes every formula.
2. The computed results are written into the ORIGINAL package as cached
   values (`<v>` + type); nothing else is rewritten.
3. Report: changed results, error cells (`#REF!`, `#DIV/0!`…), NEW errors
   (exit 1), `--expect "Sheet!B4=20000"` checks.

`recalc_check.py` (Gotenberg → PDF → text scan) remains as a fallback when
LibreOffice is not available; it cannot write results back.

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
