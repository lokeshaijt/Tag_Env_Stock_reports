# Tag & Env Stock Report – Streamlit app

Builds the weekly **Tag & Env Stock** report automatically from three ERP exports:

| Input | What it is used for |
|---|---|
| **Stock Summary Report** (`StockSummaryReport-…xls/xlsx`) | Envelope and tag stock in kg, and the machine line of each item |
| **Order Status Report** (`OrderStatusReport-…xlsx`) | Open orders → the *Remarks* column |
| **Exploded BOM Report** (`Exploded_bom-…xlsx`) | Which envelope / tag each finished good uses, and pieces per CFC |

The page has three upload boxes, one per report. Each file is checked by its columns, so a file put
in the wrong box is flagged straight away. Once all three are in, the report is built automatically;
click **Generate Report**, check the preview, optionally change a remark, and download the Excel file.

## Output workbook

The reference report (`tag_env/templates/Tag_Env_Stock_Template.xlsx`, the manually prepared
*Tag_Env_Stock_as_on_28-10-26*) is used as the template. Sheet names, order and visibility, header
rows, merged headers, fonts, fills, borders, number formats, column widths, row heights, frozen
panes and print settings all come from it; only the data rows are replaced, each new cell taking the
style of the template's data cells. Links to other workbooks are removed.

| Sheet | Contents |
|---|---|
| **Env & Tag Stock** (visible) | Make Cases per machine line: kg, `=(kg × factor in row 4) / BOM pieces per case`, *Max Can Make*, total `=F+L+P+T` (as in the reference), Remarks. Kg totals in the last row. |
| **Tag & Env Stock** (hidden, short names) and **Tag & Env Stock (2)** (hidden, full names) | Stock in kg by machine line, shared tags merged across blends, Remarks, `TOTAL` (SUBTOTAL) and `Grand Total`. |
| **Ghana & AB & AG Stock** | Ghana, Abidjan and Angola stock with totals and grand total. |

### Make Cases

`Make Cases = (Available KG × Conversion Factor) / Exploded BOM value`

* **Conversion factor** – per machine line, from the *Conversion Round Off* chart
  (`config.DEFAULT_CONVERSIONS`, editable in the app sidebar), written in row 4 above each
  Make Cases column. Each machine line uses its own factor.
* **Exploded BOM value** – pieces of that envelope/tag per case (CFC), read from the uploaded BOM file,
  never hard-coded: the FG on open order first; otherwise the blend's own case pack (standard pack
  on a tie); otherwise the largest case pack. Items missing from the *Exploded BOM* sheet are looked
  up in the file's single-level *BOM* sheet (cartons per case × pieces per carton). BOM quantities
  below one piece are ignored as data errors; a tag without a usable value takes its envelope's
  value on the same row (reported).
* Missing factor or BOM value → the cell is left empty and listed in *Data checks*.

The app also compares every generated workbook with the template (*Format check* tab) and the tests
assert there are no differences.

## Business rules

**Rows.** One row per blend/flavour family. The family is the item name without the machine-line
words (`DC`, `SF`, `C250`, `ENV`, `TBGS`, `PENV`, `HS`, `PTAG`, `PK45`, `-FT`, `(70 GSM)`, `(WO GLUE)`, `(v2)`),
so `ENV AL FINA DC SF C250 TBGS`, `ENV AL FINA DC ENV (70 GSM)`, `ENV AL FINA DC SF TBGS`,
`TAG AL FINA DC SF C250 TBGS` and `TAG AL FINA DC SF TBGS` share one row. The row is named after its
Perfecta, else C250, else MD-20, else MD HS, else Maisa, else Constanta envelope. Rows are A–Z, then
tag-only rows (PK45/Perfecta tags, MD-20 tags, Constanta tags).

**Columns.** By the *Machine Line Name* in the Stock Summary:
Constanta Envelope → Constanta; CONSTANTA-250 → C250 Envelope; Constanta Tag → C250 Tag;
MD-20 Crimp Envelope / MD-20 Tag → MD20; Perfecta Envelope (PENV) → Perfecta; MD HS Envelope → MD20 HS;
Maisa Envelope / Maisa Tag → Maisa; PERFACTA TAG / PK-45 → Perfecta/PK45 Tag.

**Tags.** A tag joins the row of the same name when that row has an envelope on the matching line
(Constanta tag ↔ C250 envelope, MD-20 tag ↔ MD-20 envelope …). Otherwise the BOM decides which
blends it is packed with; a tag used by several blends is merged across them. Tags that match no
envelope get their own row at the bottom.

**Export stock.** Items marked `(GHANA)`, `(ABIDJAN)` or `(ANGOLA)`, and `ENV INDUS DC TBGS (70GSM)`,
go to the *Ghana & AB & AG Stock* sheet instead of the main sheet.

**Remarks** (from open orders of every finished good that uses the row's envelopes/tags):

1. **Running** – production still pending and already started (stock in process or ready stock > 0)
2. **Order Received** – production still pending, nothing made yet
3. **Multiline Stock** – nothing pending, stock on two or more envelope lines
4. **Leftover Stock** – nothing pending (no order, or orders fully covered by ready stock)

The difference between *Running* and *Order Received* depends on the shop-floor plan, which is not in
the three files, so any remark can be changed in the preview before downloading.

## Data checks (*Data checks* tab in the app)

* items whose machine line is unknown (left out, with their kg)
* duplicate items in the Stock Summary (quantities added) and negative stock
* several stock items combined in one cell
* order FGs missing from the BOM (matched by blend name where possible, otherwise listed)
* reconciliation: total envelope/tag kg in the Stock Summary must equal the kg placed in the report

## Run locally

```bash
pip install -r requirements.txt
```

```bash
streamlit run app.py
```

Tests: `python -m pytest -q`

## Deploy on Streamlit Community Cloud

1. Push this repository to GitHub (already done if you are reading this there).
2. Go to <https://share.streamlit.io>, sign in with GitHub and click **Create app**.
3. Pick this repository, branch `main` (or the branch you use) and main file `app.py`, then **Deploy**.
4. Share the app URL with the planning team. The upload limit is set to 200 MB in
   `.streamlit/config.toml` (the Exploded BOM is ~13 MB).

## Changing the rules

Everything a planner may want to change is in [`tag_env/config.py`](tag_env/config.py): machine-line
mapping, conversion factors, export sections and extra export items, and the words stripped to
build the family key.

## Project layout

```
app.py                         Streamlit UI (3 uploads, Generate Report, preview, download)
tag_env/readers.py             readers for the 3 ERP exports (.xls/.xlsx/HTML, header detection)
tag_env/grouping.py            stock lines -> report rows, shared tags, export stock
tag_env/remarks.py             order linking, remarks, BOM pieces per case
tag_env/make_cases.py          Make Cases calculation
tag_env/template_writer.py     fills the reference workbook (openpyxl)
tag_env/format_check.py        compares the output with the reference format
tag_env/templates/             the reference Tag & Env Stock workbook
tests/                         synthetic end-to-end tests
```
