"""Fill the reference Tag & Env Stock workbook with new data.

The manually prepared report (templates/Tag_Env_Stock_Template.xlsx) is used as
the template: sheet names, order and visibility, header rows, merged headers,
fonts, fills, borders, number formats, column widths, row heights, frozen panes
and print settings all come from it. Only the data rows are replaced, each new
cell taking the style of the matching cell in the template's data rows.
"""
from __future__ import annotations

import io
from copy import copy
from datetime import date, datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter

from . import config as C
from .make_cases import BLOCKS, TOTAL_OF, calculate, conversion_for

TEMPLATE = Path(__file__).parent / "templates" / "Tag_Env_Stock_Template.xlsx"

MAIN_SHEETS = ["Tag & Env Stock (2)", "Tag & Env Stock"]
MAKE_SHEET = "Env & Tag Stock"
EXPORT_SHEET = "Ghana & AB & AG Stock"

# Where things are in the template
MAIN = {"first": 4, "style_row": 5, "shared_cell": "E30", "total_row": 187, "grand_row": 188,
        "last_col": "S"}
MAKE = {"first": 5, "style_row": 6, "shared_kg": "B31", "shared_cases": "C31", "total_row": 188,
        "last_col": "AM"}

# Short names used on the "Tag & Env Stock" and "Env & Tag Stock" sheets
_SHORTEN = ["DC SF C250 ENV TBGS", "DC SF C250 TBGS", "DC SF HS ENV TBGS", "DC SF ENV TBGS",
            "DC C250 TBGS", "DC ENV TBGS", "DC PK45 TBGS", "C250 TAG TBGS", "DC TBGS"]


def short_name(name: str) -> str:
    """'ENV AL FINA DC SF C250 TBGS' -> 'ENV AL FINA ' (the manual report's short form)."""
    s = name.replace(" (WO GLUE)", "")
    for phrase in _SHORTEN:
        if phrase in s:
            s = s.replace(phrase, "", 1)
            s = s.replace(" (70GSM)", "").replace(" (70 GSM)", "")
            break
    return s


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def _cols(last: str) -> list[str]:
    return [get_column_letter(i) for i in range(1, column_index_from_string(last) + 1)]


def _row_styles(ws, row: int, last: str) -> dict[str, object]:
    return {c: copy(ws[f"{c}{row}"]._style) for c in _cols(last)}


def _clear_from(ws, first: int):
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row >= first:
            ws.unmerge_cells(str(rng))
    if ws.max_row >= first:
        ws.delete_rows(first, ws.max_row - first + 1)
    for r in [r for r in ws.row_dimensions if r >= first]:
        del ws.row_dimensions[r]


def _put(ws, ref: str, value, style):
    cell = ws[ref]
    cell.value = value
    cell._style = copy(style)
    return cell


def _merge(ws, col: str, r1: int, r2: int, style):
    for r in range(r1, r2 + 1):
        ws[f"{col}{r}"]._style = copy(style)
    if r2 > r1:
        ws.merge_cells(f"{col}{r1}:{col}{r2}")


def _scroll_top(ws, cell: str):
    sv = ws.sheet_view
    if sv.pane is not None:
        sv.pane.topLeftCell = cell
    for sel in sv.selection:
        sel.activeCell = cell
        sel.sqref = cell


def _dt(d: date | None) -> datetime:
    d = d or date.today()
    return datetime(d.year, d.month, d.day)


# ---------------------------------------------------------------------------
# Tag & Env Stock (and its "(2)" copy with full item names)
# ---------------------------------------------------------------------------
def _fill_main(ws, report, remarks, full_names: bool):
    data = _row_styles(ws, MAIN["style_row"], MAIN["last_col"])
    shared = copy(ws[MAIN["shared_cell"]]._style)
    total = _row_styles(ws, MAIN["total_row"], MAIN["last_col"])
    grand = _row_styles(ws, MAIN["grand_row"], MAIN["last_col"])
    h_data = ws.row_dimensions[MAIN["style_row"]].height
    h_grand = ws.row_dimensions[MAIN["grand_row"]].height
    _clear_from(ws, MAIN["first"])

    ws["A1"].value = _dt(report.report_date)
    first = MAIN["first"]
    last = first + len(report.rows) - 1
    for i, row in enumerate(report.rows):
        r = first + i
        for col, st in data.items():
            ws[f"{col}{r}"]._style = copy(st)
        ws[f"A{r}"].value = row.name if full_names else short_name(row.name)
        for cid, (letter, *_x) in C.MAIN_COLUMNS.items():
            ws[f"{letter}{r}"].value = round(row.qty(cid), 2)
        ws[f"S{r}"].value = remarks.get(row.key, row.remark)
        ws.row_dimensions[r].height = h_data
    for st, lo, hi in report.shared_spans():
        letter = C.MAIN_COLUMNS[st.column][0]
        _merge(ws, letter, first + lo, first + hi, shared)
        ws[f"{letter}{first + lo}"].value = round(st.qty, 2)

    t, g = last + 1, last + 2
    for col in _cols(MAIN["last_col"]):
        ws[f"{col}{t}"]._style = copy(total[col])
        ws[f"{col}{g}"]._style = copy(grand[col])
    ws[f"A{t}"].value = "TOTAL"
    for letter in (v[0] for v in C.MAIN_COLUMNS.values()):
        ws[f"{letter}{t}"].value = f"=SUBTOTAL(9,{letter}{first}:{letter}{last})"
    ws[f"A{g}"].value = "Grand Total"
    ws[f"B{g}"].value = "=" + "+".join(f"{v[0]}{t}" for v in C.MAIN_COLUMNS.values())
    ws.row_dimensions[g].height = h_grand
    ws.auto_filter.ref = f"R2:S{last}"
    _scroll_top(ws, "B4")


# ---------------------------------------------------------------------------
# Env & Tag Stock (make cases)
# ---------------------------------------------------------------------------
def _fill_make(ws, report, remarks, conversions, cells):
    data = _row_styles(ws, MAKE["style_row"], MAKE["last_col"])
    s_kg = copy(ws[MAKE["shared_kg"]]._style)
    s_cases = copy(ws[MAKE["shared_cases"]]._style)
    total = _row_styles(ws, MAKE["total_row"], MAKE["last_col"])
    h_data = ws.row_dimensions[MAKE["style_row"]].height
    _clear_from(ws, MAKE["first"])

    ws["A1"].value = _dt(report.report_date)
    # Row 4: conversion factor of each machine line, above its make-cases column
    for _, cols, _ in BLOCKS:
        for cid, _, cm_col in cols:
            ws[f"{cm_col}4"].value = conversion_for(cid, conversions)

    first = MAKE["first"]
    last = first + len(report.rows) - 1
    for i, row in enumerate(report.rows):
        r = first + i
        for col, st in data.items():
            ws[f"{col}{r}"]._style = copy(st)
        ws.row_dimensions[r].height = h_data
        ws[f"A{r}"].value = short_name(row.name)
        for _, cols, max_col in BLOCKS:
            for cid, kg_col, cm_col in cols:
                cell = cells[(i, cid)]
                if cell.covered:
                    continue
                ws[f"{kg_col}{r}"].value = round(cell.kg, 2)
                if cell.cases is None:
                    ws[f"{cm_col}{r}"].value = None
                elif cell.bom_value:
                    ws[f"{cm_col}{r}"].value = f"=({kg_col}{r}*${cm_col}$4)/{cell.bom_value:g}"
                else:
                    ws[f"{cm_col}{r}"].value = 0
                if cell.merged_to is not None:
                    _merge(ws, kg_col, r, first + cell.merged_to, s_kg)
                    _merge(ws, cm_col, r, first + cell.merged_to, s_cases)
            parts = [f"{cm}{r}" for _, _, cm in cols]
            ws[f"{max_col}{r}"].value = (f"=MAX({','.join(parts)})" if len(parts) > 1
                                         else f"=MAX({parts[0]})" if max_col == "P" else f"={parts[0]}")
        ws[f"AL{r}"].value = "=" + "+".join(f"{c}{r}" for c in TOTAL_OF)
        ws[f"AM{r}"].value = remarks.get(row.key, row.remark)

    t = last + 1
    for col in _cols(MAKE["last_col"]):
        ws[f"{col}{t}"]._style = copy(total[col])
    for _, cols, max_col in BLOCKS:
        for _, kg_col, _ in cols:
            ws[f"{kg_col}{t}"].value = f"=SUM({kg_col}{first}:{kg_col}{last})"
    ws[f"F{t}"].value = " "
    _scroll_top(ws, "B5")


# ---------------------------------------------------------------------------
# Ghana & AB & AG Stock
# ---------------------------------------------------------------------------
def _fill_export(ws, report):
    def grab(row, cols="ABC"):
        return {c: copy(ws[f"{c}{row}"]._style) for c in cols}

    st = {
        3: {"title": grab(1), "h1": grab(2), "h2": grab(3), "data": grab(4), "total": grab(14)},
        2: {"title": grab(16, "AB"), "h1": grab(17, "AB"), "h2": grab(18, "AB"), "data": grab(19, "AB"),
            "total": grab(20, "AB")},
    }
    heights = {k: ws.row_dimensions[r].height for k, r in
               (("title", 1), ("h1", 2), ("h2", 3), ("data", 4), ("total", 14), ("grand", 29))}
    grand = grab(29, "AB")
    _clear_from(ws, 1)

    r = 1
    total_refs = []
    for title, _marker in C.EXPORT_SECTIONS:
        lines = [l for l in report.export_lines if l.section == title]
        ncol = 3 if title == "GHANA STOCK" or any(l.column == "tag" for l in lines) else 2
        s = st[ncol]
        last_c = "C" if ncol == 3 else "B"
        for c in s["title"]:
            ws[f"{c}{r}"]._style = copy(s["title"][c])
        ws[f"A{r}"].value = title
        ws.merge_cells(f"A{r}:{last_c}{r}")
        ws.row_dimensions[r].height = heights["title"]
        for c in s["h1"]:
            ws[f"{c}{r + 1}"]._style = copy(s["h1"][c])
            ws[f"{c}{r + 2}"]._style = copy(s["h2"][c])
        ws[f"A{r + 1}"].value = "Brand"
        ws[f"B{r + 1}"].value = "Constanta"
        ws[f"B{r + 2}"].value = "Envelope in Kgs"
        if ncol == 3:
            ws[f"C{r + 1}"].value = "Constanta Tag"
            ws[f"C{r + 2}"].value = "Tag in Kgs"
        ws.merge_cells(f"A{r + 1}:A{r + 2}")
        ws.row_dimensions[r + 1].height = heights["h1"]
        r += 3
        start = r
        extras = set(C.EXPORT_EXTRA_ITEMS)
        envs = sorted([l for l in lines if l.column == "env"],
                      key=lambda l: (l.item.upper() in extras, l.item.upper()))
        tags = sorted([l for l in lines if l.column == "tag"], key=lambda l: l.item.upper())
        for l in envs + tags:
            for c in s["data"]:
                ws[f"{c}{r}"]._style = copy(s["data"][c])
            ws[f"A{r}"].value = l.item
            ws[f"B{r}"].value = round(l.qty, 2) if l.column == "env" else 0
            if ncol == 3:
                ws[f"C{r}"].value = round(l.qty, 2) if l.column == "tag" else 0
            ws.row_dimensions[r].height = heights["data"]
            r += 1
        for c in s["total"]:
            ws[f"{c}{r}"]._style = copy(s["total"][c])
        ws[f"A{r}"].value = "Total" if ncol == 3 else "TOTAL"
        end = max(r - 1, start)
        ws[f"B{r}"].value = f"=SUM(B{start}:B{end})" if end > start else f"=SUM(B{start})"
        total_refs.append(f"B{r}")
        if ncol == 3:
            ws[f"C{r}"].value = f"=SUM(C{start}:C{end})"
            total_refs.append(f"C{r}")
        ws.row_dimensions[r].height = heights["total"]
        r += 2
    r += 1
    for c in grand:
        ws[f"{c}{r}"]._style = copy(grand[c])
    ws[f"A{r}"].value = "GRAND TOTAL"
    ws[f"B{r}"].value = "=" + "+".join(total_refs)
    ws.row_dimensions[r].height = heights["grand"]


# ---------------------------------------------------------------------------
def build_workbook(report, remarks: dict[str, str] | None = None,
                   conversions: dict[str, float] | None = None) -> bytes:
    remarks = remarks or {}
    conv = dict(C.DEFAULT_CONVERSIONS)
    conv.update(conversions or {})
    cells, _ = calculate(report, conv)

    wb = openpyxl.load_workbook(TEMPLATE, keep_links=False)
    for name in MAIN_SHEETS:
        _fill_main(wb[name], report, remarks, full_names=name.endswith("(2)"))
    _fill_make(wb[MAKE_SHEET], report, remarks, conv, cells)
    _fill_export(wb[EXPORT_SHEET], report)
    wb.calculation.fullCalcOnLoad = True
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
