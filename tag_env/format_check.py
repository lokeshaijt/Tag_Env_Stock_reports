"""Compare a generated report with the reference template, cell style by cell style.

Checks sheet names / order / visibility, column widths, frozen panes, header
rows (text, merges and style), and that every data, total and grand-total cell
carries the same font, fill, border, alignment and number format as the
matching cell of the template. Returns a list of differences (empty = identical
format).
"""
from __future__ import annotations

import io

import openpyxl
from openpyxl.utils import get_column_letter

from .template_writer import EXPORT_SHEET, MAIN, MAIN_SHEETS, MAKE, MAKE_SHEET, TEMPLATE, _cols


def _sig(c):
    f, fl, b, a = c.font, c.fill, c.border, c.alignment

    def col(x):
        return None if x is None else (x.type, x.rgb if x.type == "rgb" else x.theme, round(x.tint or 0, 3))
    return (f.name, f.sz, bool(f.b), bool(f.i), col(f.color),
            fl.fill_type, col(fl.fgColor) if fl.fill_type else None,
            b.left.style, b.right.style, b.top.style, b.bottom.style,
            a.horizontal, a.vertical, bool(a.wrap_text), c.number_format)


def _widths(ws):
    return {k: round(v.width or 0, 2) for k, v in ws.column_dimensions.items() if v.width}


def compare(output: bytes) -> list[str]:
    ref = openpyxl.load_workbook(TEMPLATE, keep_links=False)
    out = openpyxl.load_workbook(io.BytesIO(output))
    diffs: list[str] = []

    if ref.sheetnames != out.sheetnames:
        diffs.append(f"Sheet names/order differ: {out.sheetnames} vs {ref.sheetnames}")
        return diffs
    for name in ref.sheetnames:
        r, o = ref[name], out[name]
        if r.sheet_state != o.sheet_state:
            diffs.append(f"{name}: visibility {o.sheet_state} vs {r.sheet_state}")
        if _widths(r) != _widths(o):
            diffs.append(f"{name}: column widths differ")
        rp, op = r.sheet_view.pane, o.sheet_view.pane
        if (rp and (rp.xSplit, rp.ySplit)) != (op and (op.xSplit, op.ySplit)):
            diffs.append(f"{name}: frozen panes differ")

    def same_cell(sheet, ref_ref, out_ref, what):
        rc, oc = ref[sheet][ref_ref], out[sheet][out_ref]
        if _sig(rc) != _sig(oc):
            diffs.append(f"{sheet}!{out_ref}: {what} style differs from template {ref_ref}")

    # header rows: text, merges, style
    for name, hdr_rows, last in [(n, 3, MAIN["last_col"]) for n in MAIN_SHEETS] + \
                                [(MAKE_SHEET, 4, MAKE["last_col"])]:
        for row in range(1, hdr_rows + 1):
            for c in _cols(last):
                ref_cell, out_cell = ref[name][f"{c}{row}"], out[name][f"{c}{row}"]
                if row >= 2 and not (name == MAKE_SHEET and row == 4) and ref_cell.value != out_cell.value:
                    diffs.append(f"{name}!{c}{row}: header '{out_cell.value}' vs '{ref_cell.value}'")
                same_cell(name, f"{c}{row}", f"{c}{row}", "header")
        rm = {str(m) for m in ref[name].merged_cells.ranges if m.max_row <= hdr_rows}
        om = {str(m) for m in out[name].merged_cells.ranges if m.max_row <= hdr_rows}
        if rm != om:
            diffs.append(f"{name}: merged header cells differ")

    # data / total rows
    for name, spec in [(n, MAIN) for n in MAIN_SHEETS] + [(MAKE_SHEET, MAKE)]:
        o = out[name]
        merged = {(get_column_letter(cc), rr) for m in o.merged_cells.ranges if m.min_row >= spec["first"]
                  for rr in range(m.min_row, m.max_row + 1) for cc in range(m.min_col, m.max_col + 1)}
        last_row = o.max_row
        total_row = last_row if name == MAKE_SHEET else last_row - 1
        for row in range(spec["first"], total_row):
            for c in _cols(spec["last_col"]):
                if (c, row) in merged:
                    continue
                same_cell(name, f"{c}{spec['style_row']}", f"{c}{row}", "data")
        for c in _cols(spec["last_col"]):
            same_cell(name, f"{c}{spec['total_row']}", f"{c}{total_row}", "total")
            if name != MAKE_SHEET:
                same_cell(name, f"{c}{spec['grand_row']}", f"{c}{last_row}", "grand total")

    # Ghana & AB & AG Stock: each role keeps the template style
    o = out[EXPORT_SHEET]
    roles = {"GHANA STOCK": 1, "ABHIDJAN STOCK": 16, "ANGOLA STOCK": 22}
    for row in range(1, o.max_row + 1):
        v = o[f"A{row}"].value
        if v in roles:
            cols = "ABC" if v == "GHANA STOCK" else "AB"
            for k in range(3):
                for c in cols:
                    same_cell(EXPORT_SHEET, f"{c}{roles[v] + k}", f"{c}{row + k}", "section header")
        elif v == "GRAND TOTAL":
            for c in "AB":
                same_cell(EXPORT_SHEET, f"{c}29", f"{c}{row}", "grand total")
    return diffs
