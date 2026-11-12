"""Excel output in the layout of the manually prepared Tag & Env Stock report.

Written with XlsxWriter so every formula also carries its calculated value -
the file shows correct numbers even in viewers that do not recalculate
(mobile previews, e-mail attachments), and stays live when opened in Excel.
"""
from __future__ import annotations

import colorsys
import io
from datetime import date

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

from . import config as C

# ---------------------------------------------------------------------------
# Colours: the manual report uses Office theme colours with tints; these are
# the same colours resolved to RGB.
# ---------------------------------------------------------------------------
_THEME = {3: "44546A", 4: "4472C4", 9: "70AD47"}


def _tint(hex_rgb: str, tint: float) -> str:
    r, g, b = (int(hex_rgb[i:i + 2], 16) / 255 for i in (0, 2, 4))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = l * (1 - tint) + tint if tint > 0 else l * (1 + tint)
    r, g, b = colorsys.hls_to_rgb(h, l, s)
    return "#" + "".join(f"{round(x * 255):02X}" for x in (r, g, b))


HEAD_GREEN = _tint(_THEME[9], 0.7999816888943144)    # brand headers, date cell
HEAD_GREY = _tint(_THEME[3], 0.8999908444471572)     # sub headers
TOTAL_BLUE = _tint(_THEME[4], 0.5999938962981048)    # TOTAL row
GRAND_GREEN = _tint(_THEME[9], 0.5999938962981048)   # Grand Total row
SPACER_YELLOW = "#FFFF00"

NUM0 = '_ * #,##0_ ;_ * \\-#,##0_ ;_ * "-"??_ ;_ @_ '
NUM2 = '_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * "-"??_ ;_ @_ '

MAIN_WIDTHS = {"A": 45.22, "B": 9.22, "C": 1.78, "D": 9.11, "E": 7.0, "F": 1.78, "G": 9.89,
               "H": 7.11, "I": 1.78, "J": 9.66, "K": 1.78, "L": 9.55, "M": 7.0, "N": 1.78,
               "O": 9.0, "P": 7.11, "Q": 1.78, "R": 17.22, "S": 18.22}
MAIN_SPACERS = ["C", "F", "I", "K", "N", "Q"]


class _Fmt:
    """Cache of XlsxWriter formats built from keyword properties."""

    def __init__(self, wb):
        self.wb = wb
        self.cache = {}

    def __call__(self, **props):
        base = {"font_name": "Calibri", "font_size": 11}
        base.update(props)
        key = tuple(sorted(base.items()))
        if key not in self.cache:
            self.cache[key] = self.wb.add_format(base)
        return self.cache[key]


def _col(letter: str) -> int:
    n = 0
    for ch in letter:
        n = n * 26 + ord(ch) - 64
    return n - 1


# ---------------------------------------------------------------------------
# Sheet 1: Tag & Env Stock
# ---------------------------------------------------------------------------
def _write_main(wb, F, report, remarks):
    ws = wb.add_worksheet("Tag & Env Stock")
    rows = report.rows
    first, last = 3, 3 + len(rows) - 1          # zero based data rows (Excel 4..)
    total_r, grand_r = last + 1, last + 2

    for letter, w in MAIN_WIDTHS.items():
        ws.set_column(f"{letter}:{letter}", w)
    ws.set_row(1, 15)
    ws.set_row(2, 28.8)
    ws.set_row(grand_r, 19.95)

    # Row 1 - report date
    ws.write_datetime(0, 0, _dt(report.report_date), F(num_format="mm-dd-yy", bg_color=HEAD_GREEN))
    for c in (1, 2):
        ws.write_blank(0, c, None, F(bg_color=HEAD_GREEN))

    # Rows 2-3 - headers
    border = {"border": 1}
    ws.merge_range(1, 0, 2, 0, "Brand", F(bg_color=HEAD_GREY, align="left", valign="vcenter", **border))
    brand_cells: dict[str, list[str]] = {}
    for cid, (letter, brand, sub, kind) in C.MAIN_COLUMNS.items():
        brand_cells.setdefault(brand, []).append(letter)
        ws.write(2, _col(letter), sub, F(bg_color=HEAD_GREY, align="center", valign="vcenter",
                                          text_wrap=True, **border))
    for brand, letters in brand_cells.items():
        a, b = _col(letters[0]), _col(letters[-1])
        fmt = F(bold=True, bg_color=HEAD_GREEN, align="center", valign="vcenter", **border)
        if a == b:
            ws.write(1, a, brand, fmt)
        else:
            ws.merge_range(1, a, 1, b, brand, fmt)
    ws.merge_range(1, _col("S"), 2, _col("S"), "Remarks",
                   F(bg_color=HEAD_GREY, align="left", valign="vcenter", **border))
    for letter in MAIN_SPACERS:
        for r in range(1, total_r + 1):
            ws.write_blank(r, _col(letter), None, F(bg_color=SPACER_YELLOW, **border))

    # Data
    cell_fmt = {cid: F(num_format=NUM0 if kind == "env" else NUM2, **border)
                for cid, (_, _, _, kind) in C.MAIN_COLUMNS.items()}
    for i, row in enumerate(rows):
        r = first + i
        ws.write(r, 0, row.name, F(**border))
        for cid, (letter, *_rest) in C.MAIN_COLUMNS.items():
            ws.write_number(r, _col(letter), round(row.qty(cid), 2), cell_fmt[cid])
        ws.write(r, _col("S"), remarks.get(row.key, row.remark), F(**border))

    # Shared tags: one value merged over the families it serves
    for st, lo, hi in report.shared_spans():
        letter = C.MAIN_COLUMNS[st.column][0]
        fmt = F(num_format=NUM2, align="right", valign="vcenter", **border)
        if hi > lo:
            ws.merge_range(first + lo, _col(letter), first + hi, _col(letter), round(st.qty, 2), fmt)
        else:
            ws.write_number(first + lo, _col(letter), round(st.qty, 2), fmt)

    # Totals
    tot_fmt = F(bold=True, bg_color=TOTAL_BLUE, num_format=NUM0, **border)
    ws.write(total_r, 0, "TOTAL", F(bold=True, bg_color=TOTAL_BLUE, **border))
    totals = report.column_totals()
    for cid, (letter, *_rest) in C.MAIN_COLUMNS.items():
        col = _col(letter)
        ws.write_formula(total_r, col, f"=SUBTOTAL(9,{letter}{first + 1}:{letter}{last + 1})",
                         tot_fmt, round(totals[cid], 2))
    ws.write_blank(total_r, _col("S"), None, F(**border))
    g_fmt = F(bold=True, bg_color=GRAND_GREEN, num_format=NUM0, **border)
    ws.write(grand_r, 0, "Grand Total", F(bold=True, bg_color=GRAND_GREEN, **border))
    refs = "+".join(f"{v[0]}{total_r + 1}" for v in C.MAIN_COLUMNS.values())
    ws.write_formula(grand_r, 1, "=" + refs, g_fmt, round(sum(totals.values()), 2))
    for c in range(2, _col("R") + 1):
        ws.write_blank(grand_r, c, None, F(bold=True, bg_color=GRAND_GREEN))

    ws.freeze_panes(3, 1)
    ws.autofilter(1, _col("R"), last, _col("S"))
    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(1, 2)
    return ws


# ---------------------------------------------------------------------------
# Sheet 2: Env & Tag Stock (can-make cases)
# ---------------------------------------------------------------------------
# (block title, [(column id, kg col, can-make col)], max col, max formula kind)
CAN_MAKE_BLOCKS = [
    ("C250", [("C250_TAG", "B", "C"), ("C250_ENV", "D", "E")], "F"),
    ("MD20", [("MD20_TAG", "H", "I"), ("MD20_ENV", "J", "K")], "L"),
    ("Constanta", [("CONSTANTA_ENV", "N", "O")], "P"),
    ("Perfecta ", [("PENV_ENV", "R", "S")], "T"),
    ("MD20 HS", [("HS_TAG", "V", "W"), ("HS_ENV", "X", "Y")], "Z"),
    ("MAISA", [("MAISA_TAG", "AB", "AC"), ("MAISA_ENV", "AD", "AE")], "AF"),
    ("Perfecta Tag / PK45", [("PTAG_TAG", "AH", "AI")], "AJ"),
]
CAN_MAKE_SPACERS = ["G", "M", "Q", "U", "AA", "AG", "AK"]
CAN_MAKE_WIDTHS = {"A": 45.22, "B": 8.89, "C": 8.22, "D": 9.5, "E": 8.22, "F": 8.78, "G": 1.78,
                   "H": 7.5, "I": 8.0, "J": 9.5, "K": 8.22, "L": 8.78, "M": 1.78, "N": 9.89,
                   "O": 8.22, "P": 8.78, "Q": 1.78, "R": 9.5, "S": 8.22, "T": 8.78, "U": 1.78,
                   "V": 8.0, "W": 8.0, "X": 9.5, "Y": 8.22, "Z": 8.78, "AA": 1.78, "AB": 8.0,
                   "AC": 8.0, "AD": 9.5, "AE": 8.22, "AF": 8.78, "AG": 1.78, "AH": 8.33,
                   "AI": 8.22, "AJ": 8.78, "AK": 1.78, "AL": 10.0, "AM": 21.55}


def _write_can_make(wb, F, report, remarks, conversions):
    ws = wb.add_worksheet("Env & Tag Stock")
    rows = report.rows
    first = 4                                   # zero based -> Excel row 5
    last = first + len(rows) - 1
    total_r = last + 1
    border = {"border": 1}
    main = "'Tag & Env Stock'"

    for letter, w in CAN_MAKE_WIDTHS.items():
        ws.set_column(f"{letter}:{letter}", w)
    ws.set_row(1, 15)
    ws.set_row(2, 43.2)

    ws.write_datetime(0, 0, _dt(report.report_date), F(num_format="mm-dd-yy", bg_color=HEAD_GREEN))
    ws.merge_range(1, 0, 2, 0, "Brand", F(bg_color=HEAD_GREY, align="left", valign="vcenter", **border))
    hdr = F(bg_color=HEAD_GREY, align="center", valign="vcenter", text_wrap=True, **border)
    hdr_b = F(bold=True, bg_color=HEAD_GREY, align="center", valign="vcenter", text_wrap=True, **border)
    for title, cols, max_col in CAN_MAKE_BLOCKS:
        ws.merge_range(1, _col(cols[0][1]), 1, _col(max_col), title,
                       F(bold=True, bg_color=HEAD_GREEN, align="center", valign="vcenter", **border))
        for cid, kg_col, cm_col in cols:
            kind = C.MAIN_COLUMNS[cid][3]
            ws.write(2, _col(kg_col), "Tag Kgs" if kind == "tag" else "Envelope Kgs", hdr)
            ws.write(2, _col(cm_col), "Can make case" if kind == "tag" else "Can Make cases", hdr)
            conv = conversions[C.COLUMN_CONVERSION_KEY[cid]]
            ws.write_number(3, _col(cm_col), conv, F(bg_color=SPACER_YELLOW, bold=True, **border))
        ws.write(2, _col(max_col), "Max Can Make", hdr_b)
    ws.write(2, _col("AL"), "Total Can Make", hdr_b)
    ws.write(2, _col("AM"), "Remarks", hdr)
    ws.write(3, 0, "Conversion (pcs per kg)", F(italic=True, align="left", **border))
    for letter in CAN_MAKE_SPACERS:
        for r in range(1, total_r + 1):
            ws.write_blank(r, _col(letter), None, F(bg_color=SPACER_YELLOW))

    kg_fmt = {"env": F(num_format=NUM0, **border), "tag": F(num_format=NUM2, **border)}
    cm_fmt = F(num_format=NUM0, **border)
    max_fmt = F(num_format=NUM0, bold=True, **border)
    shared_first = {}
    for st, lo, hi in report.shared_spans():
        shared_first[(lo, st.column)] = (st, hi)
    covered = {(i, st.column) for st, lo, hi in report.shared_spans() for i in range(lo + 1, hi + 1)}

    col_values = {}                              # (excel col letter) -> list of values for totals
    for i, row in enumerate(rows):
        r = first + i
        main_r = 4 + i                           # Excel row in the main sheet
        ws.write(r, 0, row.name, F(**border))
        maxes = []
        for title, cols, max_col in CAN_MAKE_BLOCKS:
            cm_cells, cm_vals = [], []
            for cid, kg_col, cm_col in cols:
                kind = C.MAIN_COLUMNS[cid][3]
                main_letter = C.MAIN_COLUMNS[cid][0]
                conv = conversions[C.COLUMN_CONVERSION_KEY[cid]]
                if (i, cid) in covered:
                    continue                     # inside a merged shared-tag block
                if (i, cid) in shared_first:
                    st, hi = shared_first[(i, cid)]
                    kg = st.qty
                    size = report.pack_size("SHARED::" + st.lines[0].item_norm, cid)
                    kg_f = f"={main}!{main_letter}{main_r}"
                    cm_val = kg * conv / size if size else 0
                    cm_f = f"={kg_col}{r + 1}*${cm_col}$4/{size:g}" if size else "=0"
                    span = hi - i
                    if span:
                        ws.merge_range(r, _col(kg_col), r + span, _col(kg_col), "", F(num_format=NUM2,
                                       align="right", valign="vcenter", **border))
                        ws.merge_range(r, _col(cm_col), r + span, _col(cm_col), "", F(num_format=NUM0,
                                       align="right", valign="vcenter", **border))
                    ws.write_formula(r, _col(kg_col), kg_f, F(num_format=NUM2, align="right",
                                     valign="vcenter", **border), round(kg, 2))
                    ws.write_formula(r, _col(cm_col), cm_f, F(num_format=NUM0, align="right",
                                     valign="vcenter", **border), cm_val)
                else:
                    kg = row.qty(cid)
                    size = report.pack_size(row.key, cid)
                    ws.write_formula(r, _col(kg_col), f"={main}!{main_letter}{main_r}", kg_fmt[kind],
                                     round(kg, 2))
                    cm_val = kg * conv / size if size else 0
                    cm_f = f"={kg_col}{r + 1}*${cm_col}$4/{size:g}" if size else "=0"
                    ws.write_formula(r, _col(cm_col), cm_f, cm_fmt, cm_val)
                col_values.setdefault(kg_col, []).append(kg)
                col_values.setdefault(cm_col, []).append(cm_val)
                cm_cells.append(f"{cm_col}{r + 1}")
                cm_vals.append(cm_val)
            mx = max(cm_vals) if cm_vals else 0
            if cm_cells:
                f = f"=MAX({','.join(cm_cells)})" if len(cm_cells) > 1 else f"={cm_cells[0]}"
            else:
                f = "=0"
            ws.write_formula(r, _col(max_col), f, max_fmt, mx)
            col_values.setdefault(max_col, []).append(mx)
            maxes.append((max_col, mx))
        total = sum(v for _, v in maxes)
        ws.write_formula(r, _col("AL"), "=" + "+".join(f"{c}{r + 1}" for c, _ in maxes),
                         max_fmt, total)
        col_values.setdefault("AL", []).append(total)
        ws.write(r, _col("AM"), remarks.get(row.key, row.remark), F(**border))

    tot = F(bold=True, bg_color=TOTAL_BLUE, num_format=NUM0, **border)
    ws.write(total_r, 0, "TOTAL", F(bold=True, bg_color=TOTAL_BLUE, **border))
    letters = [c for _, cols, mx in CAN_MAKE_BLOCKS for x in cols for c in x[1:]] + \
              [mx for _, _, mx in CAN_MAKE_BLOCKS] + ["AL"]
    for letter in letters:
        ws.write_formula(total_r, _col(letter), f"=SUM({letter}{first + 1}:{letter}{last + 1})", tot,
                         sum(col_values.get(letter, [])))
    ws.freeze_panes(4, 1)
    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)
    ws.repeat_rows(1, 3)
    return ws


# ---------------------------------------------------------------------------
# Sheet 3: Ghana & AB & AG Stock
# ---------------------------------------------------------------------------
def _write_export(wb, F, report):
    ws = wb.add_worksheet("Ghana & AB & AG Stock")
    ws.set_column("A:A", 39.66)
    ws.set_column("B:B", 13.55)
    ws.set_column("C:C", 13.11)
    border = {"border": 1}
    title_f = F(bold=True, bg_color=SPACER_YELLOW, align="center", **border)
    env_f = F(num_format=NUM0, font_color="#0D0D0D", **border)
    tag_f = F(num_format=NUM2, **border)
    tot_lbl = F(bold=True, bg_color=GRAND_GREEN, **border)
    tot_f = F(bold=True, bg_color=GRAND_GREEN, num_format=NUM0, **border)

    r = 0
    total_cells = []
    grand = 0.0
    for title, _marker in C.EXPORT_SECTIONS:
        lines = [l for l in report.export_lines if l.section == title]
        has_tag = any(l.column == "tag" for l in lines) or title == "GHANA STOCK"
        ncols = 3 if has_tag else 2
        if ncols == 3:
            ws.merge_range(r, 0, r, 2, title, title_f)
        else:
            ws.merge_range(r, 0, r, 1, title, title_f)
        ws.merge_range(r + 1, 0, r + 2, 0, "Brand", F(bg_color=HEAD_GREY, align="left",
                                                      valign="vcenter", **border))
        ws.write(r + 1, 1, "Constanta", F(bold=True, bg_color=HEAD_GREEN, align="center", **border))
        ws.write(r + 2, 1, "Envelope in Kgs", F(bg_color=HEAD_GREY, **border))
        if ncols == 3:
            ws.write(r + 1, 2, "Constanta Tag", F(bold=True, bg_color=HEAD_GREEN, align="center", **border))
            ws.write(r + 2, 2, "Tag in Kgs", F(bg_color=HEAD_GREY, **border))
        r += 3
        start = r
        envs = sorted([l for l in lines if l.column == "env"], key=lambda l: l.item.upper())
        tags = sorted([l for l in lines if l.column == "tag"], key=lambda l: l.item.upper())
        if not envs and not tags:
            ws.write(r, 0, "(no stock)", F(italic=True, **border))
            ws.write_number(r, 1, 0, env_f)
            if ncols == 3:
                ws.write_number(r, 2, 0, tag_f)
            r += 1
        for l in envs + tags:
            ws.write(r, 0, l.item, F(**border))
            ws.write_number(r, 1, round(l.qty, 2) if l.column == "env" else 0, env_f)
            if ncols == 3:
                ws.write_number(r, 2, round(l.qty, 2) if l.column == "tag" else 0, tag_f)
            r += 1
        env_sum = sum(l.qty for l in envs)
        tag_sum = sum(l.qty for l in tags)
        ws.write(r, 0, "Total" if title == "GHANA STOCK" else "TOTAL", tot_lbl)
        ws.write_formula(r, 1, f"=SUM(B{start + 1}:B{r})", tot_f, round(env_sum, 2))
        total_cells.append(f"B{r + 1}")
        if ncols == 3:
            ws.write_formula(r, 2, f"=SUM(C{start + 1}:C{r})", tot_f, round(tag_sum, 2))
            total_cells.append(f"C{r + 1}")
        grand += env_sum + tag_sum
        r += 2
    r += 1
    gt = F(bold=True, bg_color=_tint(_THEME[4], 0.3999755851924192), num_format=NUM0)
    ws.write(r, 0, "GRAND TOTAL", F(bold=True, bg_color=_tint(_THEME[4], 0.3999755851924192)))
    ws.write_formula(r, 1, "=" + "+".join(total_cells), gt, round(grand, 2))
    ws.set_row(r, 15.6)
    return ws


# ---------------------------------------------------------------------------
# Sheet 4: Validation
# ---------------------------------------------------------------------------
def _write_validation(wb, F, report):
    ws = wb.add_worksheet("Validation")
    ws.set_column("A:A", 28)
    ws.set_column("B:B", 55)
    ws.set_column("C:C", 110)
    head = F(bold=True, bg_color=HEAD_GREEN, border=1)
    ws.write_row(0, 0, ["Check", "Item", "Detail"], head)
    issues = report.issues or [{"Check": "No issues", "Item": "", "Detail": "All inputs matched cleanly."}]
    for i, iss in enumerate(issues, start=1):
        ws.write_row(i, 0, [iss["Check"], iss["Item"], iss["Detail"]], F(border=1, text_wrap=True,
                                                                          valign="top"))
    ws.autofilter(0, 0, len(issues), 2)
    ws.freeze_panes(1, 0)
    return ws


def _dt(d: date | None):
    from datetime import datetime
    d = d or date.today()
    return datetime(d.year, d.month, d.day)


def build_workbook(report, remarks: dict[str, str] | None = None,
                   conversions: dict[str, float] | None = None) -> bytes:
    remarks = remarks or {}
    conv = dict(C.DEFAULT_CONVERSIONS)
    conv.update(conversions or {})
    buf = io.BytesIO()
    wb = xlsxwriter.Workbook(buf, {"in_memory": True})
    F = _Fmt(wb)
    _write_main(wb, F, report, remarks)
    _write_can_make(wb, F, report, remarks, conv)
    _write_export(wb, F, report)
    _write_validation(wb, F, report)
    wb.close()
    return buf.getvalue()


__all__ = ["build_workbook", "xl_col_to_name", "xl_rowcol_to_cell"]
