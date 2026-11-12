"""End-to-end: three uploaded files -> report model -> Excel bytes."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

import pandas as pd

from . import config as C
from .grouping import ExportLine, ReportRow, SharedTag, build_rows, family_key, merge_span
from .readers import (BomData, OrderData, StockData, norm_name, read_exploded_bom,
                      read_order_status, read_previous_remarks, read_stock_summary)
from .remarks import RowOrders, link_orders, pack_sizes, suggest_remark
from .writer import build_workbook


@dataclass
class Report:
    rows: list[ReportRow]
    shared_tags: list[SharedTag]
    export_lines: list[ExportLine]
    orders: dict[str, RowOrders]
    sizes: dict[tuple[str, str], float]
    issues: list[dict]
    report_date: date | None
    sources: dict[str, str] = field(default_factory=dict)
    _spans: list | None = None

    # -- helpers used by the writer and the UI ---------------------------
    def pack_size(self, key: str, col: str):
        return self.sizes.get((key, col))

    def shared_spans(self):
        """[(shared tag, first row index, last row index)] with no overlaps."""
        if self._spans is not None:
            return self._spans
        taken: dict[str, set[int]] = {}
        out = []
        for st in sorted(self.shared_tags, key=lambda s: merge_span(self.rows, s)[0]):
            lo, hi, clean = merge_span(self.rows, st)
            used = taken.setdefault(st.column, set())
            if not clean or any(i in used for i in range(lo, hi + 1)):
                members = [i for i, r in enumerate(self.rows) if r.key in st.member_keys and i not in used]
                if not members:
                    continue
                lo = hi = members[0]
            used |= set(range(lo, hi + 1))
            out.append((st, lo, hi))
        self._spans = out
        return out

    def column_totals(self) -> dict[str, float]:
        tot = {c: sum(r.qty(c) for r in self.rows) for c in C.MAIN_COLUMNS}
        for st, _, _ in self.shared_spans():
            tot[st.column] += st.qty
        return tot

    def preview_frame(self) -> pd.DataFrame:
        """Main sheet as a DataFrame (shared tag shown on its first row)."""
        spans = {(lo, st.column): st.qty for st, lo, _ in self.shared_spans()}
        recs = []
        for i, r in enumerate(self.rows):
            rec = {"Item": r.name}
            for cid, (letter, brand, sub, _) in C.MAIN_COLUMNS.items():
                rec[f"{brand.strip()} {sub}"] = round(r.qty(cid) + spans.get((i, cid), 0.0), 2)
            o = self.orders.get(r.key)
            rec["Remarks"] = r.remark
            rec["Why"] = r.remark_reason
            rec["Pending CFC"] = round(o.pending_cfc, 0) if o else 0
            rec["_key"] = r.key
            recs.append(rec)
        return pd.DataFrame(recs)

    def export_frame(self) -> pd.DataFrame:
        return pd.DataFrame([{"Section": e.section, "Item": e.item,
                              "Envelope Kgs": e.qty if e.column == "env" else 0.0,
                              "Tag Kgs": e.qty if e.column == "tag" else 0.0}
                             for e in self.export_lines])

    def issues_frame(self) -> pd.DataFrame:
        return pd.DataFrame(self.issues, columns=["Check", "Item", "Detail"])

    def to_excel(self, remarks: dict[str, str] | None = None,
                 conversions: dict[str, float] | None = None) -> bytes:
        return build_workbook(self, remarks, conversions)


def build_report(stock: StockData, orders: OrderData, bom: BomData,
                 previous_remarks: dict[str, str] | None = None) -> Report:
    grouping = build_rows(stock.items, bom.lines)
    orders_by_row, order_issues = link_orders(grouping, orders.lines, bom.lines)
    prev_by_family = {family_key(k): v for k, v in (previous_remarks or {}).items()}
    active = (C.REMARK_RUNNING, C.REMARK_ORDER)
    for row in grouping.rows:
        row.remark, row.remark_reason = suggest_remark(row, orders_by_row[row.key])
        if previous_remarks and row.remark in active:
            prev = previous_remarks.get(norm_name(row.name)) or prev_by_family.get(row.key)
            if prev in active and prev != row.remark:
                row.remark = prev
                row.remark_reason += f" (kept '{prev}' from the previous report)"
    sizes = pack_sizes(grouping, bom.lines, orders_by_row)

    issues: list[dict] = []
    for src, warns in (("Stock Summary", stock.warnings), ("Order Status", orders.warnings),
                       ("Exploded BOM", bom.warnings)):
        for w in warns:
            issues.append({"Check": "Input warning", "Item": src, "Detail": w})
    issues += grouping.issues
    issues += order_issues
    for row in grouping.rows:
        for col in row.lines:
            if (row.key, col) not in sizes:
                issues.append({"Check": "Pack size not in BOM", "Item": row.name,
                               "Detail": f"No BOM line gives pieces per CFC for "
                                         f"{', '.join(l.item for l in row.lines[col])} - "
                                         "can-make cases left at 0 on the 'Env & Tag Stock' sheet."})
    # Reconciliation: every kg of mapped envelope/tag stock must land in the report
    mapped = stock.items[stock.items["machine_line"].map(C.machine_line_column).notna()]["qty"].sum()
    placed = sum(r.qty(c) for r in grouping.rows for c in C.MAIN_COLUMNS) \
        + sum(st.qty for st in grouping.shared_tags) + sum(e.qty for e in grouping.export_lines)
    if abs(mapped - placed) > 0.01:
        issues.insert(0, {"Check": "Reconciliation", "Item": "Stock Summary",
                          "Detail": f"Envelope/tag stock in the file is {mapped:,.2f} kg but the report "
                                    f"shows {placed:,.2f} kg."})
    if stock.report_date is None:
        issues.append({"Check": "Input warning", "Item": "Stock Summary",
                       "Detail": "Report date not found in the file - today's date is used."})

    return Report(rows=grouping.rows, shared_tags=grouping.shared_tags,
                  export_lines=grouping.export_lines, orders=orders_by_row, sizes=sizes,
                  issues=issues, report_date=stock.report_date,
                  sources={"Stock Summary sheet": stock.sheet_used,
                           "Order Status sheet": orders.sheet_used,
                           "Exploded BOM sheet": bom.sheet_used,
                           "BOM run date": bom.run_date or "-",
                           "Envelope/tag stock in Stock Summary (kg)": f"{mapped:,.2f}",
                           "Envelope/tag stock placed in report (kg)": f"{placed:,.2f}"})


def run(stock_bytes: bytes, order_bytes: bytes, bom_bytes: bytes,
        stock_name: str | None = None, previous_bytes: bytes | None = None) -> Report:
    previous = read_previous_remarks(previous_bytes) if previous_bytes else None
    return build_report(read_stock_summary(stock_bytes, stock_name),
                        read_order_status(order_bytes),
                        read_exploded_bom(bom_bytes), previous)
