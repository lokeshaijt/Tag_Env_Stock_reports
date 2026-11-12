"""Remarks (Running / Order Received / Multiline Stock / Leftover Stock) and
pack sizes, worked out from the Order Status report and the Exploded BOM.

A row is linked to the finished goods (FGs) whose BOM uses any of its envelope
or tag items. Order FGs missing from the BOM are matched by name instead
("GV CLOVE 24/25 DC SF ENV TBGS" -> family "GV CLOVE").

Rules, applied in order:
  1. Running         - an open order still needs production AND production of it
                       has already started (Stock In Process or Ready Stock > 0)
  2. Order Received  - an open order still needs production, nothing made yet
  3. Multiline Stock - no production pending, stock sits on 2+ envelope lines
  4. Leftover Stock  - no production pending
"""
from __future__ import annotations

import re
from collections import defaultdict
from dataclasses import dataclass

import pandas as pd

from . import config as C
from .grouping import Grouping, ReportRow, family_key


def fg_family_key(fg_name: str) -> str:
    n = re.sub(r"\b\d+\s*/\s*\d+\b", " ", str(fg_name).upper())
    n = re.sub(r"\(\s*(AF|V\d+)\s*\)", " ", n)
    n = re.sub(r"\b\d+\b", " ", n)
    n = re.sub(r"\bWITH\b.*$", " ", n)
    return family_key(n)


@dataclass
class RowOrders:
    fgs: list[str]
    pending: float
    sip: float
    ready: float
    pending_cfc: float


def link_orders(grouping: Grouping, orders: pd.DataFrame, bom: pd.DataFrame):
    """Attach order totals to every row. Returns (row key -> RowOrders, issues)."""
    issues: list[dict] = []
    agg = orders.groupby("fg_norm")[["pending", "sip", "ready", "pending_cfc"]].sum()
    names = orders.groupby("fg_norm")["fg"].first()
    bom_fgs = set(bom["fg_norm"])
    item_to_fgs = bom.groupby("item_norm")["fg_norm"].apply(set).to_dict()

    missing = [f for f in agg.index if f not in bom_fgs]
    by_family: dict[str, set[str]] = defaultdict(set)
    for f in missing:
        by_family[fg_family_key(f)].add(f)

    used_fallback: set[str] = set()
    result: dict[str, RowOrders] = {}
    for row in grouping.rows:
        fgs: set[str] = set()
        for ln in row.all_lines():
            fgs |= item_to_fgs.get(ln.item_norm, set())
        fb = by_family.get(row.key, set())
        if fb:
            fgs |= fb
            used_fallback |= fb
        on_order = sorted(f for f in fgs if f in agg.index)
        tot = agg.loc[on_order].sum() if on_order else pd.Series(0.0, index=agg.columns)
        result[row.key] = RowOrders([names[f] for f in on_order], float(tot["pending"]),
                                    float(tot["sip"]), float(tot["ready"]), float(tot["pending_cfc"]))

    for f in missing:
        if f not in used_fallback and agg.loc[f, "pending_cfc"] > 0 and "TBGS" in f:
            issues.append({"Check": "Order FG not in BOM", "Item": names[f],
                           "Detail": f"Pending production {agg.loc[f, 'pending_cfc']:,.0f} CFC, but this FG "
                                     "is not in the Exploded BOM and no stock row matches its name - "
                                     "its envelope/tag could not be identified."})
        elif f in used_fallback:
            issues.append({"Check": "Order FG matched by name", "Item": names[f],
                           "Detail": "FG is not in the Exploded BOM; linked to the stock row with the "
                                     f"same blend name ('{fg_family_key(f)}'). Please verify."})
    return result, issues


def suggest_remark(row: ReportRow, o: RowOrders) -> tuple[str, str]:
    if o.pending_cfc > 0 and (o.sip > 0 or o.ready > 0):
        return C.REMARK_RUNNING, (f"Pending {o.pending_cfc:,.0f} CFC to produce; "
                                  f"{o.sip + o.ready:,.0f} CFC already in process/ready")
    if o.pending_cfc > 0:
        return C.REMARK_ORDER, f"Pending {o.pending_cfc:,.0f} CFC to produce, not started"
    if len(row.env_columns) >= 2:
        lines = ", ".join(C.MAIN_COLUMNS[c][1].strip() for c in row.env_columns)
        return C.REMARK_MULTILINE, f"No pending production; envelope stock on {lines}"
    if o.pending > 0:
        return C.REMARK_LEFTOVER, "Open orders are fully covered by ready stock"
    return C.REMARK_LEFTOVER, "No open order uses this item"


def pack_sizes(grouping: Grouping, bom: pd.DataFrame, orders_by_row: dict[str, RowOrders]):
    """Pieces of each envelope/tag per CFC, per (row key, column).

    Taken from the BOM line of the FG with the biggest open order; when none of
    the FGs is on order, the most common CFC pack size is used.
    """
    pcs = bom[(bom["uom"].isin(["PCS", ""])) & (bom["qty"] > 0)]
    cfc = pcs[pcs["fg_uom"].isin(["CFC", "CASE", ""])]
    by_item = {k: v for k, v in cfc.groupby("item_norm")}
    sizes: dict[tuple[str, str], float] = {}

    def size_for(item_norms: list[str], ordered: set[str]):
        frames = [by_item[i] for i in item_norms if i in by_item]
        if not frames:
            return None
        df = pd.concat(frames)
        hit = df[df["fg"].isin(ordered)]
        if not hit.empty:
            return float(hit["qty"].mode().iloc[0])
        return float(df["qty"].mode().iloc[0])

    for row in grouping.rows:
        ordered = set(orders_by_row.get(row.key, RowOrders([], 0, 0, 0, 0)).fgs)
        for col, ls in row.lines.items():
            s = size_for([l.item_norm for l in ls], ordered)
            if s:
                sizes[(row.key, col)] = s
    for st in grouping.shared_tags:
        ordered = set()
        for k in st.member_keys:
            ordered |= set(orders_by_row.get(k, RowOrders([], 0, 0, 0, 0)).fgs)
        s = size_for([l.item_norm for l in st.lines], ordered)
        if s:
            sizes[("SHARED::" + st.lines[0].item_norm, st.column)] = s
    return sizes
