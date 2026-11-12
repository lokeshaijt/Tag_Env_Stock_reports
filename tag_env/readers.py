"""Readers for the three input reports.

Each reader takes the raw bytes of an uploaded file, works out the real file
format (ERP exports often save .xlsx data with a .xls extension, or HTML with a
.xls extension), finds the header row by its column names rather than a fixed
position, and returns a tidy DataFrame plus a list of warnings.
"""
from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import date, datetime

import pandas as pd


class InputError(ValueError):
    """Raised when an uploaded file is not the report we expected."""


def norm_name(value) -> str:
    """Upper-case, single-spaced item name used as a join key."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().upper()


def _norm_header(value) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).lower()) if value is not None else ""


def _engine_for(data: bytes) -> str:
    head = data[:8]
    if head.startswith(b"PK"):
        return "openpyxl"
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        return "xlrd"
    sniff = data[:2048].lower()
    if b"<html" in sniff or b"<table" in sniff:
        return "html"
    raise InputError("The file is not a readable Excel workbook (.xlsx / .xls).")


class _Workbook:
    """Lazy access to the sheets of an uploaded workbook."""

    def __init__(self, data: bytes):
        self.engine = _engine_for(data)
        self._data = data
        self._tables = None
        self._xl = None
        if self.engine == "html":
            self._tables = {f"Table{i + 1}": t for i, t in
                            enumerate(pd.read_html(io.BytesIO(data), header=None))}
        else:
            try:
                self._xl = pd.ExcelFile(io.BytesIO(data), engine=self.engine)
            except ImportError as exc:  # xlrd missing
                raise InputError(f"Cannot read this file format: {exc}") from exc

    @property
    def sheet_names(self) -> list[str]:
        return list(self._tables) if self._tables is not None else list(self._xl.sheet_names)

    def read(self, name: str, nrows: int | None = None) -> pd.DataFrame:
        if self._tables is not None:
            t = self._tables[name]
            return t.head(nrows) if nrows else t
        return self._xl.parse(name, header=None, nrows=nrows)

    def find_sheets(self, required: list[str]):
        """[(sheet name, header row index)] for sheets containing all required headers."""
        out = []
        for name in self.sheet_names:
            hdr = _find_header_row(self.read(name, nrows=40), required)
            if hdr is not None:
                out.append((name, hdr))
        return out


def _find_header_row(raw: pd.DataFrame, required: list[str], max_scan: int = 40):
    """Index of the first row that contains every required header."""
    req = [_norm_header(r) for r in required]
    for i in range(min(max_scan, len(raw))):
        cells = [_norm_header(v) for v in raw.iloc[i].tolist()]
        if all(r in cells for r in req):
            return i
    return None


def _col_index(header_cells: list[str], *candidates: str, occurrence: int = 0):
    """Column position of the first matching candidate header."""
    for cand in candidates:
        c = _norm_header(cand)
        hits = [i for i, h in enumerate(header_cells) if h == c]
        if len(hits) > occurrence:
            return hits[occurrence]
    return None


def _to_number(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


@dataclass
class StockData:
    items: pd.DataFrame            # item, item_norm, machine_line, qty
    report_date: date | None
    sheet_used: str
    warnings: list[str] = field(default_factory=list)


@dataclass
class OrderData:
    lines: pd.DataFrame            # fg, fg_norm, pending, sip, ready, pending_cfc, ship_date
    sheet_used: str
    warnings: list[str] = field(default_factory=list)


@dataclass
class BomData:
    lines: pd.DataFrame            # fg, fg_norm, item, item_norm, item_type, qty, uom, fg_uom
    sheet_used: str
    run_date: str | None
    warnings: list[str] = field(default_factory=list)


# ---------------------------------------------------------------------------
# Stock Summary Report
# ---------------------------------------------------------------------------
_DATE_RE = re.compile(r"(\d{4})-(\d{2})-(\d{2})")


def _parse_report_date(raw: pd.DataFrame, filename: str | None) -> date | None:
    for i in range(min(5, len(raw))):
        for v in raw.iloc[i].tolist():
            if isinstance(v, (datetime, pd.Timestamp)):
                return v.date()
            m = _DATE_RE.search(str(v)) if v is not None else None
            if m:
                return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    if filename:
        m = _DATE_RE.search(filename)
        if m:
            return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return None


def read_stock_summary(data: bytes, filename: str | None = None) -> StockData:
    wb = _Workbook(data)
    candidates = wb.find_sheets(["Item Name", "Qty"])
    if not candidates:
        raise InputError("Stock Summary: could not find a sheet with 'Item Name' and 'Qty' columns. "
                         "Please upload the Stock Summary Report exported from the ERP.")
    # Prefer the original export over a user-made copy such as "Sheet (2)".
    candidates.sort(key=lambda c: bool(re.search(r"\(\d+\)\s*$", c[0])))
    name, hdr = candidates[0]
    raw = wb.read(name)
    warnings: list[str] = []

    header = [_norm_header(v) for v in raw.iloc[hdr].tolist()]
    i_item = _col_index(header, "Item Name", "Item")
    i_ml = _col_index(header, "Machine Line Name", "Machine Line", "MachineLine")
    i_qty = _col_index(header, "Qty", "Quantity", "Closing Qty")
    if i_ml is None:
        raise InputError("Stock Summary: the 'Machine Line Name' column is missing. It is needed to "
                         "place each item in the right column of the report.")

    body = raw.iloc[hdr + 1:]
    df = pd.DataFrame({
        "item": body.iloc[:, i_item].astype(object),
        "machine_line": body.iloc[:, i_ml].astype(object),
        "qty": _to_number(body.iloc[:, i_qty]),
    })
    df["item"] = df["item"].map(lambda v: re.sub(r"\s+", " ", str(v)).strip() if pd.notna(v) else "")
    df = df[df["item"] != ""]
    df = df[~df["item"].str.upper().isin(["TOTAL", "GRAND TOTAL"])]
    df["machine_line"] = df["machine_line"].map(lambda v: str(v).strip() if pd.notna(v) else "")
    df["item_norm"] = df["item"].map(norm_name)

    dups = df[df.duplicated("item_norm", keep=False)]
    if not dups.empty:
        names = sorted(dups["item"].unique())
        warnings.append(f"Stock Summary: {len(names)} item(s) appear more than once; their "
                        f"quantities were added together: {', '.join(names[:10])}"
                        + (" ..." if len(names) > 10 else ""))
        df = (df.groupby("item_norm", as_index=False)
                .agg(item=("item", "first"), machine_line=("machine_line", "first"), qty=("qty", "sum")))

    neg = df[df["qty"] < 0]
    if not neg.empty:
        warnings.append(f"Stock Summary: {len(neg)} item(s) have a negative quantity: "
                        + ", ".join(neg["item"].head(10)))

    return StockData(items=df.reset_index(drop=True), report_date=_parse_report_date(raw, filename),
                     sheet_used=name, warnings=warnings)


# ---------------------------------------------------------------------------
# Order Status Report
# ---------------------------------------------------------------------------
def read_order_status(data: bytes) -> OrderData:
    wb = _Workbook(data)
    found = wb.find_sheets(["Prod Name", "Pending Ord Qty"])
    if not found:
        raise InputError("Order Status: could not find a sheet with 'Prod Name' and 'Pending Ord Qty' "
                         "columns. Please upload the Order Status Report.")
    # The by-buyer and by-shipment-date sheets list the same orders; use one only.
    found.sort(key=lambda f: (0 if "buyer" in f[0].lower() else 1))
    name, hdr = found[0]
    raw = wb.read(name)
    warnings: list[str] = []

    header = [_norm_header(v) for v in raw.iloc[hdr].tolist()]
    idx = {
        "fg": _col_index(header, "Prod Name"),
        "pending": _col_index(header, "Pending Ord Qty"),
        "sip": _col_index(header, "Stock In Process"),
        "ready": _col_index(header, "Ready Stock"),
        "pending_cfc": _col_index(header, "Pending Prod In Cfc"),
        "ship_date": _col_index(header, "Plan To Ship", "Buyer Requested Shipment Date"),
        "status": _col_index(header, "Order Status"),
        "prod_id": _col_index(header, "Prod Id"),
    }
    body = raw.iloc[hdr + 1:]
    df = pd.DataFrame({"fg": body.iloc[:, idx["fg"]]})
    for key in ("pending", "sip", "ready", "pending_cfc"):
        df[key] = _to_number(body.iloc[:, idx[key]]) if idx[key] is not None else 0.0
    if idx["pending_cfc"] is None:
        df["pending_cfc"] = (df["pending"] - df["sip"] - df["ready"]).clip(lower=0)
        warnings.append("Order Status: 'Pending Prod In Cfc' column not found - calculated as "
                        "Pending Ord Qty - Stock In Process - Ready Stock.")
    df["ship_date"] = body.iloc[:, idx["ship_date"]] if idx["ship_date"] is not None else None
    df["fg"] = df["fg"].map(lambda v: re.sub(r"\s+", " ", str(v)).strip() if pd.notna(v) else "")
    df = df[(df["fg"] != "") & (df["fg"].str.upper() != "PROD NAME")]
    df["fg_norm"] = df["fg"].map(norm_name)
    if df.empty:
        warnings.append("Order Status: no order lines found - all items will be marked as leftover.")
    return OrderData(lines=df.reset_index(drop=True), sheet_used=name, warnings=warnings)


# ---------------------------------------------------------------------------
# Exploded BOM Report
# ---------------------------------------------------------------------------
def read_exploded_bom(data: bytes) -> BomData:
    wb = _Workbook(data)
    chosen = wb.find_sheets(["FG Name", "Item Name", "Qty"])
    if not chosen:
        raise InputError("Exploded BOM: could not find a sheet with 'FG Name', 'Item Name' and 'Qty' "
                         "columns. Please upload the Exploded BOM Report.")
    name, hdr = chosen[0]
    raw = wb.read(name)
    warnings: list[str] = []
    header = [_norm_header(v) for v in raw.iloc[hdr].tolist()]
    i_fg = _col_index(header, "FG Name")
    i_item = _col_index(header, "Item Name")
    i_type = _col_index(header, "ITEM TYPE", "Item Type")
    i_qty = _col_index(header, "Qty")
    i_uom = _col_index(header, "Uom", "UOM")
    i_fguom = _col_index(header, "FG Uom", "FG UOM")
    body = raw.iloc[hdr + 1:]
    df = pd.DataFrame({
        "fg": body.iloc[:, i_fg],
        "item": body.iloc[:, i_item],
        "item_type": body.iloc[:, i_type] if i_type is not None else "",
        "qty": _to_number(body.iloc[:, i_qty]),
        "uom": body.iloc[:, i_uom] if i_uom is not None else "",
        "fg_uom": body.iloc[:, i_fguom] if i_fguom is not None else "",
    })
    df = df[df["fg"].notna() & df["item"].notna()]
    for c in ("fg", "item"):
        df[c] = df[c].map(lambda v: re.sub(r"\s+", " ", str(v)).strip())
    for c in ("item_type", "uom", "fg_uom"):
        df[c] = df[c].map(lambda v: str(v).strip().upper() if pd.notna(v) else "")
    df["fg_norm"] = df["fg"].map(norm_name)
    df["item_norm"] = df["item"].map(norm_name)
    before = len(df)
    df = df.drop_duplicates(["fg_norm", "item_norm"])
    if len(df) < before:
        warnings.append(f"Exploded BOM: {before - len(df)} duplicate FG/item line(s) ignored.")

    run_date = None
    for v in raw.iloc[: max(hdr, 1)].values.ravel().tolist():
        if v is not None and "run date" in str(v).lower():
            run_date = str(v).split("-", 1)[-1].strip()
    return BomData(lines=df.reset_index(drop=True), sheet_used=name, run_date=run_date,
                   warnings=warnings)


# ---------------------------------------------------------------------------
# File type detection (so users can drop the files in any order)
# ---------------------------------------------------------------------------
KIND_STOCK, KIND_ORDER, KIND_BOM, KIND_PREVIOUS = "stock", "order", "bom", "previous"
KIND_LABELS = {
    KIND_STOCK: "Stock Summary Report",
    KIND_ORDER: "Order Status Report",
    KIND_BOM: "Exploded BOM Report",
    KIND_PREVIOUS: "Previous Tag & Env Stock report (optional)",
}


def identify(data: bytes) -> str | None:
    """Which of the input reports this file is, judged by its column headers."""
    try:
        wb = _Workbook(data)
    except Exception:
        return None
    if wb.find_sheets(["Prod Name", "Pending Ord Qty"]):
        return KIND_ORDER
    if wb.find_sheets(["FG Name", "Item Name", "Qty"]):
        return KIND_BOM
    if wb.find_sheets(["Item Name", "Machine Line Name"]) or wb.find_sheets(["Item Name", "Qty"]):
        return KIND_STOCK
    if wb.find_sheets(["Brand", "Remarks"]):
        return KIND_PREVIOUS
    return None


def read_previous_remarks(data: bytes) -> dict[str, str]:
    """{normalised row name: remark} from an earlier Tag & Env Stock report."""
    wb = _Workbook(data)
    found = wb.find_sheets(["Brand", "Remarks"])
    # prefer the visible main sheet layout (remarks in the last header column)
    out: dict[str, str] = {}
    for name, hdr in found:
        raw = wb.read(name)
        header = [_norm_header(v) for v in raw.iloc[hdr].tolist()]
        i_rem = _col_index(header, "Remarks")
        i_name = _col_index(header, "Brand")
        for _, rec in raw.iloc[hdr + 1:].iterrows():
            n, r = rec.iloc[i_name], rec.iloc[i_rem]
            if pd.notna(n) and pd.notna(r) and str(r).strip():
                out.setdefault(norm_name(n), str(r).strip())
    return out
