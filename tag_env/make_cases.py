"""Make Cases calculation for the "Env & Tag Stock" sheet.

    Make Cases = (Available KG x Conversion Factor) / Exploded BOM value

* Available KG       - the Tag or Envelope kg of the row on that machine line.
* Conversion Factor  - pcs per kg of that machine line (config.DEFAULT_CONVERSIONS,
                       from the "Conversion Round Off" chart), one per machine line.
* Exploded BOM value - pieces of that envelope / tag per case (CFC), read from the
                       uploaded Exploded BOM for the finished good that uses it
                       (the FG with an open order first, otherwise the most common
                       pack size). Never hard-coded.

When a factor or BOM value is missing the cell is left empty (and listed in the
data checks) instead of showing a wrong number.
"""
from __future__ import annotations

from dataclasses import dataclass

from . import config as C

# Sheet layout of "Env & Tag Stock" (as in the reference report):
# (block title, [(column id, kg column, make-cases column)], max column)
# The conversion factor of each machine line sits in row 4 above its make-cases column.
BLOCKS = [
    ("C250", [("C250_TAG", "B", "C"), ("C250_ENV", "D", "E")], "F"),
    ("MD20", [("MD20_TAG", "H", "I"), ("MD20_ENV", "J", "K")], "L"),
    ("Constanta", [("CONSTANTA_ENV", "N", "O")], "P"),
    ("Perfecta ", [("PENV_ENV", "R", "S")], "T"),
    ("MD20 HS", [("HS_TAG", "V", "W"), ("HS_ENV", "X", "Y")], "Z"),
    ("MAISA", [("MAISA_TAG", "AB", "AC"), ("MAISA_ENV", "AD", "AE")], "AF"),
    ("Perfecta Tag / PK45", [("PTAG_TAG", "AH", "AI")], "AJ"),
]
# The reference report's "Total" column adds these Max Can Make columns.
TOTAL_OF = ["F", "L", "P", "T"]


@dataclass
class Cell:
    kg: float
    factor: float | None
    bom_value: float | None
    cases: float | None          # None = cannot be calculated (missing factor / BOM value)
    merged_to: int | None = None  # last row index of a shared-tag block starting here
    covered: bool = False         # inside a shared-tag block (no own value)


def conversion_for(cid: str, conversions: dict[str, float]) -> float | None:
    v = conversions.get(C.COLUMN_CONVERSION_KEY.get(cid, ""))
    return float(v) if v else None


def calculate(report, conversions: dict[str, float]):
    """Return (cells, issues): cells[(row index, column id)] -> Cell."""
    cells: dict[tuple[int, str], Cell] = {}
    issues: list[dict] = []
    first_of = {(lo, st.column): (st, hi) for st, lo, hi in report.shared_spans()}
    covered = {(i, st.column) for st, lo, hi in report.shared_spans() for i in range(lo + 1, hi + 1)}

    for i, row in enumerate(report.rows):
        row_sizes = [report.pack_size(row.key, c) for c in row.lines]
        fallback = next((s for s in row_sizes if s), None)
        for _, cols, _ in BLOCKS:
            for cid, _, _ in cols:
                factor = conversion_for(cid, conversions)
                if (i, cid) in covered:
                    cells[(i, cid)] = Cell(0.0, factor, None, None, covered=True)
                    continue
                merged_to = None
                if (i, cid) in first_of:
                    st, hi = first_of[(i, cid)]
                    kg = st.qty
                    size = report.pack_size("SHARED::" + st.lines[0].item_norm, cid)
                    label = ", ".join(l.item for l in st.lines)
                    merged_to = hi
                else:
                    kg = row.qty(cid)
                    size = report.pack_size(row.key, cid)
                    label = ", ".join(l.item for l in row.lines.get(cid, [])) or row.name
                if kg == 0:
                    size = size or fallback
                    cases = 0.0
                elif not factor:
                    cases = None
                    issues.append({"Check": "Missing conversion factor", "Item": label,
                                   "Detail": f"No conversion factor for {C.MAIN_COLUMNS[cid][1].strip()} "
                                             f"{C.MAIN_COLUMNS[cid][2]} - make cases left empty."})
                elif not size:
                    cases = None
                    issues.append({"Check": "Missing Exploded BOM value", "Item": label,
                                   "Detail": f"{kg:,.2f} kg in stock but no BOM line gives pieces per case "
                                             "for this item - make cases left empty."})
                else:
                    cases = kg * factor / size
                cells[(i, cid)] = Cell(kg, factor, size, cases, merged_to=merged_to)
    return cells, issues


def block_max(cells, i: int, cols) -> float:
    vals = [cells[(i, cid)].cases for cid, _, _ in cols
            if not cells[(i, cid)].covered and cells[(i, cid)].cases is not None]
    return max(vals) if vals else 0.0
