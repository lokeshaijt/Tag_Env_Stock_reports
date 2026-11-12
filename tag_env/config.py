"""Business configuration for the Tag & Env Stock report.

Everything a planner may need to tweak lives here: machine-line -> report
column mapping, conversion factors (pcs per kg), export sections and the
name tokens stripped to build a "family" key.
"""
from __future__ import annotations

import re

# ---------------------------------------------------------------------------
# Main sheet ("Tag & Env Stock") columns
# ---------------------------------------------------------------------------
# Logical column id -> (excel column letter, brand header, sub header, kind)
# kind: "env" (envelope, kgs, 0 decimals) or "tag" (tag, kgs, 2 decimals)
MAIN_COLUMNS = {
    "CONSTANTA_ENV": ("B", "Constanta", "Envelope in Kgs", "env"),
    "C250_ENV": ("D", "C250", "Envelope in Kgs", "env"),
    "C250_TAG": ("E", "C250", "Tag Kgs", "tag"),
    "MD20_ENV": ("G", "MD20", "Envelope in Kgs", "env"),
    "MD20_TAG": ("H", "MD20", "Tag Kgs", "tag"),
    "PENV_ENV": ("J", "Perfecta ", "Envelope in Kgs", "env"),
    "HS_ENV": ("L", "MD20 HS", "Envelope in Kgs", "env"),
    "HS_TAG": ("M", "MD20 HS", "Tag Kgs", "tag"),
    "MAISA_ENV": ("O", "Maisa", "Envelope in Kgs", "env"),
    "MAISA_TAG": ("P", "Maisa", "Tag Kgs", "tag"),
    "PTAG_TAG": ("R", "Perfecta/PK45 Tag", "Tag in Kgs", "tag"),
}
ENV_COLUMNS = [k for k, v in MAIN_COLUMNS.items() if v[3] == "env"]
TAG_COLUMNS = [k for k, v in MAIN_COLUMNS.items() if v[3] == "tag"]

# A tag only joins an envelope family row when that family has an envelope on
# the paired machine line (e.g. a Constanta tag goes with a C250 envelope).
TAG_PAIRED_ENV = {
    "C250_TAG": "C250_ENV",
    "MD20_TAG": "MD20_ENV",
    "HS_TAG": "HS_ENV",
    "MAISA_TAG": "MAISA_ENV",
    "PTAG_TAG": "PENV_ENV",
}

# Which envelope names the row (column A) when a family has several lines.
ROW_NAME_PRIORITY = ["PENV_ENV", "C250_ENV", "MD20_ENV", "HS_ENV", "MAISA_ENV", "CONSTANTA_ENV"]

# Order of tag-only rows at the bottom of the sheet (as in the manual report).
TAG_ONLY_ORDER = ["PTAG_TAG", "MAISA_TAG", "HS_TAG", "MD20_TAG", "C250_TAG"]


def _mlkey(s: str) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())


# Machine Line Name (as printed in the Stock Summary) -> logical column.
_MACHINE_LINE_MAP_RAW = {
    "Constanta Envelope": "CONSTANTA_ENV",
    "CONSTANTA-250": "C250_ENV",
    "Constanta 250": "C250_ENV",
    "Constanta Tag": "C250_TAG",
    "Constanta Tag Big": "C250_TAG",
    "MD-20 Crimp Envelope": "MD20_ENV",
    "MD-20 Envelope": "MD20_ENV",
    "MD-20 Tag": "MD20_TAG",
    "Perfecta Envelope (PENV)": "PENV_ENV",
    "Perfecta Envelope": "PENV_ENV",
    "MD HS Envelope": "HS_ENV",
    "MD HS Envelope Paper": "HS_ENV",
    "MD HS Envelope Foil": "HS_ENV",
    "MD-20 HS Envelope": "HS_ENV",
    "MD HS Tag": "HS_TAG",
    "MD-20 HS Tag": "HS_TAG",
    "Maisa Envelope": "MAISA_ENV",
    "Maisa Tag": "MAISA_TAG",
    "PERFACTA TAG": "PTAG_TAG",
    "Perfecta Tag": "PTAG_TAG",
    "PK-45": "PTAG_TAG",
    "PK45": "PTAG_TAG",
}
MACHINE_LINE_MAP = {_mlkey(k): v for k, v in _MACHINE_LINE_MAP_RAW.items()}


def machine_line_column(machine_line) -> str | None:
    if machine_line is None:
        return None
    return MACHINE_LINE_MAP.get(_mlkey(machine_line))


# ---------------------------------------------------------------------------
# Conversion factors (pcs per kg) - from the "Conversion Round Off" chart.
# Used by the "Env & Tag Stock" (can-make cases) sheet. Editable in the UI.
# ---------------------------------------------------------------------------
DEFAULT_CONVERSIONS = {
    "Constanta Envelope": 1525,
    "MD-20 Crimp Envelope": 1700,
    "MD HS Envelope Paper": 1425,
    "MD HS Envelope Foil": 1190,
    "Perfecta Envelope (PENV)": 1520,
    "Maisa Envelope": 1410,
    "CONSTANTA-250": 1740,
    "CONSTANTA FAMILY SIZE": 6320,
    "Constanta Tag Big": 6040,
    "MD-20 Tag": 12600,
    "Maisa Tag": 9080,
    "PERFACTA TAG & PK-45 & Constanta Tag": 10780,
}

# Which conversion factor each logical column uses.
COLUMN_CONVERSION_KEY = {
    "C250_TAG": "PERFACTA TAG & PK-45 & Constanta Tag",
    "C250_ENV": "CONSTANTA-250",
    "MD20_TAG": "MD-20 Tag",
    "MD20_ENV": "MD-20 Crimp Envelope",
    "CONSTANTA_ENV": "Constanta Envelope",
    "PENV_ENV": "Perfecta Envelope (PENV)",
    "HS_TAG": "MD-20 Tag",
    "HS_ENV": "MD HS Envelope Paper",
    "MAISA_TAG": "Maisa Tag",
    "MAISA_ENV": "Maisa Envelope",
    "PTAG_TAG": "PERFACTA TAG & PK-45 & Constanta Tag",
}

# ---------------------------------------------------------------------------
# Export stock ("Ghana & AB & AG Stock" sheet)
# ---------------------------------------------------------------------------
# (section title, marker in item name)
EXPORT_SECTIONS = [
    ("GHANA STOCK", "(GHANA)"),
    ("ABHIDJAN STOCK", "(ABIDJAN)"),
    ("ANGOLA STOCK", "(ANGOLA)"),
]
# Items without a marker that the planners still book under an export section.
EXPORT_EXTRA_ITEMS = {
    "ENV INDUS DC TBGS (70GSM)": "GHANA STOCK",
}

# ---------------------------------------------------------------------------
# Family key: strip machine-line / construction words from an item name so the
# envelopes and tags of one blend/flavour land on the same report row.
# ---------------------------------------------------------------------------
STRIP_PARENS = re.compile(
    r"\((?:\s*70\s*GSM\s*|\s*WO\s*GLUE\s*|\s*V\d+\s*|\s*3\s*MIN\s*)\)", re.IGNORECASE
)
STRIP_TOKENS = {
    "DC", "SF", "C250", "ENV", "TBGS", "PENV", "HS", "PTAG", "PK45", "TAG",
    "FT", "TBGS-FT",
}

REMARK_RUNNING = "Running"
REMARK_ORDER = "Order Received"
REMARK_MULTILINE = "Multiline Stock"
REMARK_LEFTOVER = "Leftover Stock"
REMARK_CHOICES = [REMARK_RUNNING, REMARK_ORDER, REMARK_MULTILINE, REMARK_LEFTOVER,
                  "Loose Item", "Yet to Plan", "No Order"]
