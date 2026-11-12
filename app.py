"""Tag & Env Stock Report - Streamlit app.

Upload the Stock Summary, Order Status and Exploded BOM reports; the app builds
the Tag & Env Stock report in the same layout as the manual Excel and offers it
for download.
"""
from __future__ import annotations

import hashlib
from datetime import date

import pandas as pd
import streamlit as st

from tag_env import config as C
from tag_env.pipeline import build_report
from tag_env.readers import (KIND_BOM, KIND_LABELS, KIND_ORDER, KIND_STOCK, InputError,
                             identify, read_exploded_bom, read_order_status, read_stock_summary)

st.set_page_config(page_title="Tag & Env Stock Report", page_icon="📦", layout="wide")


# ---------------------------------------------------------------------------
# Cached steps (keyed on file content, so re-runs are instant)
# ---------------------------------------------------------------------------
@st.cache_data(show_spinner=False, max_entries=20)
def _identify(data: bytes) -> str | None:
    return identify(data)


@st.cache_data(show_spinner=False, max_entries=6)
def _stock(data: bytes, name: str):
    return read_stock_summary(data, name)


@st.cache_data(show_spinner=False, max_entries=6)
def _orders(data: bytes):
    return read_order_status(data)


@st.cache_data(show_spinner=False, max_entries=3)
def _bom(data: bytes):
    return read_exploded_bom(data)


def _digest(*parts: bytes | None) -> str:
    h = hashlib.sha1()
    for p in parts:
        h.update(p or b"-")
    return h.hexdigest()


# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.title("📦 Tag & Env Stock Report")
st.write(
    "Upload the **Stock Summary**, **Order Status** and **Exploded BOM** reports exported from the "
    "ERP. The report is built automatically, in the same format as the manual Excel report."
)

with st.sidebar:
    st.header("How to use")
    st.markdown(
        "1. Upload each report in its own box: Stock Summary, Order Status, Exploded BOM.\n"
        "2. Check the preview and, if needed, change a **Remarks** value.\n"
        "3. Click **Download report**."
    )
    st.divider()
    st.divider()
    st.caption("Remarks rules")
    st.markdown(
        "- **Running** – open order still to produce, production already started\n"
        "- **Order Received** – open order still to produce, nothing made yet\n"
        "- **Multiline Stock** – nothing to produce; stock on 2+ envelope lines\n"
        "- **Leftover Stock** – nothing to produce\n\n"
        "Any remark can be changed in the preview before downloading."
    )

# ---------------------------------------------------------------------------
# Upload - one box per input report
# ---------------------------------------------------------------------------
SLOTS = [
    (KIND_STOCK, "1. Stock Summary Report", "e.g. StockSummaryReport-COIMBATORE-EOU-2026-11-03.xls"),
    (KIND_ORDER, "2. Order Status Report", "e.g. OrderStatusReport-03-11-2026.xlsx"),
    (KIND_BOM, "3. Exploded BOM Report", "e.g. Exploded_bom-20-08-2026.xlsx"),
]
files: dict[str, tuple[str, bytes]] = {}
problems = False
cols = st.columns(3)
for col, (kind, label, hint) in zip(cols, SLOTS):
    with col:
        up = st.file_uploader(label, type=["xlsx", "xls", "xlsm"], key=f"upload_{kind}", help=hint)
        if up is None:
            st.caption("⏳ Waiting for file")
            continue
        data = up.getvalue()
        found = _identify(data)
        if found == kind:
            files[kind] = (up.name, data)
            st.caption(f"✅ {KIND_LABELS[kind]} recognised")
        elif found in (KIND_STOCK, KIND_ORDER, KIND_BOM):
            problems = True
            st.error(f"This looks like the **{KIND_LABELS[found]}**. Please upload it in the "
                     f"'{dict((k, l) for k, l, _ in SLOTS)[found]}' box.")
        else:
            problems = True
            st.error(f"'{up.name}' is not a {KIND_LABELS[kind]}. Please check you exported the right "
                     "report from the ERP.")

required = [KIND_STOCK, KIND_ORDER, KIND_BOM]
if problems or not all(k in files for k in required):
    if not problems:
        st.info("Upload all three reports to build the Tag & Env Stock report.", icon="📂")
    st.stop()

# ---------------------------------------------------------------------------
# Build the report
# ---------------------------------------------------------------------------
digest = _digest(*(files[k][1] for k in required))
try:
    with st.status("Building the report...", expanded=False) as status:
        status.write("Reading Stock Summary...")
        stock = _stock(files[KIND_STOCK][1], files[KIND_STOCK][0])
        status.write("Reading Order Status...")
        orders = _orders(files[KIND_ORDER][1])
        status.write("Reading Exploded BOM (large file, may take a few seconds)...")
        bom = _bom(files[KIND_BOM][1])
        status.write("Applying business rules...")
        report = build_report(stock, orders, bom)
        status.update(label="Report ready", state="complete")
except InputError as exc:
    st.error(str(exc))
    st.stop()
except Exception as exc:  # pragma: no cover - shown to the user
    st.error("Something went wrong while processing the files. Please check that the three files are "
             f"the standard ERP exports.\n\nTechnical detail: {exc}")
    st.stop()

if st.session_state.get("digest") != digest:
    st.session_state["digest"] = digest
    st.session_state["remarks"] = {}

report_date = report.report_date or date.today()
totals = report.column_totals()
env_total = sum(v for k, v in totals.items() if k in C.ENV_COLUMNS)
tag_total = sum(v for k, v in totals.items() if k in C.TAG_COLUMNS)
export_total = sum(e.qty for e in report.export_lines)

m = st.columns(5)
m[0].metric("Stock date", report_date.strftime("%d-%m-%Y"))
m[1].metric("Report rows", len(report.rows))
m[2].metric("Envelope stock (kg)", f"{env_total:,.0f}")
m[3].metric("Tag stock (kg)", f"{tag_total:,.0f}")
m[4].metric("Ghana / Abidjan / Angola (kg)", f"{export_total:,.0f}")

issues = report.issues_frame()
blocking = issues[issues["Check"].isin(["Unknown machine line", "Order FG not in BOM"])]
if len(issues):
    st.info(f"{len(issues)} data check(s) to review - see the **Data checks** tab. "
            + (f"{len(blocking)} need attention." if len(blocking) else ""), icon="🔎")

download_slot = st.container()

# ---------------------------------------------------------------------------
# Preview
# ---------------------------------------------------------------------------
tab_main, tab_export, tab_checks, tab_info = st.tabs(
    ["Tag & Env Stock", "Ghana & AB & AG Stock", f"Data checks ({len(issues)})", "Sources"])

with tab_main:
    df = report.preview_frame()
    saved = st.session_state["remarks"]
    df["Remarks"] = [saved.get(k, r) for k, r in zip(df["_key"], df["Remarks"])]
    remark_filter = st.multiselect("Show remarks", C.REMARK_CHOICES[:4], default=C.REMARK_CHOICES[:4])
    search = st.text_input("Search item", placeholder="e.g. JUMBO")
    view = df[df["Remarks"].isin(remark_filter) | ~df["Remarks"].isin(C.REMARK_CHOICES[:4])]
    if search:
        view = view[view["Item"].str.contains(search, case=False, regex=False)]
    num_cols = [c for c in df.columns if c not in ("Item", "Remarks", "Why", "_key")]
    edited = st.data_editor(
        view,
        hide_index=True,
        width="stretch",
        height=560,
        disabled=[c for c in df.columns if c != "Remarks"],
        column_order=["Item"] + [c for c in num_cols if c != "Pending CFC"] +
                     ["Remarks", "Pending CFC", "Why"],
        column_config={
            "Item": st.column_config.TextColumn("Item", width="large", pinned=True),
            "Remarks": st.column_config.SelectboxColumn("Remarks", options=C.REMARK_CHOICES,
                                                        required=True, width="medium"),
            "Why": st.column_config.TextColumn("Why this remark", width="large"),
            **{c: st.column_config.NumberColumn(c, format="%.2f") for c in num_cols},
        },
        key=f"editor_{digest}",
    )
    for k, r, orig in zip(edited["_key"], edited["Remarks"],
                          df.set_index("_key").loc[edited["_key"], "Remarks"]):
        if r != orig or k in saved:
            saved[k] = r
    st.caption("Shared tags (e.g. TAG GV FLAVOUR) are shown on the first row of their group; in the "
               "Excel file the cell is merged across the group, as in the manual report.")

with tab_export:
    ex = report.export_frame()
    if ex.empty:
        st.write("No Ghana / Abidjan / Angola stock in this Stock Summary.")
    else:
        for section, grp in ex.groupby("Section", sort=False):
            st.subheader(section.title())
            st.dataframe(grp.drop(columns="Section"), hide_index=True, width="stretch")

with tab_checks:
    if issues.empty:
        st.success("No data issues found.")
    else:
        st.caption("These do not stop the report; they show where the inputs did not match cleanly. "
                   "They are shown here only; the downloaded report contains just the report sheets.")
        st.dataframe(issues, hide_index=True, width="stretch", height=480)

with tab_info:
    st.json(report.sources)

# ---------------------------------------------------------------------------
# Download
# ---------------------------------------------------------------------------
xlsx = report.to_excel(remarks=st.session_state["remarks"])
download_slot.download_button(
    "⬇️ Download report",
    data=xlsx,
    file_name=f"Tag_Env_Stock_as_on_{report_date.strftime('%d-%m-%y')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    width="stretch",
)
