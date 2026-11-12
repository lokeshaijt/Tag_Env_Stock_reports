"""Tag & Env Stock Report - Streamlit app.

Upload the Stock Summary, Order Status and Exploded BOM reports, click
Generate Report, check the preview and download the report. The Excel file is
the reference Tag & Env Stock workbook with only its data replaced.
"""
from __future__ import annotations

import hashlib
from datetime import date

import streamlit as st

from tag_env import config as C
from tag_env.format_check import compare
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


@st.cache_data(show_spinner=False, max_entries=6)
def _format_check(xlsx: bytes) -> list[str]:
    return compare(xlsx)


def _digest(*parts: bytes | None) -> str:
    h = hashlib.sha1()
    for p in parts:
        h.update(p or b"-")
    return h.hexdigest()


# ---------------------------------------------------------------------------
# Header and sidebar
# ---------------------------------------------------------------------------
st.title("📦 Tag & Env Stock Report")
st.write("Upload the three ERP reports and click **Generate Report**. The Excel file has the same "
         "sheets and format as the manual Tag & Env Stock report.")

with st.sidebar:
    st.header("How to use")
    st.markdown(
        "1. Upload each report in its own box.\n"
        "2. Click **Generate Report**.\n"
        "3. Check the preview; change a **Remarks** value if needed.\n"
        "4. Click **Download report**."
    )
    st.divider()
    with st.expander("Conversion factors (pcs per kg)"):
        st.caption("From the Conversion Round Off chart. Used in the Make Cases formula: "
                   "(KG × conversion factor) ÷ Exploded BOM pieces per case.")
        conversions = {k: float(st.number_input(k, min_value=1, value=int(v), step=10, key=f"conv_{k}"))
                       for k, v in C.DEFAULT_CONVERSIONS.items()}
    st.divider()
    st.caption("Remarks rules")
    st.markdown(
        "- **Running** – open order still to produce, production already started\n"
        "- **Order Received** – open order still to produce, nothing made yet\n"
        "- **Multiline Stock** – nothing to produce; stock on 2+ envelope lines\n"
        "- **Leftover Stock** – nothing to produce"
    )

# ---------------------------------------------------------------------------
# 1. Upload - one box per input report, each validated
# ---------------------------------------------------------------------------
SLOTS = [
    (KIND_STOCK, "1. Stock Summary Report", "e.g. StockSummaryReport-COIMBATORE-EOU-2026-11-03.xls"),
    (KIND_ORDER, "2. Order Status Report", "e.g. OrderStatusReport-03-11-2026.xlsx"),
    (KIND_BOM, "3. Exploded BOM Report", "e.g. Exploded_bom-20-08-2026.xlsx"),
]
files: dict[str, tuple[str, bytes]] = {}
problems = False
for col, (kind, label, hint) in zip(st.columns(3), SLOTS):
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
            st.error(f"'{up.name}' is not a {KIND_LABELS[kind]} (the expected columns were not found). "
                     "Please check you exported the right report from the ERP.")

required = [KIND_STOCK, KIND_ORDER, KIND_BOM]
ready = not problems and all(k in files for k in required)
digest = _digest(*(files[k][1] for k in required)) if ready else None

# ---------------------------------------------------------------------------
# 2. Generate
# ---------------------------------------------------------------------------
clicked = st.button("▶️ Generate Report", type="primary", disabled=not ready, width="stretch")
if not ready:
    st.info("Upload all three reports to enable **Generate Report**.", icon="📂")
    st.stop()
if clicked:
    st.session_state["generated"] = digest
    st.session_state["remarks"] = {}
if st.session_state.get("generated") != digest:
    st.stop()

try:
    with st.status("Generating the report...", expanded=False) as status:
        status.write("Reading Stock Summary...")
        stock = _stock(files[KIND_STOCK][1], files[KIND_STOCK][0])
        status.write("Reading Order Status...")
        orders = _orders(files[KIND_ORDER][1])
        status.write("Reading Exploded BOM (large file, up to ~30 seconds the first time)...")
        bom = _bom(files[KIND_BOM][1])
        status.write("Matching items, orders and BOM; calculating Make Cases...")
        report = build_report(stock, orders, bom)
        status.update(label="Report generated", state="complete")
except InputError as exc:
    st.error(str(exc))
    st.stop()
except Exception as exc:  # pragma: no cover - shown to the user
    st.error("Something went wrong while processing the files. Please check that they are the "
             f"standard ERP exports.\n\nTechnical detail: {exc}")
    st.stop()

# ---------------------------------------------------------------------------
# 3. Summary
# ---------------------------------------------------------------------------
report_date = report.report_date or date.today()
totals = report.column_totals()
issues = report.issues_frame()
m = st.columns(5)
m[0].metric("Stock date", report_date.strftime("%d-%m-%Y"))
m[1].metric("Report rows", len(report.rows))
m[2].metric("Envelope stock (kg)", f"{sum(v for k, v in totals.items() if k in C.ENV_COLUMNS):,.0f}")
m[3].metric("Tag stock (kg)", f"{sum(v for k, v in totals.items() if k in C.TAG_COLUMNS):,.0f}")
m[4].metric("Ghana / Abidjan / Angola (kg)", f"{sum(e.qty for e in report.export_lines):,.0f}")

summary = issues["Check"].value_counts()
if len(summary):
    st.warning("Processing summary - " + ", ".join(f"{n} × {k}" for k, n in summary.items())
               + ". Details in the **Data checks** tab.", icon="🔎")
else:
    st.success("All inputs matched cleanly.", icon="✅")

download_slot = st.container()

# ---------------------------------------------------------------------------
# 4. Preview
# ---------------------------------------------------------------------------
tab_make, tab_stock, tab_export, tab_checks, tab_format = st.tabs(
    ["Env & Tag Stock (Make Cases)", "Tag & Env Stock", "Ghana & AB & AG Stock",
     f"Data checks ({len(issues)})", "Format check"])

saved = st.session_state.setdefault("remarks", {})

with tab_stock:
    df = report.preview_frame()
    df["Remarks"] = [saved.get(k, r) for k, r in zip(df["_key"], df["Remarks"])]
    search = st.text_input("Search item", placeholder="e.g. JUMBO", key="search_stock")
    view = df[df["Item"].str.contains(search, case=False, regex=False)] if search else df
    num_cols = [c for c in df.columns if c not in ("Item", "Remarks", "Why", "_key", "Pending CFC")]
    edited = st.data_editor(
        view, hide_index=True, width="stretch", height=560,
        disabled=[c for c in df.columns if c != "Remarks"],
        column_order=["Item"] + num_cols + ["Remarks", "Pending CFC", "Why"],
        column_config={
            "Item": st.column_config.TextColumn("Item", width="large", pinned=True),
            "Remarks": st.column_config.SelectboxColumn("Remarks", options=C.REMARK_CHOICES,
                                                        required=True, width="medium"),
            "Why": st.column_config.TextColumn("Why this remark", width="large"),
            **{c: st.column_config.NumberColumn(c, format="%.2f") for c in num_cols},
        },
        key=f"editor_{digest}",
    )
    original = df.set_index("_key")["Remarks"]
    for k, r in zip(edited["_key"], edited["Remarks"]):
        if r != original[k] or k in saved:
            saved[k] = r
    st.caption("Remarks changed here are written to both sheets of the downloaded report.")

with tab_make:
    mc = report.make_cases_frame(conversions)
    mc["Remarks"] = [saved.get(r.key, r.remark) for r in report.rows]
    st.caption("Make Cases = (KG × machine-line conversion factor) ÷ pieces per case from the "
               "Exploded BOM. Empty = conversion factor or BOM value missing (see Data checks).")
    st.dataframe(mc, hide_index=True, width="stretch", height=560,
                 column_config={"Item": st.column_config.TextColumn("Item", pinned=True)})

with tab_export:
    ex = report.export_frame()
    if ex.empty:
        st.write("No Ghana / Abidjan / Angola stock in this Stock Summary.")
    for section, grp in ex.groupby("Section", sort=False):
        st.subheader(section.title())
        st.dataframe(grp.drop(columns="Section"), hide_index=True, width="stretch")

with tab_checks:
    if issues.empty:
        st.success("No data issues found.")
    else:
        st.caption("These do not stop the report; they show where the inputs did not match cleanly.")
        st.dataframe(issues, hide_index=True, width="stretch", height=480)
    st.json(report.sources, expanded=False)

# ---------------------------------------------------------------------------
# 5. Excel + format check + download
# ---------------------------------------------------------------------------
xlsx = report.to_excel(remarks=saved, conversions=conversions)
with tab_format:
    diffs = _format_check(xlsx)
    if diffs:
        st.error(f"{len(diffs)} formatting difference(s) from the reference report:")
        st.write(diffs)
    else:
        st.success("The generated workbook matches the reference Tag & Env Stock report: same sheets "
                   "(and visibility), column widths, frozen panes, headers and merged headers, and the "
                   "same font, fill, border, alignment and number format on every data and total cell.")

download_slot.download_button(
    "⬇️ Download report",
    data=xlsx,
    file_name=f"Tag_Env_Stock_as_on_{report_date.strftime('%d-%m-%y')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
    width="stretch",
)
