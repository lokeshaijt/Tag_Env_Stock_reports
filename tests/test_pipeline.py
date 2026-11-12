"""Tests on small synthetic copies of the three ERP reports.

Run with:  python -m pytest -q
"""
import io
import sys
from pathlib import Path

import openpyxl
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from tag_env import run  # noqa: E402
from tag_env.grouping import family_key  # noqa: E402
from tag_env.readers import InputError, identify  # noqa: E402


def _xlsx(sheets: dict[str, list[list]]) -> bytes:
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for name, rows in sheets.items():
        ws = wb.create_sheet(name)
        for r in rows:
            ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


STOCK = _xlsx({"Stock SummaryNov 03,2026": [
    ["MJIL-COIMBATORE-EOU", None, "Stock Register: [Report on: 2026-11-03]", None],
    ["Item Name", "Machine Line Name", "Qty", "Amount"],
    ["PAPER 70 GSM", None, 9000, 1],
    ["ENV GV CLOVE DC C250 TBGS", "CONSTANTA-250", 97.22, 1],
    ["ENV GV CLOVE DC ENV TBGS", "Constanta Envelope", 950.2, 1],
    ["ENV GV MINT DC C250 TBGS", "CONSTANTA-250", 180.41, 1],
    ["TAG GV FLAVOUR DC C250 TBGS", "Constanta Tag", 141.75, 1],
    ["ENV AL FINA DC SF C250 TBGS", "CONSTANTA-250", 318.8, 1],
    ["ENV AL FINA DC ENV (70 GSM)", "Constanta Envelope", 191.3, 1],
    ["ENV AL FINA DC SF TBGS", "MD-20 Crimp Envelope", 144.96, 1],
    ["TAG AL FINA DC SF C250 TBGS", "Constanta Tag", 26.49, 1],
    ["TAG AL FINA DC SF TBGS", "MD-20 Tag", 44.03, 1],
    ["ENV KHAIMA MANGO DC TBGS", "Constanta Envelope", 98.76, 1],
    ["ENV GV MINT DC ENV TBGS (GHANA)", "Constanta Envelope", 25.42, 1],
    ["TAG GV DC TBGS (GHANA)", "Constanta Tag", 186.1, 1],
    ["ENV AL FINA DC ENV TBGS (ANGOLA)", "Constanta Envelope", 4898.39, 1],
    ["TAG ELDORADO C250 TAG TBGS", "Constanta Tag", 185.97, 1],
    ["ENV MYSTERY DC TBGS", "Some New Line", 5, 1],
]})

ORDERS = _xlsx({"Order Status-By Buyer": [
    [None, None, "Order Status-Buyer Wise Shipment Dt Wise On 03-11-2026"],
    [],
    ["Ord Id", "Order Status", "Prod Id", "Prod Name", "Pending Ord Qty", "Stock In Process",
     "Ready Stock", "Pending Prod In Cfc"],
    [1, "Planning Pending", 10, "GV CLOVE 24/25 DC SF ENV TBGS", 266, 0, 0, 266],
    [None, None, 11, "GV MINT 24/25 DC SF ENV TBGS", 522, 0, 100, 422],
    [None, None, 12, "KHAIMA MANGO 24/25 DC ENV TBGS", 300, 0, 300, 0],
    [None, None, None, None, 1088, 0, 400, 688],
]})

BOM = _xlsx({"Exploded BOM": [
    ["Exploded BOM (New)", None, None, None, None, None, None, None, None, None, "Run Date-20-08-2026"],
    ["ITEM ID", "FG Name", "ITEM ID", "Item Name", "ITEM TYPE", "ITEM GROUP", "Qty", "Uom", "Brand", "FG Uom"],
    [1, "GV CLOVE 24/25 DC SF ENV TBGS", 2, "ENV GV CLOVE DC C250 TBGS", "ENV", "PM", 600, "PCS", "GV", "CFC"],
    [1, "GV CLOVE 24/25 DC SF ENV TBGS", 3, "TAG GV FLAVOUR DC C250 TBGS", "TAG", "PM", 600, "PCS", "GV", "CFC"],
    [4, "GV MINT 24/25 DC SF ENV TBGS", 5, "ENV GV MINT DC C250 TBGS", "ENV", "PM", 600, "PCS", "GV", "CFC"],
    [4, "GV MINT 24/25 DC SF ENV TBGS", 3, "TAG GV FLAVOUR DC C250 TBGS", "TAG", "PM", 600, "PCS", "GV", "CFC"],
    [6, "KHAIMA MANGO 24/25 DC ENV TBGS", 7, "ENV KHAIMA MANGO DC TBGS", "ENV", "PM", 600, "PCS", "KH", "CFC"],
]})


@pytest.fixture(scope="module")
def report():
    return run(STOCK, ORDERS, BOM, "StockSummaryReport-COIMBATORE-EOU-2026-11-03_.xls")


def test_family_key():
    assert family_key("ENV AL FINA DC SF C250 TBGS") == "AL FINA"
    assert family_key("ENV AL FINA DC ENV (70 GSM)") == "AL FINA"
    assert family_key("TAG AL FINA DC SF TBGS") == "AL FINA"
    assert family_key("ENV GARANT EARL GREY DC SF ENV TBGS-FT") == "GARANT EARL GREY"
    assert family_key("ENV CHAYA DC PENV TBGS (WO GLUE)") == "CHAYA"
    assert family_key("ENV BOM DIA BLACK TEA DC SF C250 TBGS (MATADI)") == "BOM DIA BLACK TEA (MATADI)"


def test_identify():
    assert identify(STOCK) == "stock"
    assert identify(ORDERS) == "order"
    assert identify(BOM) == "bom"
    assert identify(b"not an excel file") is None


def test_wrong_file_gives_clear_error():
    with pytest.raises(InputError):
        run(ORDERS, ORDERS, BOM)


def test_rows_and_columns(report):
    rows = {r.name: r for r in report.rows}
    al = rows["ENV AL FINA DC SF C250 TBGS"]
    assert al.qty("CONSTANTA_ENV") == pytest.approx(191.3)
    assert al.qty("C250_ENV") == pytest.approx(318.8)
    assert al.qty("C250_TAG") == pytest.approx(26.49)
    assert al.qty("MD20_ENV") == pytest.approx(144.96)
    assert al.qty("MD20_TAG") == pytest.approx(44.03)
    clove = rows["ENV GV CLOVE DC C250 TBGS"]
    assert clove.qty("CONSTANTA_ENV") == pytest.approx(950.2)
    # the shared GV tag is not on a single row but merged across GV CLOVE + GV MINT
    assert clove.qty("C250_TAG") == 0
    (st,) = report.shared_tags
    assert st.qty == pytest.approx(141.75) and set(st.member_keys) == {"GV CLOVE", "GV MINT"}
    assert "TAG ELDORADO C250 TAG TBGS" in rows                     # stand-alone tag row
    assert report.rows[-1].kind == "tag"


def test_remarks(report):
    rem = {r.name: r.remark for r in report.rows}
    assert rem["ENV GV CLOVE DC C250 TBGS"] == "Order Received"      # pending, nothing made
    assert rem["ENV GV MINT DC C250 TBGS"] == "Running"              # pending, partly ready
    assert rem["ENV KHAIMA MANGO DC TBGS"] == "Leftover Stock"       # order fully ready
    assert rem["ENV AL FINA DC SF C250 TBGS"] == "Multiline Stock"   # no order, 3 env lines


def test_export_and_checks(report):
    sections = {(e.section, e.item) for e in report.export_lines}
    assert ("GHANA STOCK", "ENV GV MINT DC ENV TBGS (GHANA)") in sections
    assert ("GHANA STOCK", "TAG GV DC TBGS (GHANA)") in sections
    assert ("ANGOLA STOCK", "ENV AL FINA DC ENV TBGS (ANGOLA)") in sections
    checks = report.issues_frame()
    assert "Unknown machine line" in set(checks["Check"])
    assert "Reconciliation" not in set(checks["Check"])
    assert str(report.report_date) == "2026-11-03"


def test_excel_layout(report):
    data = report.to_excel(remarks={"KHAIMA MANGO": "Yet to Plan"})
    wb = openpyxl.load_workbook(io.BytesIO(data))
    # same sheets, order and visibility as the reference report
    assert wb.sheetnames == ["Tag & Env Stock (2)", "Env & Tag Stock", "Tag & Env Stock",
                             "Ghana & AB & AG Stock"]
    assert [wb[n].sheet_state for n in wb.sheetnames] == ["hidden", "visible", "hidden", "visible"]
    ws = wb["Tag & Env Stock (2)"]
    assert [ws[c].value for c in ("A2", "B2", "D2", "G2", "J2", "L2", "O2", "R2", "S2")] == \
        ["Brand", "Constanta", "C250", "MD20", "Perfecta ", "MD20 HS", "Maisa", "Perfecta/PK45 Tag", "Remarks"]
    names = [ws.cell(row=r, column=1).value for r in range(4, ws.max_row + 1)]
    assert names[-2:] == ["TOTAL", "Grand Total"]
    khaima = names.index("ENV KHAIMA MANGO DC TBGS") + 4
    assert ws.cell(row=khaima, column=19).value == "Yet to Plan"   # user override kept
    assert any(str(m).startswith("E") for m in ws.merged_cells.ranges if m.min_row > 3)  # shared tag
    assert wb["Tag & Env Stock"]["A5"].value != ws["A5"].value     # short names on this sheet
    assert not wb._external_links                                  # no links to other workbooks


def test_make_cases_formula(report):
    data = report.to_excel()
    ws = openpyxl.load_workbook(io.BytesIO(data))["Env & Tag Stock"]
    assert ws["E4"].value == 1740 and ws["O4"].value == 1525 and ws["C4"].value == 10780
    rows = {ws.cell(row=r, column=1).value: r for r in range(5, ws.max_row)}
    r = rows["ENV GV CLOVE DC C250 TBGS".replace("DC C250 TBGS", "")]
    # (KG x conversion factor) / pieces per case from the BOM (600 for GV CLOVE 24/25)
    assert ws[f"E{r}"].value == f"=(D{r}*$E$4)/600"
    assert ws[f"F{r}"].value == f"=MAX(C{r},E{r})"
    cases = report.make_cases_frame()
    clove = cases[cases["Item"] == "ENV GV CLOVE DC C250 TBGS"].iloc[0]
    assert clove["C250 Env Cases"] == round(97.22 * 1740 / 600)
    # no BOM value -> left empty, not a wrong number
    r = rows["ENV AL FINA "]
    assert ws[f"E{r}"].value is None


def test_format_matches_reference(report):
    from tag_env.format_check import compare
    assert compare(report.to_excel()) == []
