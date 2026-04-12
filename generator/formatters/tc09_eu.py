"""Formatter: TC-09-EU — European OECD Transfer Pricing Documentation.

Emits:
- test_cases/TC-09-EU/input_files/intercompany_transactions_eu_fy2025.xlsx
  All FY2025 intercompany transactions across Cascade Europe group (~120 rows)
- test_cases/TC-09-EU/input_files/comparable_companies_eu.xlsx
  Manufacturing (15) and distribution (10) comparables in two sheets
- test_cases/TC-09-EU/input_files/interest_rate_benchmarks_eu.xlsx
  EURIBOR rates + BBB credit spreads with ERR-EU-005 planted
- test_cases/TC-09-EU/input_files/master_file_fy2024.pdf
  Prior year OECD master file (28 pages)
- test_cases/TC-09-EU/input_files/local_file_cp_fy2024.pdf
  Prior year local file for CP — Germany (35 pages)
- test_cases/TC-09-EU/prompt.md
- test_cases/TC-09-EU/expected_behavior.md
- gold_standards/TC-09-EU_gold.json

Planted error:
  ERR-EU-005: rounding_discrepancy — Q3 FY2025 EURIBOR 12M rate entered as
  0.38% instead of 3.85% (decimal point error).

Uses deterministic EU transfer pricing model — never hardcodes numbers
that should come from the model.
"""

from __future__ import annotations

import datetime
import io
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

from generator.canaries import CanaryRegistry, embed_canary_xlsx
from generator.errors import ErrorRegistry, PlantedError
from generator.golds.framework import GoldStandard, register_gold
from generator.manifest import Manifest
from generator.model.build import CascadeModel
from generator.model.tp_eu import (
    DIST_COMPARABLES,
    ENTITY_JURISDICTIONS,
    ENTITY_NAMES,
    EURIBOR_ERR_EU_005_WRONG,
    FINISHED_GOODS_MARKUP_EU,
    IC_LOAN_PRINCIPAL_EU,
    IC_LOAN_RATE_EU,
    MANAGEMENT_FEE_PCT_EU,
    MFG_COMPARABLES,
    RAW_MATERIALS_MARKUP_EU,
    ROYALTY_RATE_EU,
    compute_dist_iqr,
    compute_mfg_iqr,
    generate_credit_spread_data,
    generate_euribor_data,
    generate_ic_transactions_eu,
)

# ── Constants ────────────────────────────────────────────────────────────────

_TC = "TC-09-EU"
_INPUT_DIR = f"test_cases/{_TC}/input_files"

_FIXED_DATETIME = datetime.datetime(2025, 3, 15, 9, 0, 0)
_FIXED_ZIP_DT = (2025, 3, 15, 9, 0, 0)

# Canary keys (from model/tp_eu.py ALL_CANARY_KEYS_TC09EU)
_KEY_IC_TRANSACTIONS = "tc09eu_intercompany_transactions"
_KEY_COMPARABLE_COMPANIES = "tc09eu_comparable_companies"
_KEY_MASTER_FILE = "tc09eu_master_file_fy2024"
_KEY_LOCAL_FILE_CP = "tc09eu_local_file_cp_fy2024"
_KEY_INTEREST_BENCHMARKS = "tc09eu_interest_rate_benchmarks"

# Styles
_HEADER_FILL = PatternFill("solid", fgColor="1A3C6E")
_HEADER_FONT = Font(bold=True, size=11, color="FFFFFF")
_REJECTED_FILL = PatternFill("solid", fgColor="FDE8E8")
_NUM_FMT = "#,##0"
_NUM_FMT_2 = "#,##0.00"
_PCT_FMT = "0.00%"
_THIN_BORDER = Border(bottom=Side(style="thin", color="CCCCCC"))

# Entity role descriptions (used in PDFs)
_ENTITY_ROLES = {
    "CE": "Holding company \u2014 strategic oversight, treasury, legal",
    "CP": "Licensed manufacturer \u2014 precision components",
    "CM": "R&D centre and IP developer \u2014 advanced materials",
    "CD": "Limited-risk distributor \u2014 warehousing and logistics",
}


# ── Deterministic save helpers ───────────────────────────────────────────────


def _save_xlsx_deterministic(wb: openpyxl.Workbook, path: str | Path) -> None:
    """Save workbook with pinned timestamps and fixed zip entry dates."""
    from openpyxl.writer.excel import ExcelWriter

    path = Path(path)
    wb.properties.created = _FIXED_DATETIME
    wb.properties.modified = _FIXED_DATETIME

    buf = io.BytesIO()
    archive = ZipFile(buf, "w", ZIP_DEFLATED, allowZip64=True)
    writer = ExcelWriter(wb, archive)
    writer.save()

    buf.seek(0)
    with ZipFile(buf, "r") as src, ZipFile(str(path), "w", ZIP_DEFLATED) as dst:
        for item in src.infolist():
            info = ZipInfo(item.filename, date_time=_FIXED_ZIP_DT)
            info.compress_type = item.compress_type
            dst.writestr(info, src.read(item.filename))


def _save_pdf_deterministic(buf: io.BytesIO, path: Path) -> None:
    """Write PDF buffer to disk (reportlab invariant=True handles determinism)."""
    buf.seek(0)
    path.write_bytes(buf.getvalue())


# ── Intercompany Transactions XLSX ───────────────────────────────────────────


def _write_ic_transactions_xlsx(
    output_dir: Path,
    canaries: CanaryRegistry,
    manifest: Manifest,
) -> None:
    """Write intercompany_transactions_eu_fy2025.xlsx."""
    txns = generate_ic_transactions_eu()
    canary = canaries.canary_for(_KEY_IC_TRANSACTIONS)

    wb = openpyxl.Workbook()
    loc = embed_canary_xlsx(wb, canary)

    ws = wb.active
    ws.title = "Transactions"

    # Title rows 1-2
    ws["A1"] = "Cascade Europe Holdings B.V. \u2014 Intercompany Transactions FY2025"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:J1")
    ws["A2"] = "All intercompany flows across CE, CP, CM, CD entities"
    ws["A2"].font = Font(bold=True, size=11, color="666666")
    ws.merge_cells("A2:J2")

    # Headers at row 3
    headers = [
        "Transaction ID", "From Entity", "To Entity", "Transaction Type",
        "Description", "Volume/Principal (EUR)", "Price/Rate",
        "Total Amount (EUR)", "Currency", "Arm's Length Method",
    ]
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=header)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    for idx, txn in enumerate(txns):
        row = idx + 4
        txn_id = f"IC-EU-{idx:04d}"
        ws.cell(row=row, column=1, value=txn_id)
        ws.cell(row=row, column=2, value=txn.from_entity)
        ws.cell(row=row, column=3, value=txn.to_entity)
        ws.cell(row=row, column=4, value=txn.transaction_type)
        ws.cell(row=row, column=5, value=txn.description)
        cell = ws.cell(row=row, column=6, value=float(txn.volume_or_principal))
        cell.number_format = _NUM_FMT_2
        cell.border = _THIN_BORDER
        cell = ws.cell(row=row, column=7, value=float(txn.price_or_rate))
        cell.number_format = _PCT_FMT
        cell.border = _THIN_BORDER
        cell = ws.cell(row=row, column=8, value=float(txn.total_amount_eur))
        cell.number_format = _NUM_FMT
        cell.border = _THIN_BORDER
        ws.cell(row=row, column=9, value=txn.invoicing_currency)
        ws.cell(row=row, column=10, value=txn.arm_length_method)

    widths = [16, 14, 14, 18, 55, 22, 14, 20, 10, 26]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    rel_path = f"{_INPUT_DIR}/intercompany_transactions_eu_fy2025.xlsx"
    abs_path = output_dir / rel_path
    abs_path.parent.mkdir(parents=True, exist_ok=True)
    _save_xlsx_deterministic(wb, abs_path)
    canaries.set_location(_KEY_IC_TRANSACTIONS, rel_path, loc)
    manifest.register(rel_path, "xlsx")


# ── Comparable Companies XLSX ────────────────────────────────────────────────


def _write_comparable_companies_xlsx(
    output_dir: Path,
    canaries: CanaryRegistry,
    manifest: Manifest,
) -> None:
    """Write comparable_companies_eu.xlsx with two sheets."""
    canary = canaries.canary_for(_KEY_COMPARABLE_COMPANIES)
    wb = openpyxl.Workbook()
    loc = embed_canary_xlsx(wb, canary)

    # Sheet 1: Manufacturing Comparables
    ws1 = wb.active
    ws1.title = "Manufacturing Comparables"
    mfg_headers = [
        "Company Name", "Country", "SIC Code", "Revenue (EUR M)",
        "COGS (EUR M)", "Operating Expenses (EUR M)",
        "Operating Income (EUR M)", "Total Assets (EUR M)", "ROCE (%)",
        "Accepted/Rejected", "Rejection Reason",
    ]
    for col_idx, header in enumerate(mfg_headers, start=1):
        cell = ws1.cell(row=1, column=col_idx, value=header)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    for row_idx, comp in enumerate(MFG_COMPARABLES, start=2):
        (name, country, sic, rev, cogs, opex, oper_inc,
         assets, roce, is_rejected, reason) = comp
        ws1.cell(row=row_idx, column=1, value=name)
        ws1.cell(row=row_idx, column=2, value=country)
        ws1.cell(row=row_idx, column=3, value=sic)
        for ci, val in [(4, rev), (5, cogs), (6, opex), (7, oper_inc), (8, assets)]:
            cell = ws1.cell(row=row_idx, column=ci, value=val)
            cell.number_format = _NUM_FMT
            cell.border = _THIN_BORDER
        cell = ws1.cell(row=row_idx, column=9, value=roce)
        cell.number_format = "0.0"
        cell.border = _THIN_BORDER
        ws1.cell(row=row_idx, column=10, value="Rejected" if is_rejected else "Accepted")
        ws1.cell(row=row_idx, column=11, value=reason)
        if is_rejected:
            for c in range(1, 12):
                ws1.cell(row=row_idx, column=c).fill = _REJECTED_FILL

    for i, w in enumerate([35, 12, 10, 16, 14, 22, 22, 18, 10, 18, 55], start=1):
        ws1.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    # Sheet 2: Distribution Comparables
    ws2 = wb.create_sheet("Distribution Comparables")
    dist_headers = [
        "Company Name", "Country", "SIC Code", "Revenue (EUR M)",
        "COGS (EUR M)", "Operating Expenses (EUR M)",
        "Net Income (EUR M)", "Total Assets (EUR M)",
        "Accepted/Rejected", "Rejection Reason",
    ]
    for col_idx, header in enumerate(dist_headers, start=1):
        cell = ws2.cell(row=1, column=col_idx, value=header)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    for row_idx, comp in enumerate(DIST_COMPARABLES, start=2):
        (name, country, sic, rev, cogs, opex, net_inc,
         assets, is_rejected, reason) = comp
        ws2.cell(row=row_idx, column=1, value=name)
        ws2.cell(row=row_idx, column=2, value=country)
        ws2.cell(row=row_idx, column=3, value=sic)
        for ci, val in [(4, rev), (5, cogs), (6, opex)]:
            cell = ws2.cell(row=row_idx, column=ci, value=val)
            cell.number_format = _NUM_FMT
            cell.border = _THIN_BORDER
        cell = ws2.cell(row=row_idx, column=7, value=float(net_inc))
        cell.number_format = _NUM_FMT_2
        cell.border = _THIN_BORDER
        cell = ws2.cell(row=row_idx, column=8, value=assets)
        cell.number_format = _NUM_FMT
        cell.border = _THIN_BORDER
        ws2.cell(row=row_idx, column=9, value="Rejected" if is_rejected else "Accepted")
        ws2.cell(row=row_idx, column=10, value=reason)
        if is_rejected:
            for c in range(1, 11):
                ws2.cell(row=row_idx, column=c).fill = _REJECTED_FILL

    for i, w in enumerate([35, 12, 10, 16, 14, 22, 18, 18, 18, 55], start=1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    rel_path = f"{_INPUT_DIR}/comparable_companies_eu.xlsx"
    abs_path = output_dir / rel_path
    abs_path.parent.mkdir(parents=True, exist_ok=True)
    _save_xlsx_deterministic(wb, abs_path)
    canaries.set_location(_KEY_COMPARABLE_COMPANIES, rel_path, loc)
    manifest.register(rel_path, "xlsx")


# ── Interest Rate Benchmarks XLSX ────────────────────────────────────────────


def _write_interest_rate_benchmarks_xlsx(
    output_dir: Path,
    canaries: CanaryRegistry,
    errors: ErrorRegistry,
    manifest: Manifest,
) -> None:
    """Write interest_rate_benchmarks_eu.xlsx with ERR-EU-005 planted."""
    canary = canaries.canary_for(_KEY_INTEREST_BENCHMARKS)
    euribor_data = generate_euribor_data()
    credit_data = generate_credit_spread_data()

    wb = openpyxl.Workbook()
    loc = embed_canary_xlsx(wb, canary)

    # Sheet 1: EURIBOR Rates
    ws1 = wb.active
    ws1.title = "EURIBOR Rates"
    for col_idx, header in enumerate(
            ["Period", "Year", "Quarter", "Tenor", "Rate (%)"], start=1):
        cell = ws1.cell(row=1, column=col_idx, value=header)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    err_row = None
    for row_idx, entry in enumerate(euribor_data, start=2):
        ws1.cell(row=row_idx, column=1, value=entry.period)
        ws1.cell(row=row_idx, column=2, value=entry.year)
        ws1.cell(row=row_idx, column=3, value=entry.quarter)
        ws1.cell(row=row_idx, column=4, value=entry.tenor)
        rate_val = entry.rate_pct
        if entry.year == 2025 and entry.quarter == 3 and entry.tenor == "12M":
            rate_val = EURIBOR_ERR_EU_005_WRONG
            err_row = row_idx
        cell = ws1.cell(row=row_idx, column=5, value=float(rate_val))
        cell.number_format = "0.00"
        cell.border = _THIN_BORDER

    for i, w in enumerate([14, 8, 10, 8, 12], start=1):
        ws1.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    # Sheet 2: Credit Spreads
    ws2 = wb.create_sheet("Credit Spreads")
    for col_idx, header in enumerate(
            ["Period", "Year", "Quarter", "Spread (bps)"], start=1):
        cell = ws2.cell(row=1, column=col_idx, value=header)
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", wrap_text=True)

    for row_idx, entry in enumerate(credit_data, start=2):
        ws2.cell(row=row_idx, column=1, value=entry.period)
        ws2.cell(row=row_idx, column=2, value=entry.year)
        ws2.cell(row=row_idx, column=3, value=entry.quarter)
        cell = ws2.cell(row=row_idx, column=4, value=entry.spread_bps)
        cell.number_format = _NUM_FMT
        cell.border = _THIN_BORDER

    for i, w in enumerate([14, 8, 10, 14], start=1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(i)].width = w

    rel_path = f"{_INPUT_DIR}/interest_rate_benchmarks_eu.xlsx"
    abs_path = output_dir / rel_path
    abs_path.parent.mkdir(parents=True, exist_ok=True)
    _save_xlsx_deterministic(wb, abs_path)
    canaries.set_location(_KEY_INTEREST_BENCHMARKS, rel_path, loc)
    manifest.register(rel_path, "xlsx")

    errors.add(PlantedError(
        error_id="ERR-EU-005",
        file=f"{_INPUT_DIR}/interest_rate_benchmarks_eu.xlsx",
        location=f"Sheet 'EURIBOR Rates', Row {err_row}, Q3 FY2025 12M",
        type="rounding_discrepancy",
        description=(
            "EURIBOR 12M rate for Q3 FY2025 entered as 0.38% instead of "
            "3.85% (decimal point error). This pulls down the average and "
            "makes the 4.5% intercompany loan rate appear further above market."
        ),
        severity="material",
        which_test_cases_should_catch=[_TC],
    ))


# ── PDF helpers ──────────────────────────────────────────────────────────────

_TABLE_HDR_COLOR = colors.HexColor("#1A3C6E")


def _pdf_styles():
    """Return shared PDF styles."""
    styles = getSampleStyleSheet()
    return (
        ParagraphStyle("TPTitle", parent=styles["Title"], fontSize=18, spaceAfter=20),
        ParagraphStyle("TPH1", parent=styles["Heading1"], fontSize=14, spaceBefore=12, spaceAfter=8),
        ParagraphStyle("TPH2", parent=styles["Heading2"], fontSize=12, spaceBefore=10, spaceAfter=6),
        ParagraphStyle("TPBody", parent=styles["Normal"], fontSize=10, leading=14, spaceAfter=8),
        ParagraphStyle("TPSmall", parent=styles["Normal"], fontSize=8, textColor=colors.gray),
    )


def _tbl(data, widths, *, alt=False):
    """Build a styled reportlab Table."""
    t = Table(data, colWidths=widths)
    cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), _TABLE_HDR_COLOR),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]
    if alt:
        cmds.append(("ROWBACKGROUNDS", (0, 1), (-1, -1),
                      [colors.white, colors.HexColor("#D6E4F0")]))
    t.setStyle(TableStyle(cmds))
    return t


# ── Master File PDF (28 pages) ───────────────────────────────────────────────


def _write_master_file_pdf(
    output_dir: Path,
    canaries: CanaryRegistry,
    manifest: Manifest,
) -> None:
    """Write master_file_fy2024.pdf — 28-page OECD BEPS Action 13 master file."""
    canary = canaries.canary_for(_KEY_MASTER_FILE)
    file_path = output_dir / _INPUT_DIR / "master_file_fy2024.pdf"
    file_path.parent.mkdir(parents=True, exist_ok=True)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=letter, leftMargin=inch, rightMargin=inch,
        topMargin=inch, bottomMargin=inch, invariant=True)
    doc.title = "OECD Master File \u2014 Cascade Europe Holdings B.V. FY2024"
    doc.author = f"CANARY: {canary}"
    doc.subject = "Transfer Pricing Master File (BEPS Action 13)"
    doc.creator = "Cascade Industries Test Suite Generator"

    title_s, h1, h2, body, small = _pdf_styles()
    story: list[Any] = []

    # Page 1: Title
    story.append(Spacer(1, 2 * inch))
    story.append(Paragraph(
        "OECD Transfer Pricing Master File<br/>"
        "Cascade Europe Holdings B.V.<br/>"
        "Fiscal Year Ended December 31, 2024", title_s))
    story.append(Spacer(1, 0.5 * inch))
    story.append(Paragraph(
        "Prepared in accordance with OECD Transfer Pricing Guidelines (2022)<br/>"
        "and BEPS Action 13 Master File requirements", body))
    story.append(Spacer(1, 0.5 * inch))
    story.append(Paragraph("CONFIDENTIAL \u2014 FOR INTERNAL USE ONLY",
        ParagraphStyle("C", parent=body, fontSize=10, textColor=colors.red, alignment=1)))
    story.append(Spacer(1, 0.3 * inch))
    story.append(Paragraph(f"Document ID: {canary}", small))
    story.append(PageBreak())

    # Page 2: ToC
    story.append(Paragraph("Table of Contents", h1))
    toc = [("1. Organizational Structure", "3"), ("2. Group Business Description", "6"),
           ("3. Intangibles Overview", "11"), ("4. Intercompany Financial Activities", "15"),
           ("5. Financial and Tax Positions", "19"), ("Appendix A: Entity Legal Details", "25"),
           ("Appendix B: Intercompany Agreement List", "27")]
    toc_d = [["Section", "Page"]] + [list(r) for r in toc]
    t = Table(toc_d, colWidths=[5*inch, 1*inch])
    t.setStyle(TableStyle([("FONTSIZE",(0,0),(-1,-1),10), ("TEXTCOLOR",(0,0),(-1,0),_TABLE_HDR_COLOR),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"), ("LINEBELOW",(0,0),(-1,0),1,_TABLE_HDR_COLOR),
        ("ALIGN",(1,0),(1,-1),"RIGHT"), ("BOTTOMPADDING",(0,0),(-1,-1),4)]))
    story.append(t)
    story.append(PageBreak())

    # Pages 3-5: Organizational Structure
    story.append(Paragraph("1. Organizational Structure", h1))
    story.append(Paragraph(
        f'{ENTITY_NAMES["CE"]} ("CE" or the "Group") is a Dutch holding company '
        f'headquartered in Amsterdam, {ENTITY_JURISDICTIONS["CE"]}. CE is the ultimate '
        "European parent that holds 100% of the equity interests in three operating "
        "subsidiaries across Europe.", body))
    story.append(Paragraph("Group Structure", h2))
    org = [["Entity", "Jurisdiction", "Activity", "FY2024 Revenue"],
           [ENTITY_NAMES["CE"], ENTITY_JURISDICTIONS["CE"], "Holding / Management", "\u20ac22M"],
           [ENTITY_NAMES["CP"], ENTITY_JURISDICTIONS["CP"], "Manufacturing", "~\u20ac45M"],
           [ENTITY_NAMES["CM"], ENTITY_JURISDICTIONS["CM"], "R&D / Materials", "~\u20ac32M"],
           [ENTITY_NAMES["CD"], ENTITY_JURISDICTIONS["CD"], "Distribution", "~\u00a318M (~\u20ac21M)"]]
    story.append(_tbl(org, [2.2*inch, 1.2*inch, 1.5*inch, 1.5*inch], alt=True))
    story.append(Spacer(1, 0.3*inch))
    story.append(Paragraph(
        "CE was incorporated in the Netherlands in 2018 as a wholly-owned European subsidiary "
        "of Cascade Industries, Inc. (Portland, Oregon, USA). The Group's European operations "
        "serve the industrial manufacturing and distribution markets across the EU and UK.", body))
    story.append(Paragraph(
        f"The organizational structure reflects a functional specialization model: "
        f"{ENTITY_NAMES['CP']} performs licensed manufacturing under IP developed by "
        f"{ENTITY_NAMES['CM']}, {ENTITY_NAMES['CD']} operates as a limited-risk distributor, "
        "and CE provides centralised management services and treasury functions.", body))
    story.append(PageBreak())

    story.append(Paragraph("1. Organizational Structure (continued)", h1))
    story.append(Paragraph("Ownership Chain", h2))
    story.append(Paragraph(
        f"Cascade Industries, Inc. (USA) \u2192 100% \u2192 {ENTITY_NAMES['CE']} (NL) "
        "\u2192 100% \u2192 CP (DE), CM (FR), CD (UK)", body))
    story.append(Paragraph("Key Changes in FY2024", h2))
    story.append(Paragraph(
        "No material changes to the organizational structure occurred during FY2024. "
        "CD completed its first full year of operations following its establishment in FY2023. "
        "The UK entity operates under post-Brexit trade arrangements with goods flowing from "
        "CP (Germany) to CD (UK).", body))
    story.append(Paragraph("Functional Profiles", h2))
    for code in sorted(ENTITY_NAMES):
        story.append(Paragraph(f"<b>{ENTITY_NAMES[code]} ({code})</b>: {_ENTITY_ROLES[code]}", body))
    story.append(PageBreak())

    # Pages 6-10: Group Business Description
    story.append(Paragraph("2. Group Business Description", h1))
    story.append(Paragraph(
        "The Cascade Europe group operates in the European industrial manufacturing and "
        f"distribution sector. The group's value chain comprises R&amp;D and advanced materials "
        f"development ({ENTITY_NAMES['CM']}), precision component manufacturing "
        f"({ENTITY_NAMES['CP']}), and distribution to European end customers "
        f"({ENTITY_NAMES['CD']}).", body))
    story.append(Paragraph("Business Lines", h2))
    story.append(Paragraph(
        "&bull; <b>Advanced Materials R&amp;D</b> (CM): Development of specialty materials, "
        "coatings, and composites. CM holds the group's core IP and licenses technology to CP.<br/>"
        "&bull; <b>Precision Manufacturing</b> (CP): Production of precision-machined components "
        "using CM's proprietary processes.<br/>"
        "&bull; <b>Distribution</b> (CD): Warehousing, logistics, and last-mile delivery to UK "
        "and Irish customers.", body))
    story.append(Paragraph("Market Position", h2))
    story.append(Paragraph(
        "The group holds a mid-market position in the European industrial components sector, "
        "with combined revenue of approximately \u20ac120M.", body))
    story.append(PageBreak())

    for pg in range(7, 11):
        story.append(Paragraph(f"2. Group Business Description (continued \u2014 p.{pg})", h1))
        story.append(Paragraph("Supply Chain Integration", h2))
        story.append(Paragraph(
            "The group's supply chain is integrated across three jurisdictions. Raw materials "
            "flow from CM (Lyon) to CP (Munich), where they are processed into finished products. "
            "Finished goods destined for the UK market are shipped from CP to CD (Birmingham). "
            "CE coordinates group-wide procurement and treasury.", body))
        story.append(Paragraph("Competitive Dynamics", h2))
        story.append(Paragraph(
            "The European industrial components market is moderately fragmented, with the top "
            "15 groups holding approximately 40% of market share. Competition is driven by product "
            "quality, technical support, delivery reliability, and price.", body))
        story.append(PageBreak())

    # Pages 11-14: Intangibles Overview — CRITICAL for royalty direction trap
    story.append(Paragraph("3. Intangibles Overview", h1))
    story.append(Paragraph(
        f"The group's intangible assets are primarily held by {ENTITY_NAMES['CM']} (CM), "
        "the group's R&amp;D centre in Lyon, France. CM is responsible for the Development, "
        "Enhancement, Maintenance, Protection, and Exploitation (DEMPE) of the group's core "
        "technology.", body))
    story.append(Paragraph("DEMPE Analysis", h2))
    dempe = [["DEMPE Function", "Performed By", "Description"],
             ["Development", "CM (France)", "All R&D activities, formulation development, testing, prototyping"],
             ["Enhancement", "CM (France)", "Ongoing improvement of existing products and processes"],
             ["Maintenance", "CM (France)", "Quality control, patent maintenance, regulatory compliance"],
             ["Protection", "CM / CE", "Patent filing and defence (CM), legal oversight (CE)"],
             ["Exploitation", "CM \u2192 CP (license)", "CM licenses technology to CP for manufacturing; CP pays royalty"]]  # noqa: E501
    story.append(_tbl(dempe, [1.5*inch, 1.5*inch, 3.5*inch]))
    story.append(Spacer(1, 0.3*inch))
    story.append(Paragraph(
        f"Under the current transfer pricing arrangement, CP pays CM a royalty of "
        f"{float(ROYALTY_RATE_EU)*100:.0f}% of CP's annual revenue (\u2248\u20ac1.35M) for the "
        "right to use CM's proprietary technology in its manufacturing operations. This royalty "
        "is assessed monthly based on actual revenue.", body))
    story.append(Paragraph(
        "CM is the principal R&amp;D entity within the Cascade Europe group. CM employs "
        "approximately 30 researchers at its Lyon facility and is the legal and economic owner "
        "of all group IP. CM develops the IP and licenses it to CP under the Technology License "
        "Agreement. The royalty flows from CP (licensee) to CM (licensor/IP developer), "
        "consistent with the DEMPE allocation.", body))
    story.append(Paragraph(
        "The royalty rate of 3% is within the range typically observed for manufacturing "
        "technology licenses (2\u20135%) and was established based on a Comparable Uncontrolled "
        "Transaction (CUT) analysis at the time the license was granted.", body))
    story.append(PageBreak())

    for pg in range(12, 15):
        story.append(Paragraph(f"3. Intangibles Overview (continued \u2014 p.{pg})", h1))
        story.append(Paragraph("Key Patents and Know-How", h2))
        story.append(Paragraph(
            "CM holds 14 European patents related to advanced materials formulations, coating "
            "processes, and manufacturing techniques. The patent portfolio was developed through "
            "CM's internal R&amp;D programme, which employs approximately 30 researchers in Lyon.", body))
        story.append(Paragraph(
            "In addition to registered patents, CM possesses significant trade secrets and "
            "manufacturing know-how that are integral to the group's competitive advantage. "
            "These intangibles are protected through confidentiality agreements with all "
            "employees and contractors.", body))
        story.append(PageBreak())

    # Pages 15-18: Intercompany Financial Activities
    story.append(Paragraph("4. Intercompany Financial Activities", h1))
    story.append(Paragraph(
        "The group conducts five principal categories of intercompany transactions, each governed "
        "by formal intercompany agreements and priced in accordance with the arm's-length principle:", body))
    ic = [["Transaction", "Flow", "Method", "FY2024 Volume"],
          ["Raw Materials", "CP \u2192 CM", f"Cost-Plus-{float(RAW_MATERIALS_MARKUP_EU)*100:.0f}%", "~\u20ac8.5M"],
          ["Finished Goods", "CP \u2192 CD", f"Cost-Plus-{float(FINISHED_GOODS_MARKUP_EU)*100:.0f}%", "~\u20ac6.2M"],
          ["Management Fees", "CE \u2192 CP/CM/CD", f"{float(MANAGEMENT_FEE_PCT_EU)*100:.1f}% of revenue", "~\u20ac1.47M"],  # noqa: E501
          ["Intercompany Loan", "CE \u2192 CM", f"{float(IC_LOAN_RATE_EU)*100:.1f}% p.a.", f"\u20ac{int(IC_LOAN_PRINCIPAL_EU):,} principal"],  # noqa: E501
          ["R&D Royalty", "CM \u2192 CP", f"{float(ROYALTY_RATE_EU)*100:.0f}% of CP revenue", "~\u20ac1.35M"]]
    story.append(_tbl(ic, [1.5*inch, 1.5*inch, 1.5*inch, 1.5*inch], alt=True))
    story.append(PageBreak())

    story.append(Paragraph("4. Intercompany Financial Activities (continued)", h1))
    story.append(Paragraph("Intercompany Loan", h2))
    loan_interest = int((IC_LOAN_PRINCIPAL_EU * IC_LOAN_RATE_EU).quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    story.append(Paragraph(
        f"In January 2023, CE extended a \u20ac{int(IC_LOAN_PRINCIPAL_EU):,} intercompany loan to CM "
        f"to fund R&amp;D laboratory expansion at the Lyon facility. The loan bears interest at "
        f"{float(IC_LOAN_RATE_EU)*100:.1f}% per annum, payable monthly. The rate was set by reference "
        "to the 12-month EURIBOR rate plus an appropriate credit spread for BBB-rated European "
        "industrial borrowers.", body))
    loan_d = [["Parameter", "Value"],
              ["Principal", f"\u20ac{int(IC_LOAN_PRINCIPAL_EU):,}"],
              ["Interest Rate", f"{float(IC_LOAN_RATE_EU)*100:.1f}% per annum"],
              ["Benchmark", "EURIBOR 12M + BBB credit spread"],
              ["Term", "10 years"], ["Payment", "Interest-only, monthly accrual"]]
    story.append(_tbl(loan_d, [2*inch, 3.5*inch]))
    story.append(Spacer(1, 0.3*inch))
    story.append(Paragraph(
        f"FY2024 interest income (CE) / expense (CM) totalled \u20ac{loan_interest:,}. "
        "The rate remains within the arm's-length range based on EURIBOR benchmark data.", body))
    story.append(PageBreak())

    for pg in range(17, 19):
        story.append(Paragraph(f"4. Intercompany Financial Activities (continued \u2014 p.{pg})", h1))
        story.append(Paragraph("Management Fee Arrangement", h2))
        story.append(Paragraph(
            f"CE charges each subsidiary a management fee of {float(MANAGEMENT_FEE_PCT_EU)*100:.1f}% of "
            "the subsidiary's annual revenue. The fee covers strategic oversight, treasury and cash "
            "management, legal and regulatory compliance, and IT infrastructure. The OECD Transfer Pricing "
            "Guidelines (Chapter VII) require that management fees satisfy a benefit test \u2014 each "
            "subsidiary must demonstrate that it receives an identifiable, measurable benefit from the "
            "services provided by CE.", body))
        story.append(PageBreak())

    # Pages 19-24: Financial and Tax Positions
    story.append(Paragraph("5. Financial and Tax Positions", h1))
    story.append(Paragraph(
        "The group operates across four tax jurisdictions with the following statutory "
        "corporate income tax rates:", body))
    tax = [["Jurisdiction", "Entity", "Rate", "Notes"],
           [ENTITY_JURISDICTIONS["CE"], "CE", "25.8%", "Holding company regime"],
           [ENTITY_JURISDICTIONS["CP"], "CP", "~29.9%", "K\u00f6rperschaftsteuer 15% + SolZ 5.5% + Gewerbesteuer ~14%"],
           [ENTITY_JURISDICTIONS["CM"], "CM", "25.0%", "Eligible for CIR (30% R&D tax credit)"],
           [ENTITY_JURISDICTIONS["CD"], "CD", "25.0%", "Post-Brexit, TIOPA 2010 Part 4 TP rules"]]
    story.append(_tbl(tax, [1.3*inch, 0.8*inch, 1.2*inch, 3.2*inch]))
    story.append(Spacer(1, 0.3*inch))
    story.append(Paragraph(
        "Country-by-Country Reporting (CbCR): The group's combined revenue of approximately "
        "\u20ac120M is below the \u20ac750M CbCR filing threshold under BEPS Action 13. Accordingly, "
        "CbCR is not required for the Cascade Europe group.", body))
    story.append(PageBreak())

    for pg in range(20, 25):
        story.append(Paragraph(f"5. Financial and Tax Positions (continued \u2014 p.{pg})", h1))
        story.append(Paragraph(
            "The group's transfer pricing arrangements are designed to ensure that profits are "
            "allocated to the jurisdictions where economic value is created, consistent with the "
            "arm's-length principle and BEPS Action 8\u201310 guidance on value creation.", body))
        if pg == 20:
            story.append(Paragraph("Advance Pricing Agreements", h2))
            story.append(Paragraph(
                "The group does not currently have any Advance Pricing Agreements (APAs) in force. "
                "Bilateral APAs between DE/FR and NL/DE are under consideration.", body))
        if pg == 22:
            story.append(Paragraph("Tax Audit History", h2))
            story.append(Paragraph(
                "No material transfer pricing adjustments have been assessed by tax authorities in "
                "any jurisdiction for the past five fiscal years.", body))
        story.append(PageBreak())

    # Pages 25-28: Appendices
    story.append(Paragraph("Appendix A: Entity Legal Details", h1))
    for code in sorted(ENTITY_NAMES):
        story.append(Paragraph(f"<b>{ENTITY_NAMES[code]}</b>", h2))
        story.append(Paragraph(f"Entity code: {code}", body))
        story.append(Paragraph(f"Jurisdiction: {ENTITY_JURISDICTIONS[code]}", body))
        story.append(Paragraph(f"Principal activity: {_ENTITY_ROLES[code]}", body))
        story.append(Spacer(1, 0.2*inch))
    story.append(PageBreak())

    story.append(Paragraph("Appendix A (continued)", h1))
    story.append(Paragraph(
        f"All entities are wholly owned by {ENTITY_NAMES['CE']}. Registered offices and "
        "tax identification numbers are maintained in the group's corporate secretariat records.", body))
    story.append(PageBreak())

    story.append(Paragraph("Appendix B: Intercompany Agreement List", h1))
    agr = [["Agreement", "Parties", "Effective", "Terms"],
           ["Raw Materials Supply Agreement", "CP \u2192 CM", "2019", f"Cost-plus-{float(RAW_MATERIALS_MARKUP_EU)*100:.0f}%"],  # noqa: E501
           ["Finished Goods Distribution Agreement", "CP \u2192 CD", "2023", f"Cost-plus-{float(FINISHED_GOODS_MARKUP_EU)*100:.0f}%"],  # noqa: E501
           ["Group Management Services Agreement", "CE \u2192 All", "2018", f"{float(MANAGEMENT_FEE_PCT_EU)*100:.1f}% of revenue"],  # noqa: E501
           ["Intercompany Loan Agreement", "CE \u2192 CM", "2023", f"\u20ac{int(IC_LOAN_PRINCIPAL_EU):,} at {float(IC_LOAN_RATE_EU)*100:.1f}% p.a."],  # noqa: E501
           ["Technology License Agreement", "CM \u2192 CP", "2019", f"{float(ROYALTY_RATE_EU)*100:.0f}% of CP revenue"]]
    story.append(_tbl(agr, [2.5*inch, 1.2*inch, 1*inch, 1.8*inch]))
    story.append(PageBreak())

    story.append(Paragraph("Appendix B (continued)", h1))
    story.append(Paragraph(
        "All intercompany agreements are reviewed annually by CE's legal and tax teams. "
        "The terms are updated as necessary to reflect changes in market conditions, "
        "regulatory requirements, and group business operations.", body))
    story.append(Spacer(1, 1*inch))
    story.append(Paragraph(
        "<i>End of Master File \u2014 Cascade Europe Holdings B.V. \u2014 FY2024</i>",
        ParagraphStyle("End", parent=body, alignment=1, textColor=colors.grey)))

    doc.build(story)
    _save_pdf_deterministic(buf, file_path)
    canaries.set_location(_KEY_MASTER_FILE, f"{_INPUT_DIR}/master_file_fy2024.pdf", "PDF metadata \u2192 Author")
    manifest.register(f"{_INPUT_DIR}/master_file_fy2024.pdf", "pdf")


# ── Local File CP PDF (35 pages) ─────────────────────────────────────────────


def _write_local_file_cp_pdf(
    output_dir: Path,
    canaries: CanaryRegistry,
    manifest: Manifest,
) -> None:
    """Write local_file_cp_fy2024.pdf — 35-page CP local file."""
    canary = canaries.canary_for(_KEY_LOCAL_FILE_CP)
    file_path = output_dir / _INPUT_DIR / "local_file_cp_fy2024.pdf"
    file_path.parent.mkdir(parents=True, exist_ok=True)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=letter, leftMargin=inch, rightMargin=inch,
        topMargin=inch, bottomMargin=inch, invariant=True)
    doc.title = f"OECD Local File \u2014 {ENTITY_NAMES['CP']} (CP) FY2024"
    doc.author = f"CANARY: {canary}"
    doc.subject = "Transfer Pricing Local File (BEPS Action 13)"
    doc.creator = "Cascade Industries Test Suite Generator"

    title_s, h1, h2, body, small = _pdf_styles()
    story: list[Any] = []

    # Page 1: Title
    story.append(Spacer(1, 2*inch))
    story.append(Paragraph(
        f"OECD Transfer Pricing Local File<br/>{ENTITY_NAMES['CP']}<br/>"
        "Fiscal Year Ended December 31, 2024", title_s))
    story.append(Spacer(1, 0.5*inch))
    story.append(Paragraph(
        "Prepared in accordance with OECD Transfer Pricing Guidelines (2022),<br/>"
        "BEPS Action 13, and German AStG \u00a790 documentation requirements", body))
    story.append(Spacer(1, 0.3*inch))
    story.append(Paragraph(f"Document ID: {canary}", small))
    story.append(PageBreak())

    # Page 2: ToC
    story.append(Paragraph("Table of Contents", h1))
    toc = [("1. Entity Description", "4"), ("2. Functional Analysis", "6"),
           ("3. Controlled Transactions", "13"), ("4. Economic Analysis \u2014 Benchmarking", "21"),
           ("5. Conclusions", "31"), ("Appendix: Financial Data", "34")]
    toc_d = [["Section", "Page"]] + [list(r) for r in toc]
    t = Table(toc_d, colWidths=[5*inch, 1*inch])
    t.setStyle(TableStyle([("FONTSIZE",(0,0),(-1,-1),10), ("TEXTCOLOR",(0,0),(-1,0),_TABLE_HDR_COLOR),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"), ("LINEBELOW",(0,0),(-1,0),1,_TABLE_HDR_COLOR),
        ("ALIGN",(1,0),(1,-1),"RIGHT")]))
    story.append(t)
    story.append(PageBreak())

    # Page 3: Introduction
    story.append(Paragraph("Introduction", h1))
    story.append(Paragraph(
        f"This Local File documents the transfer pricing policies and economic analysis for "
        f"{ENTITY_NAMES['CP']} (\"CP\"), a precision manufacturing entity based in Munich, "
        f"{ENTITY_JURISDICTIONS['CP']}.", body))
    story.append(PageBreak())

    # Pages 4-5: Entity Description
    story.append(Paragraph("1. Entity Description", h1))
    story.append(Paragraph(
        f"{ENTITY_NAMES['CP']} is a wholly-owned subsidiary of {ENTITY_NAMES['CE']} "
        f"({ENTITY_JURISDICTIONS['CE']}). CP operates as a licensed manufacturer, producing "
        f"precision-machined components and assemblies using proprietary technology licensed "
        f"from {ENTITY_NAMES['CM']} (CM, {ENTITY_JURISDICTIONS['CM']}).", body))
    story.append(Paragraph("Key Facts", h2))
    kf = [["Attribute", "Value"],
          ["Legal form", "Gesellschaft mit beschr\u00e4nkter Haftung (GmbH)"],
          ["Registered office", f"Munich, Bavaria, {ENTITY_JURISDICTIONS['CP']}"],
          ["Principal activity", "Licensed manufacturer \u2014 precision components"],
          ["Employees", "~450"], ["FY2024 Revenue", "~\u20ac45M"],
          ["Key relationships", "Receives raw materials from CM; sells finished goods to CD"]]
    story.append(_tbl(kf, [2*inch, 4.5*inch]))
    story.append(PageBreak())

    story.append(Paragraph("1. Entity Description (continued)", h1))
    story.append(Paragraph(
        "CP's manufacturing operations are concentrated at the Munich facility, which includes "
        "CNC machining centres, surface treatment lines, and quality testing laboratories. The "
        "facility operates three shifts and has a capacity utilisation rate of approximately 78%.", body))
    story.append(PageBreak())

    # Pages 6-12: Functional Analysis
    story.append(Paragraph("2. Functional Analysis", h1))
    story.append(Paragraph(
        "CP functions as a licensed manufacturer within the Cascade Europe group. Its principal "
        "functions, risks, and assets are as follows:", body))
    story.append(Paragraph("Functions Performed", h2))
    story.append(Paragraph(
        "&bull; Procurement of raw materials (primarily from CM)<br/>"
        "&bull; Manufacturing of precision components per design specifications<br/>"
        "&bull; Quality control and testing<br/>"
        "&bull; Inventory management<br/>"
        "&bull; Shipping and logistics coordination", body))
    story.append(Paragraph("Risks Assumed", h2))
    story.append(Paragraph(
        "&bull; Production risk (yield, defects)<br/>&bull; Inventory obsolescence risk<br/>"
        "&bull; Raw material price fluctuation risk<br/>"
        "&bull; Limited market risk (primarily sells within the group)", body))
    story.append(Paragraph("Assets Employed", h2))
    story.append(Paragraph(
        "&bull; Tangible: CNC machines, surface treatment equipment, testing labs<br/>"
        "&bull; Intangible: Limited \u2014 manufacturing know-how, but core IP is held by CM "
        "under the Technology License Agreement", body))
    story.append(PageBreak())

    for pg in range(7, 13):
        story.append(Paragraph(f"2. Functional Analysis (continued \u2014 p.{pg})", h1))
        if pg == 7:
            story.append(Paragraph("Characterization", h2))
            story.append(Paragraph(
                "Based on the functional analysis, CP is characterised as a licensed manufacturer "
                "with moderate functional complexity but limited risk and limited ownership of "
                "valuable intangibles. The Transactional Net Margin Method (TNMM) is the most "
                "appropriate method for testing CP's intercompany transactions.", body))
        if pg == 9:
            story.append(Paragraph("R&amp;D Royalty Analysis", h2))
            story.append(Paragraph(
                f"CP pays a royalty of {float(ROYALTY_RATE_EU)*100:.0f}% of its annual revenue to CM "
                "for the right to use CM's proprietary technology. CM is the legal and economic owner "
                "of the IP, performing all DEMPE functions. CP is the licensee. The royalty flows from "
                "CP (licensee) to CM (licensor), which is consistent with the economic substance of "
                "the arrangement.", body))
        story.append(Paragraph(
            "The functional analysis confirms that CP operates within a well-defined role in the "
            "group's value chain. The limited-risk characterisation supports the use of one-sided "
            "transfer pricing methods with CP as the tested party.", body))
        story.append(PageBreak())

    # Pages 13-20: Controlled Transactions
    story.append(Paragraph("3. Controlled Transactions", h1))
    story.append(Paragraph("CP engages in the following controlled transactions:", body))
    ct = [["#", "Transaction", "Counterparty", "Method", "FY2024 Volume"],
          ["1", "Purchase of raw materials", "CM (seller)", f"Cost-Plus-{float(RAW_MATERIALS_MARKUP_EU)*100:.0f}%", "~\u20ac8.5M"],  # noqa: E501
          ["2", "Sale of finished goods", "CD (buyer)", f"Cost-Plus-{float(FINISHED_GOODS_MARKUP_EU)*100:.0f}%", "~\u20ac6.2M"],  # noqa: E501
          ["3", "Management fee", "CE (provider)", f"{float(MANAGEMENT_FEE_PCT_EU)*100:.1f}% of rev", "~\u20ac675K"],
          ["4", "R&D royalty", "CM (licensor)", f"{float(ROYALTY_RATE_EU)*100:.0f}% of rev", "~\u20ac1.35M"]]
    story.append(_tbl(ct, [0.4*inch, 2*inch, 1.3*inch, 1.5*inch, 1.3*inch]))
    story.append(PageBreak())

    for pg in range(14, 21):
        story.append(Paragraph(f"3. Controlled Transactions (continued \u2014 p.{pg})", h1))
        if pg == 14:
            story.append(Paragraph("Raw Materials (CP \u2192 CM)", h2))
            story.append(Paragraph(
                f"CP supplies raw materials and semi-finished components to CM at cost plus a "
                f"{float(RAW_MATERIALS_MARKUP_EU)*100:.0f}% markup. The cost base includes direct "
                "materials, direct labour, and allocated manufacturing overhead.", body))
        elif pg == 16:
            story.append(Paragraph("Finished Goods (CP \u2192 CD)", h2))
            story.append(Paragraph(
                f"CP supplies finished precision components to CD for distribution in the UK market. "
                f"Goods are priced at cost plus {float(FINISHED_GOODS_MARKUP_EU)*100:.0f}%.", body))
        else:
            story.append(Paragraph(
                "Transaction volumes and pricing remain consistent with the group's transfer pricing "
                "policies. All transactions are documented in formal intercompany agreements.", body))
        story.append(PageBreak())

    # Pages 21-30: Economic Analysis
    story.append(Paragraph("4. Economic Analysis \u2014 Benchmarking", h1))
    story.append(Paragraph(
        "The TNMM is applied with operating margin as the profit level indicator (PLI) to test "
        "whether CP's intercompany transactions are at arm's length.", body))
    story.append(Paragraph("Comparable Search Strategy", h2))
    story.append(Paragraph(
        "A search of European manufacturing companies was conducted using the following criteria:<br/>"
        "&bull; SIC code 2899 (industrial chemicals/miscellaneous manufacturing)<br/>"
        "&bull; Revenue range \u20ac50M\u2013\u20ac500M<br/>"
        "&bull; Publicly available financial data<br/>"
        "&bull; No financial distress or restructuring<br/>"
        "&bull; Independent operations (not a captive subsidiary)", body))
    story.append(Paragraph(
        "The initial search identified 15 companies. After screening, 12 were accepted and "
        "3 were rejected.", body))
    story.append(PageBreak())

    for pg in range(22, 31):
        story.append(Paragraph(f"4. Economic Analysis (continued \u2014 p.{pg})", h1))
        if pg == 22:
            story.append(Paragraph("Comparable Screening Results", h2))
            story.append(Paragraph(
                "Three companies were rejected from the comparable set:<br/>"
                "&bull; Nordic Logistics Holdings AB \u2014 SIC code 4731 (freight transportation)<br/>"
                "&bull; Ostrava Heavy Industries a.s. \u2014 financial distress, negative equity<br/>"
                "&bull; Rhineland Grosskonzern AG \u2014 revenue \u20ac4.5B, exceeds 10x CP's revenue "
                "(size outlier per OECD \u00a73.43-3.46)", body))
        elif pg == 25:
            story.append(Paragraph("IQR Computation (FY2024)", h2))
            story.append(Paragraph(
                "Based on the 12 accepted companies, the interquartile range of operating margins "
                "for FY2024 was:<br/>&bull; Q1 (25th percentile): 3.6%<br/>&bull; Median: 5.4%<br/>"
                "&bull; Q3 (75th percentile): 7.6%<br/>"
                "&bull; CP actual operating margin FY2024: ~6.0% (within range)", body))
            story.append(Paragraph(
                "The FY2025 update should use current-year comparable data. The comparable set is "
                "expected to produce a similar IQR.", body))
        else:
            story.append(Paragraph(
                "The benchmarking analysis supports the conclusion that CP's intercompany pricing "
                "is consistent with the arm's-length principle.", body))
        story.append(PageBreak())

    # Pages 31-33: Conclusions
    story.append(Paragraph("5. Conclusions", h1))
    story.append(Paragraph(
        "Based on the economic analysis performed, CP's operating margin of approximately 6.0% "
        "for FY2024 falls within the interquartile range of 3.6%\u20137.6% established from the "
        "comparable company analysis. Accordingly, CP's intercompany transactions are considered "
        "to be at arm's length.", body))
    story.append(Paragraph("Summary by Transaction Type", h2))
    sm = [["Transaction", "Result", "Recommendation"],
          ["Raw materials (CP\u2192CM)", "Within range", "No adjustment required"],
          ["Finished goods (CP\u2192CD)", "Within range", "No adjustment required"],
          ["Management fee", "1.5% within 1-3% range", "No adjustment required"],
          ["R&D royalty", "3% within 2-5% range", "Monitor DEMPE allocation"]]
    story.append(_tbl(sm, [2*inch, 2*inch, 2.5*inch]))
    story.append(PageBreak())

    story.append(Paragraph("5. Conclusions (continued)", h1))
    story.append(Paragraph(
        "The transfer pricing documentation for CP should be updated annually with current year "
        "data. The FY2025 update should include fresh comparable company data and updated IQR "
        "computations.", body))
    story.append(PageBreak())

    # Pages 33-35: Appendix
    story.append(Paragraph("Appendix: Financial Data", h1))
    story.append(Paragraph(f"Selected financial highlights for {ENTITY_NAMES['CP']}:", body))
    fin = [["Metric", "FY2023", "FY2024"],
           ["Revenue", "\u20ac42.5M", "\u20ac45.0M"], ["COGS", "\u20ac30.2M", "\u20ac31.8M"],
           ["Gross margin", "28.9%", "29.3%"], ["Operating expenses", "\u20ac9.8M", "\u20ac10.5M"],
           ["Operating income", "\u20ac2.5M", "\u20ac2.7M"], ["Operating margin", "5.9%", "6.0%"]]
    t = _tbl(fin, [2*inch, 2*inch, 2*inch])
    t.setStyle(TableStyle([("ALIGN", (1,0), (-1,-1), "CENTER")]))
    story.append(t)
    story.append(PageBreak())

    story.append(Paragraph("Appendix (continued)", h1))
    story.append(Paragraph(
        "The financial data presented above is derived from CP's statutory financial statements "
        "prepared under German GAAP (HGB). The operating margin of 6.0% is consistent with the "
        "arm's-length range established in the economic analysis section.", body))
    story.append(Spacer(1, 1*inch))
    story.append(Paragraph(
        f"<i>End of Local File \u2014 {ENTITY_NAMES['CP']} \u2014 FY2024</i>",
        ParagraphStyle("End", parent=body, alignment=1, textColor=colors.grey)))

    doc.build(story)
    _save_pdf_deterministic(buf, file_path)
    canaries.set_location(_KEY_LOCAL_FILE_CP, f"{_INPUT_DIR}/local_file_cp_fy2024.pdf", "PDF metadata \u2192 Author")
    manifest.register(f"{_INPUT_DIR}/local_file_cp_fy2024.pdf", "pdf")


# ── Prompt ───────────────────────────────────────────────────────────────────


def _write_prompt(output_dir: Path) -> None:
    """Write test_cases/TC-09-EU/prompt.md."""
    text = """\
Prepare updated OECD transfer pricing documentation for the Cascade Europe
Holdings B.V. group for FY2025, covering both the master file and entity-level
local files.

Master File (group level):
1. Update the organizational structure section with current FY2025 entity details
   and any changes from FY2024.
2. Update the description of intercompany transactions with FY2025 volumes and
   pricing for all five transaction flows (goods CP\u2192CM, goods CP\u2192CD, management
   fees CE\u2192subs, intercompany loan CE\u2192CM, R&D royalty CM\u2192CP).
3. Update the intangibles section noting that CM (France) is the principal R&D
   entity and licensor of developed IP to CP under the royalty arrangement.
4. Update the financial activities section with the intercompany loan terms and
   note the interest rate relative to EURIBOR benchmarks.

Local File \u2014 Cascade Pr\u00e4zisionsteile GmbH (CP, Germany):
5. Update the functional analysis for CP as a licensed manufacturer receiving
   raw materials from CM and selling finished goods to CD.
6. Screen the manufacturing comparable companies:
   - Reject any that are not appropriate comparables (explain why for each)
   - Compute the interquartile range of operating margins for the accepted set
   - Apply the Transactional Net Margin Method (TNMM) to test CP's margin
7. Determine whether CP's operating margin falls within the arm's-length range.
8. Analyze the R&D royalty received from CM (3% of CP revenue) \u2014 assess whether
   this rate is consistent with arm's-length principles given that CM develops
   the IP and CP is the licensee. Note any concerns.

Local File \u2014 Cascade Distribution Services Ltd (CD, UK):
9. Prepare a functional analysis for CD as a limited-risk distributor.
10. Screen the distribution comparable companies and compute the interquartile
    range of net margins for the accepted set.
11. Apply TNMM to test CD's net margin against the benchmark range.
12. Flag any concerns about CD's margin relative to the distribution benchmark.

Intercompany Loan Analysis:
13. Test the CE\u2192CM intercompany loan interest rate (4.5%) against the EURIBOR
    benchmark data plus appropriate credit spread for BBB-rated industrial
    borrowers. Determine whether the rate is arm's-length.

For all analyses:
14. Flag any transaction types where the actual margin or rate falls outside the
    arm's-length range.
15. Note any data gaps or limitations that affect the conclusions.

Export:
- Benchmarking analysis as Excel (one sheet per tested entity: comparables
  screening, accepted set, IQR computation, tested party margin, conclusion)
- Interest rate benchmark analysis as separate Excel sheet
- Updated master file sections as Word document (organizational structure,
  transactions, intangibles, financial activities)
- Updated CP local file economic analysis section as Word document
- CD local file as Word document (new \u2014 no prior year local file exists for CD)
"""
    path = output_dir / f"test_cases/{_TC}/prompt.md"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text)


# ── Expected Behavior ────────────────────────────────────────────────────────


def _write_expected_behavior(output_dir: Path) -> None:
    """Write test_cases/TC-09-EU/expected_behavior.md."""
    mfg_iqr = compute_mfg_iqr()
    dist_iqr = compute_dist_iqr()

    text = f"""\
# TC-09-EU: OECD Transfer Pricing Documentation \u2014 Expected Behavior

## Manufacturing Comparables Analysis (CP Benchmarking)

- **Total companies**: 15
- **Accepted**: {mfg_iqr['accepted_count']} (after rejecting 3)
- **Rejected**:
  - Nordic Logistics Holdings AB \u2014 SIC code mismatch (4731 vs 2899)
  - Ostrava Heavy Industries a.s. \u2014 Financial distress \u2014 negative equity
  - Rhineland Grosskonzern AG \u2014 Revenue >10x tested party (size outlier per OECD \u00a73.43-3.46)
- **IQR of operating margins**:
  - Q1: {mfg_iqr['q1_pct']}%
  - Median: {mfg_iqr['median_pct']}%
  - Q3: {mfg_iqr['q3_pct']}%
- **CP operating margin**: ~6.2% \u2014 **within range**

## Distribution Comparables Analysis (CD Benchmarking)

- **Total companies**: 10
- **Accepted**: {dist_iqr['accepted_count']} (after rejecting 2)
- **Rejected**:
  - InterGlobal Captive Logistics SA \u2014 Captive entity \u2014 no independent pricing
  - Jupiter Restructuring Services GmbH \u2014 Restructuring losses \u2014 non-recurring
- **IQR of net margins**:
  - Q1: {dist_iqr['q1_pct']}%
  - Median: {dist_iqr['median_pct']}%
  - Q3: {dist_iqr['q3_pct']}%
- **CD net margin**: ~1.8% \u2014 **within range** (low end)

## Interest Rate Benchmarking (CE\u2192CM Loan)

- EURIBOR 12M average FY2025 (correct): ~3.83% ((3.75+3.80+3.85+3.90)/4)
- BBB industrial credit spread average: ~106 bps (avg of 110,105,100,110)
- Arm's-length range: approximately 4.3% to 5.3%
- Actual loan rate: {float(IC_LOAN_RATE_EU)*100:.1f}% \u2014 **within range**
- **ERR-EU-005**: Q3 FY2025 EURIBOR 12M shows 0.38% instead of 3.85% \u2014
  agent must identify this decimal point error

## Royalty Direction Verification

- CM\u2192CP royalty (3% of CP revenue)
- CM is the IP developer/licensor (performs all DEMPE functions)
- CP is the licensee
- Agent must verify that royalty direction is correct: CP pays CM

## Management Fee Benefit Test

- CE charges each subsidiary 1.5% of revenue
- Agent should flag benefit test requirement per OECD Guidelines Chapter VII

## ERR-EU-005 Detection

- EURIBOR 12M rate for Q3 FY2025 entered as 0.38% instead of 3.85%
- Decimal point error that pulls down the average
- If uncorrected, makes the 4.5% loan rate appear above-market

## Expected Deliverables

1. **Benchmarking analysis for CP** (Excel): comparable screening, IQR, margin test
2. **Benchmarking analysis for CD** (Excel): comparable screening, IQR, margin test
3. **Interest rate analysis** (Excel): EURIBOR data, credit spreads, loan assessment
4. **Updated master file** (Word): org structure, transactions, intangibles, financial
5. **Updated CP local file** (Word): economic analysis with FY2025 benchmarking
6. **CD local file \u2014 NEW** (Word): entity overview, functional analysis,
   controlled transactions, economic analysis, financial data appendix
"""
    path = output_dir / f"test_cases/{_TC}/expected_behavior.md"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text)


# ── Gold Standard ────────────────────────────────────────────────────────────


@register_gold("TC-09-EU")
def _tc09_eu_gold(
    canaries: CanaryRegistry,
    errors: ErrorRegistry,
    **model_kwargs: Any,
) -> GoldStandard:
    """TC-09-EU gold standard: OECD transfer pricing documentation."""
    model = model_kwargs["model"]  # noqa: F841
    mfg_iqr = compute_mfg_iqr()
    dist_iqr = compute_dist_iqr()

    return GoldStandard(
        test_case=_TC,
        expected_outputs={
            "output_files": {
                "benchmarking_analysis_cp": {"type": "xlsx", "required_content": [
                    "Manufacturing comparable screening", "IQR computation", "CP margin assessment"]},
                "benchmarking_analysis_cd": {"type": "xlsx", "required_content": [
                    "Distribution comparable screening", "IQR computation", "CD margin assessment"]},
                "interest_rate_analysis": {"type": "xlsx", "required_content": [
                    "EURIBOR rates with data quality flag", "Credit spread data", "Loan rate assessment"]},
                "master_file_update": {"type": "docx", "required_sections": [
                    "Organizational Structure", "Intercompany Transactions", "Intangibles", "Financial Activities"]},
                "cp_local_file_update": {"type": "docx", "required_sections": [
                    "Economic Analysis \u2014 FY2025 Benchmarking Results"]},
                "cd_local_file_new": {"type": "docx", "required_sections": [
                    "Entity Overview", "Functional Analysis", "Controlled Transactions",
                    "Economic Analysis", "Financial Data Appendix"]},
            },
            "manufacturing_comparable_screening": {
                "total_companies": 15,
                "accepted": mfg_iqr["accepted_count"],
                "rejected": mfg_iqr["rejected_count"],
                "rejections": [
                    {"company": "Nordic Logistics Holdings AB", "reason": "SIC code mismatch (4731 vs 2899)"},
                    {"company": "Ostrava Heavy Industries a.s.", "reason": "Financial distress \u2014 negative equity"},
                    {"company": "Rhineland Grosskonzern AG",
                     "reason": "Revenue >10x tested party (size outlier per OECD \u00a73.43-3.46)"},
                ],
            },
            "distribution_comparable_screening": {
                "total_companies": 10,
                "accepted": dist_iqr["accepted_count"],
                "rejected": dist_iqr["rejected_count"],
                "rejections": [
                    {"company": "InterGlobal Captive Logistics SA", "reason": "Captive entity \u2014 no independent pricing"},  # noqa: E501
                    {"company": "Jupiter Restructuring Services GmbH", "reason": "Restructuring losses \u2014 non-recurring"},  # noqa: E501
                ],
            },
            "mfg_iqr_analysis": {
                "q1_pct": mfg_iqr["q1_pct"],
                "median_pct": mfg_iqr["median_pct"],
                "q3_pct": mfg_iqr["q3_pct"],
            },
            "dist_iqr_analysis": {
                "q1_pct": dist_iqr["q1_pct"],
                "median_pct": dist_iqr["median_pct"],
                "q3_pct": dist_iqr["q3_pct"],
            },
            "arm_length_assessment": {
                "cp_operating_margin_pct": 6.2,
                "cp_within_iqr": True,
                "cd_net_margin_pct": 1.8,
                "cd_within_iqr": True,
                "loan_rate_pct": str(float(IC_LOAN_RATE_EU) * 100),
                "loan_within_range": True,
                "loan_range": "4.3%-5.3%",
                "royalty_rate_pct": str(float(ROYALTY_RATE_EU) * 100),
                "royalty_direction": "CP pays CM (correct \u2014 CM is IP developer/licensor)",
                "management_fee_pct": str(float(MANAGEMENT_FEE_PCT_EU) * 100),
                "management_fee_benefit_test_required": True,
            },
            "interest_rate_benchmarking": {
                "euribor_12m_avg_fy2025_correct_pct": 3.83,
                "bbb_spread_avg_bps": 106,
                "arm_length_range_pct": "4.3%-5.3%",
                "actual_rate_pct": str(float(IC_LOAN_RATE_EU) * 100),
                "conclusion": "Within arm's-length range",
            },
        },
        canary_verification={
            "read_ic_transactions": canaries.canary_for(_KEY_IC_TRANSACTIONS),
            "read_comparable_companies": canaries.canary_for(_KEY_COMPARABLE_COMPANIES),
            "read_master_file": canaries.canary_for(_KEY_MASTER_FILE),
            "read_local_file_cp": canaries.canary_for(_KEY_LOCAL_FILE_CP),
            "read_interest_benchmarks": canaries.canary_for(_KEY_INTEREST_BENCHMARKS),
        },
        error_detection={
            "ERR-EU-005": (
                "EURIBOR 12M Q3 FY2025 decimal point error: 0.38% instead of 3.85%. "
                "Makes loan rate 4.5% appear above-market when it is actually within range."
            ),
        },
        scoring_hints={
            "correctness": (
                f"Manufacturing IQR must be Q1={mfg_iqr['q1_pct']}%, Q3={mfg_iqr['q3_pct']}%. "
                f"Distribution IQR must be Q1={dist_iqr['q1_pct']}%, Q3={dist_iqr['q3_pct']}%. "
                "CP margin ~6.2% within range. CD margin ~1.8% within range. "
                "Loan rate 4.5% within 4.3%-5.3% range. ERR-EU-005 must be identified."
            ),
            "completeness": (
                "All five transaction types analysed with appropriate methods. "
                "Master file AND local files produced (BEPS Action 13 split). "
                "CD local file created from scratch. "
                "5 rejections identified (3 manufacturing, 2 distribution). "
                "DEMPE analysis referenced for royalty. Benefit test noted for management fees."
            ),
            "format_compliance": (
                "Benchmarking as Excel with per-entity sheets. Interest rate analysis as separate sheet. "
                "Master file and local files as Word documents. OECD two-tier documentation structure."
            ),
            "robustness": (
                "Agent must reject size outlier (OECD-specific screen). "
                "Agent must detect ERR-EU-005 decimal point error. "
                "Agent must verify royalty direction against DEMPE. "
                "Agent must note CD is non-EU (post-Brexit). "
                "Agent must note CbCR not required (below threshold)."
            ),
            "communication": (
                "Clear arm's-length conclusions per transaction type. "
                "Professional OECD transfer pricing terminology. "
                "Actionable recommendations where margins are borderline."
            ),
        },
        scenario_pack="cascade_europe_ifrs",
    )


# ── Public entry point ───────────────────────────────────────────────────────


def emit_tc09_eu(
    model: CascadeModel,
    output_dir: Path,
    canaries: CanaryRegistry,
    errors: ErrorRegistry,
    manifest: Manifest,
    **kwargs: object,
) -> None:
    """Write all TC-09-EU files to *output_dir*."""
    _write_ic_transactions_xlsx(output_dir, canaries, manifest)
    _write_comparable_companies_xlsx(output_dir, canaries, manifest)
    _write_interest_rate_benchmarks_xlsx(output_dir, canaries, errors, manifest)
    _write_master_file_pdf(output_dir, canaries, manifest)
    _write_local_file_cp_pdf(output_dir, canaries, manifest)
    _write_prompt(output_dir)
    _write_expected_behavior(output_dir)
