"""Formatter: TC-10-EU — VAT and Cross-Border Tax Position Analysis.

Emits:
- test_cases/TC-10-EU/input_files/tc10eu_intercompany_sales_fy2025.xlsx
  Intercompany and third-party sales with VAT treatment per transaction
- test_cases/TC-10-EU/input_files/tc10eu_vat_registrations.xlsx
  VAT registration details per entity (includes pending PL registration trap)
- test_cases/TC-10-EU/input_files/tc10eu_vat_returns_summary_fy2025.xlsx
  Quarterly VAT return summaries per entity
- test_cases/TC-10-EU/input_files/tc10eu_eu_vat_rules_reference.docx
  Reference document on EU VAT rules (intra-EU supply, reverse charge,
  post-Brexit, triangulation, call-off stock, VAT PE)
- test_cases/TC-10-EU/prompt.md
- test_cases/TC-10-EU/expected_behavior.md
- gold_standards/TC-10-EU_gold.json

Planted error:
  ERR-EU-010: classification_error — CE invoiced CM for Q3 management fees
  with 20% French VAT. CE is not FR-registered; reverse charge applies.

Uses deterministic European VAT model — never hardcodes numbers that should
come from the model.
"""

from __future__ import annotations

import datetime
import io
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

import docx
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from generator.canaries import (
    CanaryRegistry,
    embed_canary_docx,
    embed_canary_xlsx,
)
from generator.errors import ErrorRegistry, PlantedError
from generator.golds.framework import GoldStandard, register_gold
from generator.manifest import Manifest
from generator.model.build import CascadeModel
from generator.model.vat_eu import (
    CM_ITALIAN_SALES,
    CP_TO_CD_Q2_UNDOCUMENTED,
    ERR_EU_010_AMOUNT,
    ERR_EU_010_QUARTER,
    ERR_EU_010_VAT_CHARGED,
    VAT_RATES,
    generate_ic_sales_eu,
    generate_vat_registrations,
    generate_vat_returns,
)

# ── Constants ────────────────────────────────────────────────────────────────

_TC = "TC-10-EU"
_INPUT_DIR = f"test_cases/{_TC}/input_files"

_FIXED_DATETIME = datetime.datetime(2025, 3, 15, 9, 0, 0)
_FIXED_ZIP_DT = (2025, 3, 15, 9, 0, 0)

# Canary keys
_KEY_IC_SALES = "tc10eu_intercompany_sales_fy2025"
_KEY_VAT_REGS = "tc10eu_vat_registrations"
_KEY_VAT_RETURNS = "tc10eu_vat_returns_summary_fy2025"
_KEY_VAT_RULES = "tc10eu_eu_vat_rules_reference"

# Styles
_HEADER_FILL = PatternFill("solid", fgColor="2E5090")
_HEADER_FONT = Font(bold=True, size=11, color="FFFFFF")
_HEADER_ALIGN = Alignment(horizontal="center", wrap_text=True)
_NUM_FMT = "#,##0"
_THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


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


def _save_docx_deterministic(document: Any, path: str | Path) -> None:
    """Save python-docx Document with fixed zip entry timestamps."""
    path = Path(path)
    buf = io.BytesIO()
    document.save(buf)

    buf.seek(0)
    with ZipFile(buf, "r") as src, ZipFile(str(path), "w", ZIP_DEFLATED) as dst:
        for item in src.infolist():
            info = ZipInfo(item.filename, date_time=_FIXED_ZIP_DT)
            info.compress_type = item.compress_type
            dst.writestr(info, src.read(item.filename))


def _style_header(ws: Any, row: int, col_count: int) -> None:
    for col in range(1, col_count + 1):
        cell = ws.cell(row=row, column=col)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _HEADER_ALIGN
        cell.border = _THIN_BORDER


def _style_data_cell(cell: Any, fmt: str = "") -> None:
    cell.border = _THIN_BORDER
    if fmt:
        cell.number_format = fmt


# ── Intercompany Sales XLSX ──────────────────────────────────────────────────

def _write_ic_sales_xlsx(
    output_dir: Path,
    canaries: CanaryRegistry,
    errors: ErrorRegistry,
    manifest: Manifest,
) -> None:
    """Write tc10eu_intercompany_sales_fy2025.xlsx."""
    sales = generate_ic_sales_eu()

    wb = openpyxl.Workbook()
    canary = canaries.canary_for(_KEY_IC_SALES)
    loc = embed_canary_xlsx(wb, canary)

    # ── Sheet 1: Intercompany Sales ──
    ws = wb.active
    ws.title = "Intercompany Sales"

    headers = [
        "Seller Entity", "Buyer Entity", "Description", "Amount (EUR)",
        "VAT Treatment Applied", "Invoice VAT Rate", "Incoterms",
        "Proof of Dispatch", "Quarter",
    ]
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)
    _style_header(ws, 1, len(headers))

    err_row = None
    for i, sale in enumerate(sales, 2):
        ws.cell(row=i, column=1, value=sale.seller)
        ws.cell(row=i, column=2, value=sale.buyer)
        ws.cell(row=i, column=3, value=sale.description)
        ws.cell(row=i, column=4, value=float(sale.amount_eur))
        ws.cell(row=i, column=5, value=sale.vat_treatment)
        ws.cell(row=i, column=6, value=sale.invoice_vat_rate)
        ws.cell(row=i, column=7, value=sale.incoterms)
        ws.cell(row=i, column=8, value=sale.proof_of_dispatch)
        ws.cell(row=i, column=9, value=sale.quarter)

        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=i, column=col)
            if col == 4:
                _style_data_cell(cell, _NUM_FMT)
            else:
                _style_data_cell(cell)

        if sale.is_error:
            err_row = i

    # Column widths
    widths = [12, 12, 55, 16, 42, 14, 14, 16, 8]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w

    # ── Sheet 2: Third-Party Sales Summary ──
    ws2 = wb.create_sheet("Third-Party Sales Summary")
    tp_headers = [
        "Entity", "Domestic Sales (EUR)", "EU Sales (EUR)",
        "Non-EU Sales (EUR)", "Total Third-Party (EUR)", "Notes",
    ]
    for col, h in enumerate(tp_headers, 1):
        ws2.cell(row=1, column=col, value=h)
    _style_header(ws2, 1, len(tp_headers))

    tp_data = [
        ("CE", 500000, 0, 0, 500000, "NL domestic advisory services only"),
        ("CP", 18000000, 5500000, 2000000, 25500000,
         "Domestic 40%, EU 22%, Non-EU 8% of €45M revenue (rest is IC)"),
        ("CM", 12000000, 6000000, 1600000, 19600000,
         "Domestic 38%, EU 19% (includes €380k Italy), Non-EU 5%"),
        ("CD", 16800000, 0, 0, 16800000,
         "UK domestic only — CD does not export"),
    ]
    for i, (ent, dom, eu, non_eu, total, notes) in enumerate(tp_data, 2):
        ws2.cell(row=i, column=1, value=ent)
        ws2.cell(row=i, column=2, value=dom)
        ws2.cell(row=i, column=3, value=eu)
        ws2.cell(row=i, column=4, value=non_eu)
        ws2.cell(row=i, column=5, value=total)
        ws2.cell(row=i, column=6, value=notes)
        for col in range(1, len(tp_headers) + 1):
            cell = ws2.cell(row=i, column=col)
            if col in (2, 3, 4, 5):
                _style_data_cell(cell, _NUM_FMT)
            else:
                _style_data_cell(cell)

    for col, w in enumerate([10, 20, 16, 16, 20, 55], 1):
        ws2.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w

    # Register planted error
    errors.add(PlantedError(
        error_id="ERR-EU-010",
        file=f"{_INPUT_DIR}/tc10eu_intercompany_sales_fy2025.xlsx",
        location=(
            f"Sheet 'Intercompany Sales', Row {err_row}: "
            f"CE→CM management fee {ERR_EU_010_QUARTER} FY2025"
        ),
        type="classification_error",
        description=(
            f"CE (Netherlands) invoiced CM (France) for {ERR_EU_010_QUARTER} "
            f"management fees (€{ERR_EU_010_AMOUNT:,}) with 20% French VAT "
            f"(€{ERR_EU_010_VAT_CHARGED:,}). CE is not VAT-registered in "
            "France; reverse charge under Art. 196 VAT Directive applies. "
            "CM should self-assess French output VAT. Requires credit note."
        ),
        severity="material",
        which_test_cases_should_catch=["TC-10-EU"],
    ))

    rel_path = f"{_INPUT_DIR}/tc10eu_intercompany_sales_fy2025.xlsx"
    abs_path = output_dir / rel_path
    abs_path.parent.mkdir(parents=True, exist_ok=True)
    _save_xlsx_deterministic(wb, abs_path)

    canaries.set_location(_KEY_IC_SALES, rel_path, loc)
    manifest.register(rel_path, "xlsx")


# ── VAT Registrations XLSX ───────────────────────────────────────────────────

def _write_vat_registrations_xlsx(
    output_dir: Path,
    canaries: CanaryRegistry,
    manifest: Manifest,
) -> None:
    """Write tc10eu_vat_registrations.xlsx."""
    regs = generate_vat_registrations()

    wb = openpyxl.Workbook()
    canary = canaries.canary_for(_KEY_VAT_REGS)
    loc = embed_canary_xlsx(wb, canary)

    ws = wb.active
    ws.title = "VAT Registrations"

    headers = [
        "Entity Code", "Country", "VAT ID", "Registration Date",
        "VAT Group (Y/N)", "Fiscal Representative",
        "EC Sales List Filed (Y/N)", "Intrastat Threshold Exceeded (Y/N)",
        "Status",
    ]
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)
    _style_header(ws, 1, len(headers))

    for i, reg in enumerate(regs, 2):
        ws.cell(row=i, column=1, value=reg.entity_code)
        ws.cell(row=i, column=2, value=reg.country)
        ws.cell(row=i, column=3, value=reg.vat_id)
        ws.cell(row=i, column=4, value=reg.registration_date)
        ws.cell(row=i, column=5, value=reg.vat_group)
        ws.cell(row=i, column=6, value=reg.fiscal_representative or "—")
        ws.cell(row=i, column=7, value=reg.ecsl_filed)
        ws.cell(row=i, column=8, value=reg.intrastat_exceeded)
        ws.cell(row=i, column=9, value=reg.status)

        for col in range(1, len(headers) + 1):
            _style_data_cell(ws.cell(row=i, column=col))

    widths = [12, 16, 18, 16, 14, 20, 22, 30, 28]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w

    rel_path = f"{_INPUT_DIR}/tc10eu_vat_registrations.xlsx"
    abs_path = output_dir / rel_path
    abs_path.parent.mkdir(parents=True, exist_ok=True)
    _save_xlsx_deterministic(wb, abs_path)

    canaries.set_location(_KEY_VAT_REGS, rel_path, loc)
    manifest.register(rel_path, "xlsx")


# ── VAT Returns Summary XLSX ────────────────────────────────────────────────

def _write_vat_returns_xlsx(
    output_dir: Path,
    canaries: CanaryRegistry,
    manifest: Manifest,
) -> None:
    """Write tc10eu_vat_returns_summary_fy2025.xlsx."""
    returns = generate_vat_returns()

    wb = openpyxl.Workbook()
    canary = canaries.canary_for(_KEY_VAT_RETURNS)
    loc = embed_canary_xlsx(wb, canary)

    ws = wb.active
    ws.title = "VAT Returns Summary"

    headers = [
        "Entity", "Quarter", "Output VAT Domestic (EUR)",
        "Output VAT Intra-EU/Export 0% (EUR)",
        "Input VAT Domestic (EUR)", "Input VAT Reverse Charge (EUR)",
        "VAT Payable / (Refundable) (EUR)", "Filing Status",
    ]
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)
    _style_header(ws, 1, len(headers))

    for i, ret in enumerate(returns, 2):
        ws.cell(row=i, column=1, value=ret.entity_code)
        ws.cell(row=i, column=2, value=ret.quarter)
        ws.cell(row=i, column=3, value=float(ret.output_vat_domestic))
        ws.cell(row=i, column=4, value=float(ret.output_vat_intra_eu_export))
        ws.cell(row=i, column=5, value=float(ret.input_vat_domestic))
        ws.cell(row=i, column=6, value=float(ret.input_vat_reverse_charge))
        ws.cell(row=i, column=7, value=float(ret.vat_payable))
        ws.cell(row=i, column=8, value=ret.filing_status)

        for col in range(1, len(headers) + 1):
            cell = ws.cell(row=i, column=col)
            if col in (3, 4, 5, 6, 7):
                _style_data_cell(cell, _NUM_FMT)
            else:
                _style_data_cell(cell)

    widths = [10, 10, 22, 28, 22, 26, 28, 26]
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = w

    rel_path = f"{_INPUT_DIR}/tc10eu_vat_returns_summary_fy2025.xlsx"
    abs_path = output_dir / rel_path
    abs_path.parent.mkdir(parents=True, exist_ok=True)
    _save_xlsx_deterministic(wb, abs_path)

    canaries.set_location(_KEY_VAT_RETURNS, rel_path, loc)
    manifest.register(rel_path, "xlsx")


# ── EU VAT Rules Reference DOCX ─────────────────────────────────────────────

def _write_vat_rules_docx(
    output_dir: Path,
    canaries: CanaryRegistry,
    manifest: Manifest,
) -> None:
    """Write tc10eu_eu_vat_rules_reference.docx."""
    canary = canaries.canary_for(_KEY_VAT_RULES)

    doc = docx.Document()
    loc = embed_canary_docx(doc, canary)

    doc.add_heading("EU VAT Rules Reference — Cascade Europe Group", level=1)
    doc.add_paragraph(
        "This document summarizes the key EU VAT rules applicable to the "
        "Cascade Europe group's cross-border transactions. It is intended as "
        "a working reference for the FY2025 VAT compliance review."
    )

    # Section 1: Intra-Community supply of goods
    doc.add_heading("1. Intra-Community Supply of Goods (Article 138 VAT Directive)", level=2)
    doc.add_paragraph(
        "An intra-Community supply of goods is zero-rated (0% VAT) when the "
        "following conditions are met:"
    )
    conditions = [
        "The goods are dispatched or transported from one EU member state to another;",
        "The supplier has the acquirer's valid VAT identification number;",
        "The supplier has documentary proof that the goods were transported "
        "(proof of dispatch: CMR, bill of lading, or equivalent);",
        "The acquirer reports the intra-Community acquisition in their member state.",
    ]
    for c in conditions:
        doc.add_paragraph(c, style="List Bullet")
    doc.add_paragraph(
        "If any condition is not met, the zero rate cannot be applied and the "
        "supplier must charge domestic VAT. In particular, missing or incomplete "
        "proof of dispatch means the supply may be reassessed at the domestic "
        "rate on audit."
    )

    # Section 2: Reverse charge on cross-border services
    doc.add_heading("2. Reverse Charge on Cross-Border Services (Article 196 VAT Directive)", level=2)
    doc.add_paragraph(
        "For B2B (business-to-business) cross-border services, the general rule "
        "is that the service is taxed where the customer is established (Article "
        "44 VAT Directive). Under the reverse charge mechanism (Article 196), "
        "the supplier does not charge VAT; instead, the customer self-assesses "
        "VAT in their own member state."
    )
    doc.add_paragraph(
        "The reverse charge applies to management services, consultancy, "
        "royalties, and similar B2B services supplied between entities in "
        "different member states. The supplier issues an invoice without VAT, "
        "noting 'Reverse charge — Article 196 VAT Directive' on the invoice."
    )
    doc.add_paragraph(
        "Important: A supplier that is not established or VAT-registered in "
        "the customer's member state should never charge the customer's local "
        "VAT. Doing so creates an invalid invoice and compliance issues for "
        "both parties."
    )

    # Section 3: UK post-Brexit
    doc.add_heading("3. UK Post-Brexit Treatment", level=2)
    doc.add_paragraph(
        "Since 1 January 2021, the United Kingdom is a third country for EU "
        "VAT purposes. Key implications:"
    )
    post_brexit = [
        "Exports from an EU member state to the UK are zero-rated as exports "
        "to a third country (not intra-Community supplies). Customs declarations "
        "are required.",
        "Imports from the UK into the EU are subject to import VAT at the "
        "relevant member state rate. Postponed VAT accounting may be available.",
        "Intra-Community acquisition rules no longer apply to UK transactions.",
        "B2B services between EU and UK entities follow the general place of "
        "supply rules. UK reverse charge applies under UK domestic legislation.",
        "EC Sales Lists and Intrastat declarations do not apply to UK transactions.",
    ]
    for p in post_brexit:
        doc.add_paragraph(p, style="List Bullet")

    # Section 4: Triangulation simplification
    doc.add_heading("4. Triangulation Simplification (Article 141 VAT Directive)", level=2)
    doc.add_paragraph(
        "In a chain transaction where goods move directly from member state A "
        "to member state C, but the invoicing chain goes A → B → C, the "
        "intermediate party B may use the triangulation simplification to avoid "
        "registering for VAT in member state C. Conditions: all three parties "
        "must be registered for VAT in different member states, the goods must "
        "move directly from A to C, and the intermediate party must note "
        "'triangulation — Article 141' on its invoice."
    )

    # Section 5: Call-off stock
    doc.add_heading("5. Call-Off Stock Arrangements (Article 17a VAT Directive)", level=2)
    doc.add_paragraph(
        "Goods transferred to another member state for a known and identified "
        "customer may benefit from the call-off stock simplification. The "
        "supplier does not need to register for VAT in the destination member "
        "state, provided the goods are supplied to the identified customer "
        "within 12 months and certain record-keeping conditions are met."
    )

    # Section 6: VAT PE
    doc.add_heading("6. Permanent Establishment for VAT Purposes", level=2)
    doc.add_paragraph(
        "A fixed establishment for VAT purposes is defined differently from "
        "a permanent establishment for direct tax (income tax) purposes. For "
        "VAT, a fixed establishment requires 'a sufficient degree of permanence "
        "and a suitable structure in terms of human and technical resources to "
        "enable it to receive and use or to make supplies' (CJEU, Welmory, "
        "C-605/12)."
    )
    doc.add_paragraph(
        "IMPORTANT: A VAT fixed establishment is NOT the same as a corporate "
        "tax PE under Art. 5 OECD Model Tax Convention. The two concepts have "
        "different legal tests and different consequences. An entity may have "
        "an income tax PE without a VAT fixed establishment, or vice versa. "
        "Do not apply OECD Model Convention criteria when assessing VAT PE risk."
    )
    doc.add_paragraph(
        "For the Cascade Europe group, the primary VAT PE risk arises from "
        "CE's management service delivery: if CE personnel regularly travel to "
        "subsidiary locations and use fixed resources there to deliver services, "
        "CE could be found to have a fixed establishment in that member state."
    )

    rel_path = f"{_INPUT_DIR}/tc10eu_eu_vat_rules_reference.docx"
    abs_path = output_dir / rel_path
    abs_path.parent.mkdir(parents=True, exist_ok=True)
    _save_docx_deterministic(doc, abs_path)

    canaries.set_location(_KEY_VAT_RULES, rel_path, loc)
    manifest.register(rel_path, "docx")


# ── Prompt markdown ──────────────────────────────────────────────────────────

def _write_prompt(output_dir: Path) -> None:
    """Write test_cases/TC-10-EU/prompt.md."""
    path = output_dir / f"test_cases/{_TC}/prompt.md"
    path.parent.mkdir(parents=True, exist_ok=True)
    text = """\
# TC-10-EU: VAT and Cross-Border Tax Position Analysis

## Context

Cascade Europe Holdings B.V. is a Netherlands-based holding company with three
operating subsidiaries: Cascade Präzisionsteile GmbH (CP, Germany — precision
manufacturing), Cascade Matériaux Avancés SAS (CM, France — R&D and advanced
materials), and Cascade Distribution Services Ltd (CD, UK — distribution).

You are analyzing the group's VAT and cross-border indirect tax position for
FY2025 across four jurisdictions (NL, DE, FR, UK) under EU VAT Directive rules
and UK post-Brexit VAT legislation.

## Input Files

1. **tc10eu_intercompany_sales_fy2025.xlsx** — Intercompany and third-party
   sales data with VAT treatment per transaction.

2. **tc10eu_vat_registrations.xlsx** — VAT registration details per entity
   including VAT IDs, filing status, and registration dates.

3. **tc10eu_vat_returns_summary_fy2025.xlsx** — Quarterly VAT return summaries
   showing output VAT, input VAT, reverse charge amounts, and filing status.

4. **tc10eu_eu_vat_rules_reference.docx** — Reference document on EU VAT rules
   covering intra-Community supply, reverse charge, UK post-Brexit treatment,
   triangulation, call-off stock, and VAT PE.

## Instructions

Analyze the VAT and cross-border tax position for the Cascade Europe Holdings
B.V. group for FY2025.

1. For each intercompany transaction flow, determine the correct VAT treatment:
   - Is the supply of goods or services?
   - Where is the place of supply under EU VAT rules?
   - Does the reverse charge mechanism apply?
   - For goods: is zero-rating available, and are the conditions met
     (valid VAT ID, proof of dispatch)?
2. Review the intercompany sales data for any VAT treatment errors:
   - Are there transactions where local VAT was incorrectly charged?
   - Are there transactions where zero-rating was applied but conditions
     are not fully met?
3. Assess the group's VAT registration position:
   - Are there jurisdictions where an entity may have an unregistered
     VAT obligation?
   - Flag any pending registrations and assess the risk.
4. Review the quarterly VAT return summaries:
   - Do the returns reconcile to the transaction data?
   - Are there unusual positions (large refunds, pending assessments)
     that warrant investigation?
5. Assess permanent establishment risk:
   - Based on the intercompany transaction flows and entity activities,
     could any entity have created a VAT fixed establishment in another
     jurisdiction?
   - Note: VAT fixed establishment criteria differ from income tax PE.
6. For the UK subsidiary (post-Brexit):
   - Confirm the correct treatment of goods flows between EU entities
     and the UK entity.
   - Are customs declarations in place for CP→CD shipments?
   - Is import VAT being accounted for correctly (postponed accounting)?
7. Summarize the group's cross-border indirect tax risk profile:
   - Identified errors or incorrect treatments
   - Missing data or documentation gaps
   - Areas requiring further investigation
   - Recommendations for compliance improvements

## Export

Export as an Excel workbook with sheets for:
- Transaction-by-Transaction VAT Analysis
- VAT Registration Assessment
- Quarterly VAT Reconciliation
- PE Risk Assessment
- Risk Summary and Recommendations
"""
    path.write_text(text)


# ── Expected behavior markdown ───────────────────────────────────────────────

def _write_expected_behavior(output_dir: Path) -> None:
    """Write test_cases/TC-10-EU/expected_behavior.md."""
    path = output_dir / f"test_cases/{_TC}/expected_behavior.md"
    path.parent.mkdir(parents=True, exist_ok=True)

    vat_exposure = float(CP_TO_CD_Q2_UNDOCUMENTED * VAT_RATES["DE"])
    text = f"""\
# TC-10-EU: Expected Behavior

## Key Expectations

### Transaction-by-Transaction VAT Classification
- Agent must analyze each intercompany flow individually.
- Applying a single rule to all transactions or factor-based allocation is wrong.

### Reverse Charge Mechanism
- CE→CP management fees: Reverse charge (Art. 196)
- CE→CM management fees: Reverse charge (Art. 196) — except ERR-EU-010
- CE→CD management fees: Outside EU VAT scope (UK reverse charge)
- CM→CP R&D royalty: Reverse charge (Art. 196)

### Intra-EU Supply Conditions
- CP→CM raw materials: 0% with proof of dispatch — conditions met
- CP→CD finished goods: Export to third country — Q2 has partial proof of
  dispatch (2/5 shipments missing). Potential €{vat_exposure:,.0f} VAT exposure.

### UK Post-Brexit Treatment
- UK is a third country. CP→CD shipments are exports, not intra-Community.
- CD uses postponed accounting for import VAT.

### ERR-EU-010 Detection
- CE charged 20% French VAT on Q3 management fees to CM.
- CE is NL-established, not FR-registered — reverse charge applies.
- Requires credit note.

### Missing Data Identification (at least 3)
1. Q2 proof of dispatch gap for CP→CD shipments
2. Unexplained Polish VAT registration for CP
3. CM's €380k Italian sales without Italian VAT registration

### VAT PE vs Income Tax PE
- Different legal tests — do not conflate.
- Agent should not use OECD Art. 5 or Pillar Two thresholds for VAT PE.

### CE Q4 Pending Assessment
- Unusual for a holding company — flag for monitoring.

### No Factor-Based Allocation
- VAT is a transaction tax. Apportioning income across jurisdictions is wrong.

## Error Detection
- ERR-EU-010: CE incorrectly charged 20% French VAT on management fees to CM
  instead of applying reverse charge under Art. 196 VAT Directive.
"""
    path.write_text(text)


# ── Gold standard ────────────────────────────────────────────────────────────

@register_gold(_TC)
def _tc10_eu_gold(
    canaries: CanaryRegistry,
    errors: ErrorRegistry,
    **model_kwargs: Any,
) -> GoldStandard:
    """TC-10-EU gold standard: VAT and cross-border tax analysis."""
    vat_exposure = float(CP_TO_CD_Q2_UNDOCUMENTED * VAT_RATES["DE"])

    expected_outputs: dict[str, Any] = {
        "output_files": {
            "vat_analysis_workbook": {
                "type": "xlsx",
                "required_sheets": [
                    "Transaction-by-Transaction VAT Analysis",
                    "VAT Registration Assessment",
                    "Quarterly VAT Reconciliation",
                    "PE Risk Assessment",
                    "Risk Summary and Recommendations",
                ],
            },
        },
        "transaction_vat_analysis": {
            "cp_to_cm_raw_materials": {
                "treatment": "Intra-EU supply of goods, 0% (Art. 138)",
                "conditions_met": True,
            },
            "cp_to_cd_finished_goods": {
                "treatment": "Export to third country (post-Brexit), 0%",
                "q2_proof_of_dispatch": "Partial (2/5 missing)",
                "vat_exposure_eur": vat_exposure,
            },
            "ce_to_cp_mgmt_fee": {
                "treatment": "Reverse charge (Art. 196, B2B services)",
                "correct": True,
            },
            "ce_to_cm_mgmt_fee": {
                "treatment": "Should be reverse charge (Art. 196)",
                "error": "ERR-EU-010: CE charged 20% French VAT",
                "requires_credit_note": True,
            },
            "ce_to_cd_mgmt_fee": {
                "treatment": "Outside EU VAT scope (UK reverse charge)",
                "correct": True,
            },
            "cm_to_cp_royalty": {
                "treatment": "Reverse charge (Art. 196, B2B services)",
                "correct": True,
            },
        },
        "registration_assessment": {
            "ce_nl": "Compliant",
            "cp_de": "Compliant \u2014 but pending PL registration unexplained",
            "cm_fr": "Potential issue \u2014 IT sales without IT registration",
            "cd_uk": "Compliant \u2014 UK VAT, postponed accounting",
        },
        "missing_data_gaps": [
            "Q2 proof of dispatch for CP\u2192CD (2/5 shipments)",
            "Polish VAT registration for CP \u2014 no transactions to Poland",
            "CM Italian sales (\u20ac380k) without Italian VAT registration",
        ],
        "pe_risk": {
            "cp_in_fr": "Low \u2014 goods shipped, not installed",
            "cm_in_de": "Low \u2014 IP licensed, no personnel in DE",
            "ce_in_subs": "Medium \u2014 depends on travel patterns (data gap)",
            "cd_in_eu": "Low \u2014 UK-only operations",
        },
        "quantified_risks": {
            "q2_dispatch_exposure_eur": vat_exposure,
            "err_eu_010_vat_amount_eur": float(ERR_EU_010_VAT_CHARGED),
            "cm_italian_sales_eur": float(CM_ITALIAN_SALES),
        },
    }

    canary_verification = {
        "read_ic_sales": canaries.canary_for(_KEY_IC_SALES),
        "read_vat_registrations": canaries.canary_for(_KEY_VAT_REGS),
        "read_vat_returns": canaries.canary_for(_KEY_VAT_RETURNS),
        "read_vat_rules": canaries.canary_for(_KEY_VAT_RULES),
    }

    scoring_hints = {
        "correctness": (
            "Transaction-by-transaction VAT analysis required. "
            "Reverse charge must be identified for B2B services. "
            "ERR-EU-010 must be detected (CE charged FR VAT on management fees). "
            f"Q2 proof of dispatch gap must be flagged with ~\u20ac{vat_exposure:,.0f} exposure."
        ),
        "completeness": (
            "All 6 IC transaction flows analyzed. 3+ missing data gaps flagged. "
            "VAT registration assessment for all entities. PE risk assessment. "
            "UK post-Brexit treatment confirmed."
        ),
        "eu_vat_framework": (
            "Art. 138 (intra-EU supply conditions), Art. 196 (reverse charge), "
            "UK as third country post-Brexit, VAT PE \u2260 income tax PE. "
            "No factor-based apportionment (US TC-10 methodology is wrong here)."
        ),
        "adversarial_elements": (
            "1. ERR-EU-010: CE\u2192CM management fee with incorrect FR VAT. "
            "2. Q2 proof of dispatch partial \u2014 quantify exposure. "
            "3. Polish registration without transactions \u2014 flag gap. "
            "4. CM Italian sales without registration \u2014 assess B2B/B2C. "
            "5. VAT PE vs income tax PE \u2014 do not conflate. "
            "6. CE Q4 pending assessment \u2014 flag as unusual."
        ),
        "communication": (
            "EU VAT terminology (reverse charge, intra-Community, Art. 196). "
            "Risk quantification (not just 'there is a gap'). "
            "Actionable recommendations (credit note, documentation, registration)."
        ),
    }

    return GoldStandard(
        test_case=_TC,
        expected_outputs=expected_outputs,
        canary_verification=canary_verification,
        error_detection={
            "ERR-EU-010": (
                f"CE (Netherlands) invoiced CM (France) for {ERR_EU_010_QUARTER} "
                "management fees with 20% French VAT. CE is not VAT-registered "
                "in France; reverse charge under Art. 196 applies. Requires "
                "credit note."
            ),
        },
        scoring_hints=scoring_hints,
    )


# ── Public entry point ───────────────────────────────────────────────────────

def emit_tc10_eu(
    model: CascadeModel,
    output_dir: Path,
    canaries: CanaryRegistry,
    errors: ErrorRegistry,
    manifest: Manifest,
    **kwargs: object,
) -> None:
    """Write all TC-10-EU files to *output_dir*."""
    _write_ic_sales_xlsx(output_dir, canaries, errors, manifest)
    _write_vat_registrations_xlsx(output_dir, canaries, manifest)
    _write_vat_returns_xlsx(output_dir, canaries, manifest)
    _write_vat_rules_docx(output_dir, canaries, manifest)
    _write_prompt(output_dir)
    _write_expected_behavior(output_dir)
