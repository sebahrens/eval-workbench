"""Formatter: TC-07 — K-1 Extraction & Consolidation (Tax, Adversarial).

Emits:
- test_cases/TC-07/input_files/k1s/K1-001.pdf through K1-008.pdf
  8 Schedule K-1 PDFs (3 system-clean, 5 varying layouts)
- test_cases/TC-07/input_files/entity_org_chart.pdf
  Org chart showing which entities these K-1s flow through
- test_cases/TC-07/prompt.md
- test_cases/TC-07/expected_behavior.md
- gold_standards/TC-07_gold.json

One planted error:
  ERR-014 — K1-003 shows wrong entity name (wrong_entity).

Uses the canonical model — never hardcodes numbers.
"""

from __future__ import annotations

import datetime
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path
from typing import Any

from fpdf import FPDF

from generator.canaries import CanaryRegistry, embed_canary_pdf_fpdf2
from generator.errors import ErrorRegistry, PlantedError, wrong_entity
from generator.golds.framework import GoldStandard, register_gold
from generator.manifest import Manifest
from generator.model.build import CascadeModel
from generator.model.entities import ENTITIES
from generator.model.k1 import (
    K1Investment,
    K1LayoutType,
    consolidated_totals,
    generate_k1_investments,
)

# ── Constants ────────────────────────────────────────────────────────────────

_TC = "TC-07"
_INPUT_DIR = f"test_cases/{_TC}/input_files"
_K1S_DIR = f"{_INPUT_DIR}/k1s"

_FIXED_DATE = "2025-03-15"
_TAX_YEAR = 2025

# fpdf2 fixed creation date for determinism
_CREATION_DATE = datetime.datetime(2025, 3, 15, 9, 0, 0)


def _fmt_dollars(d: Decimal | int | None) -> str:
    """Format a Decimal as a dollar string like '$285,000'."""
    if d is None:
        return ""
    if isinstance(d, int):
        val = d
    else:
        val = int(d.quantize(Decimal("1"), rounding=ROUND_HALF_UP))
    if val < 0:
        return f"(${abs(val):,})"
    return f"${val:,}"


def _fmt_int(d: Decimal | int | None) -> int | None:
    """Convert Decimal to int for gold standard."""
    if d is None:
        return None
    if isinstance(d, int):
        return d
    return int(d.quantize(Decimal("1"), rounding=ROUND_HALF_UP))


# ── Canary file keys ────────────────────────────────────────────────────────
# One per K-1 PDF plus the org chart.

def _k1_canary_key(k1_id: str) -> str:
    """Return the canary file key for a K-1 PDF, e.g. 'tc07_k1_001'."""
    num = k1_id.split("-")[1]
    return f"tc07_k1_{num}"


_ORG_CHART_KEY = "tc07_entity_org_chart"

_ALL_CANARY_KEYS: list[str] = sorted(
    [_k1_canary_key(f"K1-{i:03d}") for i in range(1, 9)]
    + [_ORG_CHART_KEY]
)


# ── PDF helpers ──────────────────────────────────────────────────────────────


def _new_pdf(canary: str) -> FPDF:
    """Create an FPDF instance with deterministic metadata and embedded canary."""
    pdf = FPDF()
    pdf.set_creation_date(_CREATION_DATE)
    pdf.set_auto_page_break(auto=False)
    embed_canary_pdf_fpdf2(pdf, canary)
    return pdf


# ── K-1 PDF renderers ───────────────────────────────────────────────────────
# Three "system-clean" and five "varying" layout styles.


# Box label mapping (subset — boxes 1-13 + guaranteed payments)
_BOX_LABELS: list[tuple[str, str]] = [
    ("box_1_ordinary_income", "1  Ordinary business income (loss)"),
    ("box_2_net_rental_income", "2  Net rental real estate income (loss)"),
    ("box_3_other_rental_income", "3  Other net rental income (loss)"),
    ("box_4a_guaranteed_payments_services", "4a Guaranteed payments for services"),
    ("box_4b_guaranteed_payments_capital", "4b Guaranteed payments for capital"),
    ("box_4c_total_guaranteed_payments", "4c Total guaranteed payments"),
    ("box_5_interest_income", "5  Interest income"),
    ("box_6a_ordinary_dividends", "6a Ordinary dividends"),
    ("box_6b_qualified_dividends", "6b Qualified dividends"),
    ("box_7_royalties", "7  Royalties"),
    ("box_8_net_st_capital_gain", "8  Net short-term capital gain (loss)"),
    ("box_9a_net_lt_capital_gain", "9a Net long-term capital gain (loss)"),
    ("box_9b_collectibles_gain", "9b Collectibles (28%) gain (loss)"),
    ("box_9c_unrecaptured_1250", "9c Unrecaptured section 1250 gain"),
    ("box_10_net_1231_gain", "10 Net section 1231 gain (loss)"),
    ("box_11_other_income", "11 Other income (loss)"),
    ("box_12_section_179", "12 Section 179 deduction"),
    ("box_13_other_deductions", "13 Other deductions"),
]


def _render_system_clean(
    pdf: FPDF, inv: K1Investment, *, entity_name_override: str | None = None,
) -> None:
    """Render a clean, system-generated K-1 layout."""
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 8, "Schedule K-1 (Form 1065)", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(
        0, 5,
        "Partner's Share of Income, Deductions, Credits, etc.",
        new_x="LMARGIN", new_y="NEXT",
    )
    pdf.cell(
        0, 5,
        f"For calendar year {_TAX_YEAR}, or tax year beginning _________, ending _________",
        new_x="LMARGIN", new_y="NEXT",
    )
    pdf.ln(3)

    if inv.is_amended:
        pdf.set_font("Helvetica", "B", 11)
        pdf.cell(0, 6, "*** AMENDED K-1 ***", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "I", 9)
        pdf.cell(0, 5, "X  Amended K-1     Final K-1", new_x="LMARGIN", new_y="NEXT")
        pdf.ln(2)

    # Partnership info
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 6, "Part I - Information About the Partnership", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 9)
    pdf.cell(0, 5, f"A  Partnership's EIN: {inv.partnership_ein}", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, f"B  Partnership's name: {inv.partnership_name}", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    # Partner info
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 6, "Part II - Information About the Partner", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 9)
    if entity_name_override is not None:
        partner_name = entity_name_override
    else:
        entity = ENTITIES.get(inv.entity_code)
        partner_name = entity.name if entity else inv.entity_code
    pdf.cell(0, 5, f"F  Partner's name: {partner_name}", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(3)

    # Part III — boxes
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(
        0, 6,
        "Part III - Partner's Share of Current Year Income, Deductions, Credits",
        new_x="LMARGIN", new_y="NEXT",
    )
    pdf.set_font("Helvetica", "", 9)

    for attr, label in _BOX_LABELS:
        val = getattr(inv, attr)
        if val is not None:
            pdf.cell(120, 5, label)
            pdf.cell(0, 5, _fmt_dollars(val), new_x="LMARGIN", new_y="NEXT")

    # Box 20
    if inv.box_20_codes:
        pdf.ln(2)
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(0, 5, "20 Other information", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 9)
        for code in inv.box_20_codes:
            pdf.cell(120, 5, f"    Code {code.code} - {code.description}")
            pdf.cell(0, 5, _fmt_dollars(code.amount), new_x="LMARGIN", new_y="NEXT")

    # Section 199A
    if inv.section_199a_qbi is not None:
        pdf.ln(2)
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(0, 5, "Section 199A Information", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 9)
        pdf.cell(120, 5, "    Qualified Business Income (QBI)")
        pdf.cell(0, 5, _fmt_dollars(inv.section_199a_qbi), new_x="LMARGIN", new_y="NEXT")
        if inv.section_199a_wages is not None:
            pdf.cell(120, 5, "    W-2 Wages")
            pdf.cell(0, 5, _fmt_dollars(inv.section_199a_wages), new_x="LMARGIN", new_y="NEXT")
        if inv.section_199a_ubia is not None:
            pdf.cell(120, 5, "    UBIA of Qualified Property")
            pdf.cell(0, 5, _fmt_dollars(inv.section_199a_ubia), new_x="LMARGIN", new_y="NEXT")

    # Amendment details
    if inv.is_amended and inv.amendments:
        pdf.ln(3)
        pdf.set_font("Helvetica", "B", 9)
        pdf.cell(0, 5, "Changes from Original K-1:", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font("Helvetica", "", 9)
        for amend in inv.amendments:
            pdf.cell(0, 5, f"  - {amend.description}", new_x="LMARGIN", new_y="NEXT")

    # Footer
    pdf.ln(5)
    pdf.set_font("Helvetica", "I", 7)
    pdf.cell(0, 4, f"Generated {_FIXED_DATE}", new_x="LMARGIN", new_y="NEXT")


# ── Varying layout styles ────────────────────────────────────────────────────
# Each varying-layout K-1 uses a slightly different font, spacing, and format
# to simulate PDFs from different partnership preparers.

_VARYING_FONTS = ["Helvetica", "Courier", "Times", "Helvetica", "Courier"]
_VARYING_SIZES = [9, 8, 10, 9, 8]
_VARYING_LINE_HEIGHTS = [5, 4.5, 5.5, 5, 4.5]


def _render_varying(pdf: FPDF, inv: K1Investment, style_idx: int) -> None:
    """Render a K-1 with varying layout/font/spacing."""
    font_family = _VARYING_FONTS[style_idx % len(_VARYING_FONTS)]
    font_size = _VARYING_SIZES[style_idx % len(_VARYING_SIZES)]
    line_h = _VARYING_LINE_HEIGHTS[style_idx % len(_VARYING_LINE_HEIGHTS)]

    pdf.add_page()

    # Title — different styles per preparer
    pdf.set_font(font_family, "B", font_size + 4)
    if style_idx % 2 == 0:
        pdf.cell(0, line_h + 3, "SCHEDULE K-1 (FORM 1065)", new_x="LMARGIN", new_y="NEXT")
    else:
        pdf.cell(0, line_h + 3, "Schedule K-1  -  Form 1065", new_x="LMARGIN", new_y="NEXT")

    pdf.set_font(font_family, "", font_size)
    pdf.cell(
        0, line_h,
        f"Tax Year: {_TAX_YEAR}",
        new_x="LMARGIN", new_y="NEXT",
    )

    if inv.is_amended:
        pdf.set_font(font_family, "B", font_size + 1)
        pdf.cell(0, line_h + 1, "[X] CORRECTED / AMENDED", new_x="LMARGIN", new_y="NEXT")

    pdf.ln(line_h * 0.5)

    # Partnership info — varying format
    pdf.set_font(font_family, "B", font_size)
    if style_idx % 3 == 0:
        pdf.cell(0, line_h, "PARTNERSHIP INFORMATION", new_x="LMARGIN", new_y="NEXT")
    else:
        pdf.cell(0, line_h, "Part I: Partnership Details", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font(font_family, "", font_size)
    pdf.cell(0, line_h, f"EIN: {inv.partnership_ein}", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, line_h, f"Name: {inv.partnership_name}", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(line_h * 0.5)

    # Partner info
    entity = ENTITIES.get(inv.entity_code)
    partner_name = entity.name if entity else inv.entity_code
    pdf.set_font(font_family, "B", font_size)
    pdf.cell(0, line_h, "Partner:", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font(font_family, "", font_size)
    pdf.cell(0, line_h, partner_name, new_x="LMARGIN", new_y="NEXT")
    pdf.ln(line_h * 0.5)

    # Income/deduction boxes — varying column width
    pdf.set_font(font_family, "B", font_size)
    if style_idx % 2 == 0:
        pdf.cell(0, line_h, "INCOME / DEDUCTIONS / CREDITS", new_x="LMARGIN", new_y="NEXT")
    else:
        pdf.cell(0, line_h, "Part III - Income, Deductions, Credits", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font(font_family, "", font_size)

    # Varying column widths
    label_w = 100 + (style_idx % 3) * 10

    for attr, label in _BOX_LABELS:
        val = getattr(inv, attr)
        if val is not None:
            # Some varying layouts use different label formats
            if style_idx % 2 == 0:
                display_label = label.upper()
            else:
                display_label = label
            pdf.cell(label_w, line_h, display_label)
            pdf.cell(0, line_h, _fmt_dollars(val), new_x="LMARGIN", new_y="NEXT")

    # Box 20
    if inv.box_20_codes:
        pdf.ln(line_h * 0.3)
        pdf.set_font(font_family, "B", font_size)
        pdf.cell(0, line_h, "Box 20 - Other Information:", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font(font_family, "", font_size)
        for code in inv.box_20_codes:
            pdf.cell(label_w, line_h, f"  Code {code.code}: {code.description}")
            pdf.cell(0, line_h, _fmt_dollars(code.amount), new_x="LMARGIN", new_y="NEXT")

    # Section 199A
    if inv.section_199a_qbi is not None:
        pdf.ln(line_h * 0.3)
        pdf.set_font(font_family, "B", font_size)
        pdf.cell(0, line_h, "Sec. 199A Information", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font(font_family, "", font_size)
        pdf.cell(label_w, line_h, "  QBI")
        pdf.cell(0, line_h, _fmt_dollars(inv.section_199a_qbi), new_x="LMARGIN", new_y="NEXT")
        if inv.section_199a_wages is not None:
            pdf.cell(label_w, line_h, "  W-2 Wages")
            pdf.cell(0, line_h, _fmt_dollars(inv.section_199a_wages), new_x="LMARGIN", new_y="NEXT")
        if inv.section_199a_ubia is not None:
            pdf.cell(label_w, line_h, "  UBIA")
            pdf.cell(0, line_h, _fmt_dollars(inv.section_199a_ubia), new_x="LMARGIN", new_y="NEXT")

    # Amendment details
    if inv.is_amended and inv.amendments:
        pdf.ln(line_h)
        pdf.set_font(font_family, "B", font_size)
        pdf.cell(0, line_h, "AMENDMENT NOTES:", new_x="LMARGIN", new_y="NEXT")
        pdf.set_font(font_family, "", font_size)
        for amend in inv.amendments:
            pdf.cell(0, line_h, f"* {amend.description}", new_x="LMARGIN", new_y="NEXT")

    # Footer — varying positions
    pdf.ln(line_h * 2)
    pdf.set_font(font_family, "I", font_size - 2)
    pdf.cell(0, line_h, f"Prepared {_FIXED_DATE}", new_x="LMARGIN", new_y="NEXT")


# ── Entity Org Chart PDF ────────────────────────────────────────────────────


def _write_org_chart(
    investments: list[K1Investment],
    output_dir: Path,
    canaries: CanaryRegistry,
    manifest: Manifest,
) -> None:
    """Write entity_org_chart.pdf showing K-1 flow-through structure."""
    canary_code = canaries.canary_for(_ORG_CHART_KEY)
    pdf = _new_pdf(canary_code)
    pdf.add_page()

    # Title
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, "Cascade Industries, Inc.", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.set_font("Helvetica", "", 12)
    pdf.cell(0, 7, "Partnership Investment Structure", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.cell(0, 6, f"Tax Year {_TAX_YEAR}", new_x="LMARGIN", new_y="NEXT", align="C")
    pdf.ln(8)

    # Parent entity
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 7, "Cascade Industries, Inc. (C-Corporation - Parent)", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 9)
    pdf.cell(0, 5, "EIN: 93-1234567  |  Portland, OR", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)

    # Group investments by entity
    by_entity: dict[str, list[K1Investment]] = {}
    for inv in investments:
        by_entity.setdefault(inv.entity_code, []).append(inv)

    for entity_code in sorted(by_entity.keys()):
        entity = ENTITIES[entity_code]
        invs = sorted(by_entity[entity_code], key=lambda k: k.k1_id)

        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(
            0, 7,
            f"  |-- {entity.name} ({entity_code})",
            new_x="LMARGIN", new_y="NEXT",
        )
        pdf.set_font("Helvetica", "", 9)
        pdf.cell(0, 5, f"       {entity.location}", new_x="LMARGIN", new_y="NEXT")

        for inv in invs:
            amended_flag = " [AMENDED]" if inv.is_amended else ""
            pdf.cell(
                0, 5,
                f"       |-- {inv.partnership_name} (EIN: {inv.partnership_ein}){amended_flag}",
                new_x="LMARGIN", new_y="NEXT",
            )
            pdf.cell(
                0, 5,
                f"           K-1 ID: {inv.k1_id}  |  Total Income: {_fmt_dollars(inv.total_income)}",
                new_x="LMARGIN", new_y="NEXT",
            )
        pdf.ln(3)

    # Note about C-corp
    pdf.ln(5)
    pdf.set_font("Helvetica", "I", 9)
    pdf.cell(
        0, 5,
        "Note: Cascade Industries is a C-corporation. Section 199A QBI deductions",
        new_x="LMARGIN", new_y="NEXT",
    )
    pdf.cell(
        0, 5,
        "reported on K-1s are NOT applicable to C-corporation filers.",
        new_x="LMARGIN", new_y="NEXT",
    )

    path = output_dir / _INPUT_DIR / "entity_org_chart.pdf"
    path.parent.mkdir(parents=True, exist_ok=True)
    pdf.output(str(path))

    canaries.set_location(
        _ORG_CHART_KEY,
        f"{_INPUT_DIR}/entity_org_chart.pdf",
        "PDF metadata → Subject",
    )
    manifest.register(
        f"{_INPUT_DIR}/entity_org_chart.pdf",
        "pdf",
        canary=canary_code,
        test_cases=[_TC],
    )


# ── K-1 PDF writer ──────────────────────────────────────────────────────────


def _write_k1_pdfs(
    investments: list[K1Investment],
    output_dir: Path,
    canaries: CanaryRegistry,
    manifest: Manifest,
    errors: ErrorRegistry,
) -> None:
    """Write 8 K-1 PDFs to the k1s/ subdirectory."""
    k1s_path = output_dir / _K1S_DIR
    k1s_path.mkdir(parents=True, exist_ok=True)

    # ERR-014: K1-003 will show the wrong entity name
    _ERR_014_K1 = "K1-003"
    correct_entity = ENTITIES["AM"].name   # "Cascade Advanced Materials, Inc."
    wrong_name = wrong_entity(correct_entity, "Cascade Precision Components LLC")

    varying_idx = 0
    for inv in investments:
        key = _k1_canary_key(inv.k1_id)
        canary_code = canaries.canary_for(key)
        pdf = _new_pdf(canary_code)

        if inv.layout_type == K1LayoutType.SYSTEM_CLEAN:
            if inv.k1_id == _ERR_014_K1:
                _render_system_clean(pdf, inv, entity_name_override=wrong_name)
            else:
                _render_system_clean(pdf, inv)
        else:
            _render_varying(pdf, inv, varying_idx)
            varying_idx += 1

        filename = f"{inv.k1_id}.pdf"
        path = k1s_path / filename
        pdf.output(str(path))

        canaries.set_location(
            key,
            f"{_K1S_DIR}/{filename}",
            "PDF metadata → Subject",
        )
        manifest.register(
            f"{_K1S_DIR}/{filename}",
            "pdf",
            canary=canary_code,
            test_cases=[_TC],
        )

    # Register ERR-014
    errors.add(PlantedError(
        error_id="ERR-014",
        file=f"{_K1S_DIR}/{_ERR_014_K1}.pdf",
        location="Part II — Partner's name",
        type="wrong_entity",
        description=(
            f"K1-003 shows '{wrong_name}' instead of '{correct_entity}'"
        ),
        severity="material",
        which_test_cases_should_catch=[_TC],
    ))


# ── Prompt & expected behavior ───────────────────────────────────────────────


def _write_prompt(output_dir: Path) -> None:
    """Write the TC-07 prompt.md."""
    text = """\
# TC-07: K-1 Extraction & Consolidation

## Input Files
- `k1s/` — 8 Schedule K-1 PDFs from partnership investments
- `entity_org_chart.pdf` — Shows which entities hold these partnership investments

## Instructions

Cascade Industries received K-1s from 8 partnership investments.

1. Extract all income, deduction, and credit items from each K-1.
2. Organize by K-1 box number and category (ordinary income, rental income,
   guaranteed payments, interest, dividends, capital gains, Section 179,
   charitable contributions, etc.).
3. Identify the amended K-1 and note what changed from the original.
4. Consolidate all K-1 data into a single summary schedule.
5. Map each item to the appropriate line on the corporate return (Form 1120).
6. Flag any items requiring special handling (e.g., passive activity limitations,
   at-risk limitations, Section 199A deductions flowing to a C-corp).

Export the consolidated schedule as an Excel file.
"""
    path = output_dir / f"test_cases/{_TC}/prompt.md"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text)


def _write_expected_behavior(output_dir: Path) -> None:
    """Write the TC-07 expected_behavior.md."""
    text = """\
# TC-07: Expected Behavior

## Key Evaluation Criteria

### Data Extraction
- All 8 K-1s must be parsed correctly despite layout variation.
- All box values (1-13) and Box 20 codes must be extracted.
- Section 199A information must be identified on each K-1.

### Amended K-1 Detection
- The agent must identify K1-004 (Columbia Basin Energy Partners LLC) as an amended K-1.
- The amendment changed ordinary income from $340,000 to $285,000.
- A $55,000 guaranteed payment was added in the amendment.
- The agent should use the amended values (not originals) in the consolidation.

### Consolidation Accuracy
- The consolidated schedule must sum all K-1 boxes across all 8 partnerships.
- Totals must be exact — no rounding tolerance.
- Items should be organized by box number and category.

### Form 1120 Mapping
- Each K-1 item should be mapped to the correct Form 1120 line.
- Ordinary income → Form 1120 Line 4 (other income).
- Interest → Line 5. Dividends → Line 6 (via Schedule C). Etc.

### Section 199A Flag
- The agent MUST flag that Section 199A QBI deductions are NOT applicable
  to Cascade Industries because it is a C-corporation filer.
- This is a critical adversarial element — reporting 199A as a deduction
  would be an error.

### Special Handling Items
- Passive activity limitations (rental income from K1-002).
- At-risk limitations assessment.
- Capital gains requiring Schedule D treatment.
- Section 1231 gains requiring Form 4797.
"""
    path = output_dir / f"test_cases/{_TC}/expected_behavior.md"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text)


# ── Gold Standard ────────────────────────────────────────────────────────────

# Form 1120 line mapping for K-1 items
_FORM_1120_MAPPING: dict[str, str] = {
    "Box 1 - Ordinary business income": "Form 1120, Line 4 (Other income)",
    "Box 2 - Net rental real estate income": "Form 1120, Line 4 (Other income) — passive activity rules apply",
    "Box 4c - Guaranteed payments": "Form 1120, Line 4 (Other income)",
    "Box 5 - Interest income": "Form 1120, Line 5 (Interest)",
    "Box 6a - Ordinary dividends": "Form 1120, Schedule C → Line 6 (Dividends)",
    "Box 7 - Royalties": "Form 1120, Line 4 (Other income)",
    "Box 8 - Net ST capital gain": "Form 1120, Schedule D → Line 8 (Capital gain net income)",
    "Box 9a - Net LT capital gain": "Form 1120, Schedule D → Line 8 (Capital gain net income)",
    "Box 9c - Unrecaptured Sec 1250 gain": "Form 1120, Schedule D (capital gain, unrecaptured §1250)",
    "Box 10 - Net section 1231 gain": "Form 4797 → Form 1120, Line 4",
    "Box 11 - Other income": "Form 1120, Line 4 (Other income)",
    "Box 12 - Section 179 deduction": "Form 1120, Line 20 (Depreciation — §179)",
    "Box 13 - Other deductions": "Form 1120, Line 26 (Other deductions)",
}


@register_gold("TC-07")
def _tc07_gold(
    canaries: CanaryRegistry,
    errors: ErrorRegistry,
    **model_kwargs: Any,
) -> GoldStandard:
    """Build the TC-07 gold standard from the canonical model."""
    investments = generate_k1_investments()
    totals = consolidated_totals(investments)

    # Per-K-1 detail
    k1_details: dict[str, dict[str, Any]] = {}
    for inv in investments:
        detail: dict[str, Any] = {
            "partnership_name": inv.partnership_name,
            "partnership_ein": inv.partnership_ein,
            "entity": inv.entity_code,
            "is_amended": inv.is_amended,
            "layout_type": inv.layout_type.value,
        }
        # Populated boxes
        boxes: dict[str, int] = {}
        for attr, label in _BOX_LABELS:
            val = getattr(inv, attr)
            if val is not None:
                boxes[label] = _fmt_int(val)
        detail["boxes"] = boxes

        if inv.box_20_codes:
            detail["box_20"] = {
                f"Code {c.code}": {"description": c.description, "amount": _fmt_int(c.amount)}
                for c in inv.box_20_codes
            }

        if inv.section_199a_qbi is not None:
            sec_199a: dict[str, int | None] = {
                "qbi": _fmt_int(inv.section_199a_qbi),
            }
            if inv.section_199a_wages is not None:
                sec_199a["wages"] = _fmt_int(inv.section_199a_wages)
            if inv.section_199a_ubia is not None:
                sec_199a["ubia"] = _fmt_int(inv.section_199a_ubia)
            detail["section_199a"] = sec_199a

        if inv.is_amended:
            detail["amendments"] = [
                {
                    "field": a.field_changed,
                    "original": _fmt_int(a.original_value),
                    "amended": _fmt_int(a.amended_value),
                    "description": a.description,
                }
                for a in inv.amendments
            ]

        k1_details[inv.k1_id] = detail

    # Consolidated totals as ints
    totals_int = {k: _fmt_int(v) for k, v in sorted(totals.items())}

    # Total income and deductions across all K-1s
    grand_income = sum(inv.total_income for inv in investments)
    grand_deductions = sum(inv.total_deductions for inv in investments)

    # Canary verification — one per file
    canary_verification: dict[str, str] = {}
    for inv in investments:
        key = _k1_canary_key(inv.k1_id)
        canary_verification[f"read_{inv.k1_id}"] = canaries.canary_for(key)
    canary_verification["read_org_chart"] = canaries.canary_for(_ORG_CHART_KEY)

    return GoldStandard(
        test_case=_TC,
        expected_outputs={
            "file_type": "xlsx",
            "description": "Consolidated K-1 schedule with all 8 partnerships",
            "k1_count": 8,
            "k1_details": k1_details,
            "consolidated_totals": totals_int,
            "grand_total_income": _fmt_int(grand_income),
            "grand_total_deductions": _fmt_int(grand_deductions),
            "form_1120_mapping": _FORM_1120_MAPPING,
            "amended_k1": {
                "k1_id": "K1-004",
                "partnership": "Columbia Basin Energy Partners LLC",
                "changes": [
                    {
                        "field": "Ordinary income (Box 1)",
                        "original": 340000,
                        "amended": 285000,
                    },
                    {
                        "field": "Guaranteed payments (Box 4c)",
                        "original": 0,
                        "amended": 55000,
                    },
                ],
            },
            "section_199a_flag": (
                "Section 199A QBI deductions are NOT applicable to "
                "Cascade Industries as a C-corporation filer. Any K-1s "
                "reporting 199A amounts should be noted but NOT taken "
                "as a deduction on the corporate return."
            ),
            "special_handling": [
                "Passive activity limitations — rental income (K1-002)",
                "Capital gains requiring Schedule D (K1-001, K1-003, K1-007, K1-008)",
                "Section 1231 gains requiring Form 4797 (K1-004, K1-008)",
                "Unrecaptured §1250 gain (K1-002)",
            ],
        },
        canary_verification=canary_verification,
        error_detection={
            "ERR-014": (
                "K1-003 displays the wrong entity name — "
                "'Cascade Precision Components LLC' instead of "
                "'Cascade Advanced Materials, Inc.'"
            ),
        },
        scoring_hints={
            "correctness": (
                "All 8 K-1 values extracted correctly; consolidated totals "
                "match gold standard exactly; amended values used (not originals)"
            ),
            "completeness": (
                "All boxes extracted; Box 20 codes captured; Section 199A "
                "data identified; Form 1120 mapping provided; special "
                "handling items flagged"
            ),
            "format_compliance": (
                "Valid xlsx with per-K-1 detail and consolidated summary"
            ),
            "communication": (
                "Identified amended K-1 and described changes; flagged "
                "Section 199A N/A for C-corp; noted passive activity and "
                "capital gain special handling"
            ),
        },
    )


# ── Public entry point ──────────────────────────────────────────────────────


def emit_tc07(
    model: CascadeModel,
    output_dir: Path,
    canaries: CanaryRegistry,
    errors: ErrorRegistry,
    manifest: Manifest,
    **kwargs: object,
) -> None:
    """Write all TC-07 files to *output_dir*."""
    investments = generate_k1_investments()
    _write_k1_pdfs(investments, output_dir, canaries, manifest, errors)
    _write_org_chart(investments, output_dir, canaries, manifest)
    _write_prompt(output_dir)
    _write_expected_behavior(output_dir)
