"""Tests for TC-10-EU — VAT and Cross-Border Tax Position Analysis.

Verifies:
- Intercompany sales XLSX (24 IC rows + third-party summary, correct columns)
- VAT registrations XLSX (5 registrations including pending PL)
- VAT returns summary XLSX (16 quarterly return rows)
- EU VAT rules reference DOCX (6 sections)
- ERR-EU-010 planted error (classification_error, severity material)
- Canary embedding in all 4 files
- Gold standard JSON (structure, expected outputs, scoring hints)
- Prompt and expected behavior markdown files
"""

from __future__ import annotations

import json
import tempfile
from pathlib import Path

import openpyxl
from docx import Document

from generator.canaries import CanaryRegistry
from generator.canaries import build_registry as build_canary_registry
from generator.errors import ErrorRegistry
from generator.formatters.tc10_eu import emit_tc10_eu
from generator.golds.framework import emit_gold
from generator.manifest import Manifest
from generator.model.build import CascadeModel, build_model
from generator.model.vat_eu import (
    ALL_CANARY_KEYS_TC10EU,
    ERR_EU_010_QUARTER,
    generate_ic_sales_eu,
    generate_vat_registrations,
    generate_vat_returns,
)

# ---------------------------------------------------------------------------
# Module-level fixture: build the model and run emit_tc10_eu once
# ---------------------------------------------------------------------------

_MODEL: CascadeModel | None = None
_OUTPUT: Path | None = None
_CANARIES: CanaryRegistry | None = None
_ERRORS: ErrorRegistry | None = None
_MANIFEST: Manifest | None = None


def _ensure_emitted() -> tuple[CascadeModel, Path, CanaryRegistry, ErrorRegistry, Manifest]:
    global _MODEL, _OUTPUT, _CANARIES, _ERRORS, _MANIFEST  # noqa: PLW0603
    if _MODEL is None:
        _MODEL = build_model(seed=42)
        _OUTPUT = Path(tempfile.mkdtemp(prefix="tc10eu_test_"))
        _CANARIES = build_canary_registry(ALL_CANARY_KEYS_TC10EU, seed=42)
        _ERRORS = ErrorRegistry()
        _MANIFEST = Manifest(_OUTPUT)
        _MANIFEST.__enter__()

        emit_tc10_eu(_MODEL, _OUTPUT, _CANARIES, _ERRORS, _MANIFEST)
        emit_gold("TC-10-EU", _CANARIES, _ERRORS, _OUTPUT / "gold_standards", model=_MODEL)

        _MANIFEST.__exit__(None, None, None)
    return _MODEL, _OUTPUT, _CANARIES, _ERRORS, _MANIFEST


_INPUT_DIR = "test_cases/TC-10-EU/input_files"


# ---------------------------------------------------------------------------
# Intercompany Sales XLSX
# ---------------------------------------------------------------------------


class TestIntercompanySales:
    """Verify tc10eu_intercompany_sales_fy2025.xlsx."""

    def test_xlsx_exists(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_intercompany_sales_fy2025.xlsx"
        assert path.exists()

    def test_xlsx_has_canary(self) -> None:
        _, output, canaries, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_intercompany_sales_fy2025.xlsx"
        wb = openpyxl.load_workbook(str(path))
        expected = canaries.canary_for("tc10eu_intercompany_sales_fy2025")
        desc = wb.properties.description or ""
        assert expected in desc, f"Canary {expected} not in IC sales xlsx"

    def test_has_intercompany_sales_sheet(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_intercompany_sales_fy2025.xlsx"
        wb = openpyxl.load_workbook(str(path))
        assert "Intercompany Sales" in wb.sheetnames

    def test_headers_at_row_3(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_intercompany_sales_fy2025.xlsx"
        wb = openpyxl.load_workbook(str(path))
        ws = wb["Intercompany Sales"]
        assert ws.cell(row=3, column=1).value == "Seller Entity"
        assert ws.cell(row=3, column=4).value == "Amount (EUR)"
        assert ws.cell(row=3, column=5).value == "VAT Treatment Applied"
        assert ws.cell(row=3, column=8).value == "Proof of Dispatch"

    def test_ic_row_count(self) -> None:
        """24 IC rows from model, data starting at row 4."""
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_intercompany_sales_fy2025.xlsx"
        wb = openpyxl.load_workbook(str(path))
        ws = wb["Intercompany Sales"]
        ic_count = len(generate_ic_sales_eu())
        # Count data rows (row 4 to 4 + ic_count - 1)
        data_rows = 0
        for r in range(4, 4 + ic_count + 5):
            val = ws.cell(row=r, column=1).value
            if val and "Third-Party" not in str(val):
                data_rows += 1
            elif val and "Third-Party" in str(val):
                break
            elif not val:
                # Could be blank row between IC and third-party
                continue
        assert data_rows == ic_count, f"Expected {ic_count} IC rows, got {data_rows}"

    def test_err_eu_010_present(self) -> None:
        """ERR-EU-010: CE charges 20% French VAT on CE->CM Q3 management fee."""
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_intercompany_sales_fy2025.xlsx"
        wb = openpyxl.load_workbook(str(path))
        ws = wb["Intercompany Sales"]
        found = False
        for row in ws.iter_rows(min_row=4, values_only=True):
            if not row[0]:
                continue
            desc = str(row[2]) if row[2] else ""
            vat_rate = str(row[5]) if row[5] else ""
            if "Q3" in desc and "Management" in desc and "20%" in vat_rate:
                found = True
                break
        assert found, "ERR-EU-010 row not found in IC sales data"

    def test_q2_partial_dispatch(self) -> None:
        """Q2 CP->CD should have 'Partial' proof of dispatch."""
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_intercompany_sales_fy2025.xlsx"
        wb = openpyxl.load_workbook(str(path))
        ws = wb["Intercompany Sales"]
        found = False
        for row in ws.iter_rows(min_row=4, values_only=True):
            if not row[0]:
                continue
            desc = str(row[2]) if row[2] else ""
            pod = str(row[7]) if row[7] else ""
            if "Q2" in desc and "Finished goods" in desc and "Partial" in pod:
                found = True
                break
        assert found, "Q2 CP->CD 'Partial' proof of dispatch not found"

    def test_third_party_summary_exists(self) -> None:
        """Third-party sales summary section should exist below IC data."""
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_intercompany_sales_fy2025.xlsx"
        wb = openpyxl.load_workbook(str(path))
        ws = wb["Intercompany Sales"]
        found = False
        for row in ws.iter_rows(values_only=True):
            if row[0] and "Third-Party" in str(row[0]):
                found = True
                break
        assert found, "Third-Party Sales Summary section not found"

    def test_cm_italian_sales_in_summary(self) -> None:
        """CM's Italian sales should appear in third-party summary notes."""
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_intercompany_sales_fy2025.xlsx"
        wb = openpyxl.load_workbook(str(path))
        ws = wb["Intercompany Sales"]
        found = False
        for row in ws.iter_rows(values_only=True):
            for cell_val in row:
                if cell_val and "Italian" in str(cell_val):
                    found = True
                    break
        assert found, "CM Italian sales not found in third-party summary"


# ---------------------------------------------------------------------------
# VAT Registrations XLSX
# ---------------------------------------------------------------------------


class TestVATRegistrations:
    """Verify tc10eu_vat_registrations.xlsx."""

    def test_xlsx_exists(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_vat_registrations.xlsx"
        assert path.exists()

    def test_xlsx_has_canary(self) -> None:
        _, output, canaries, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_vat_registrations.xlsx"
        wb = openpyxl.load_workbook(str(path))
        expected = canaries.canary_for("tc10eu_vat_registrations")
        desc = wb.properties.description or ""
        assert expected in desc

    def test_has_vat_registrations_sheet(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_vat_registrations.xlsx"
        wb = openpyxl.load_workbook(str(path))
        assert "VAT Registrations" in wb.sheetnames

    def test_headers_at_row_3(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_vat_registrations.xlsx"
        wb = openpyxl.load_workbook(str(path))
        ws = wb["VAT Registrations"]
        assert ws.cell(row=3, column=1).value == "Entity Code"
        assert ws.cell(row=3, column=3).value == "VAT ID"
        assert ws.cell(row=3, column=9).value == "Status"

    def test_registration_count(self) -> None:
        """5 registrations: CE-NL, CP-DE, CP-PL(pending), CM-FR, CD-UK."""
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_vat_registrations.xlsx"
        wb = openpyxl.load_workbook(str(path))
        ws = wb["VAT Registrations"]
        count = sum(1 for row in ws.iter_rows(min_row=4, values_only=True) if row[0])
        assert count == 5

    def test_cp_polish_pending(self) -> None:
        """CP has a pending Polish VAT registration (missing data trap)."""
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_vat_registrations.xlsx"
        wb = openpyxl.load_workbook(str(path))
        ws = wb["VAT Registrations"]
        found = False
        for row in ws.iter_rows(min_row=4, values_only=True):
            if row[0] == "CP" and row[1] == "Poland":
                assert "Pending" in str(row[8]), "Polish registration should be Pending"
                found = True
                break
        assert found, "CP Polish registration not found"


# ---------------------------------------------------------------------------
# VAT Returns Summary XLSX
# ---------------------------------------------------------------------------


class TestVATReturnsSummary:
    """Verify tc10eu_vat_returns_summary_fy2025.xlsx."""

    def test_xlsx_exists(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_vat_returns_summary_fy2025.xlsx"
        assert path.exists()

    def test_xlsx_has_canary(self) -> None:
        _, output, canaries, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_vat_returns_summary_fy2025.xlsx"
        wb = openpyxl.load_workbook(str(path))
        expected = canaries.canary_for("tc10eu_vat_returns_summary_fy2025")
        desc = wb.properties.description or ""
        assert expected in desc

    def test_has_quarterly_vat_returns_sheet(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_vat_returns_summary_fy2025.xlsx"
        wb = openpyxl.load_workbook(str(path))
        assert "Quarterly VAT Returns" in wb.sheetnames

    def test_return_count(self) -> None:
        """16 rows: 4 entities x 4 quarters."""
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_vat_returns_summary_fy2025.xlsx"
        wb = openpyxl.load_workbook(str(path))
        ws = wb["Quarterly VAT Returns"]
        count = sum(1 for row in ws.iter_rows(min_row=4, values_only=True) if row[0])
        assert count == 16

    def test_ce_q4_pending_assessment(self) -> None:
        """CE Q4 should show 'Filed - Pending Assessment'."""
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_vat_returns_summary_fy2025.xlsx"
        wb = openpyxl.load_workbook(str(path))
        ws = wb["Quarterly VAT Returns"]
        found = False
        for row in ws.iter_rows(min_row=4, values_only=True):
            if row[0] and "Cascade Europe" in str(row[0]) and row[1] == "Q4":
                assert "Pending Assessment" in str(row[7])
                found = True
                break
        assert found, "CE Q4 row not found"


# ---------------------------------------------------------------------------
# EU VAT Rules Reference DOCX
# ---------------------------------------------------------------------------


class TestVATRulesReference:
    """Verify tc10eu_eu_vat_rules_reference.docx."""

    def test_docx_exists(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_eu_vat_rules_reference.docx"
        assert path.exists()

    def test_docx_has_canary(self) -> None:
        _, output, canaries, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_eu_vat_rules_reference.docx"
        doc = Document(str(path))
        expected = canaries.canary_for("tc10eu_eu_vat_rules_reference")
        comments = doc.core_properties.comments or ""
        assert expected in comments

    def test_has_pe_distinction(self) -> None:
        """DOCX must mention that VAT PE is not the same as income tax PE."""
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_eu_vat_rules_reference.docx"
        doc = Document(str(path))
        text = "\n".join(p.text for p in doc.paragraphs)
        assert "not the same" in text.lower()

    def test_has_article_196(self) -> None:
        """DOCX must reference Article 196 (reverse charge)."""
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_eu_vat_rules_reference.docx"
        doc = Document(str(path))
        text = "\n".join(p.text for p in doc.paragraphs)
        assert "Article 196" in text or "Art. 196" in text

    def test_has_uk_post_brexit(self) -> None:
        """DOCX must reference UK post-Brexit treatment."""
        _, output, _, _, _ = _ensure_emitted()
        path = output / _INPUT_DIR / "tc10eu_eu_vat_rules_reference.docx"
        doc = Document(str(path))
        text = "\n".join(p.text for p in doc.paragraphs)
        assert "third country" in text.lower()


# ---------------------------------------------------------------------------
# Error Registration
# ---------------------------------------------------------------------------


class TestErrorRegistration:
    """Verify ERR-EU-010 is registered correctly."""

    def test_err_eu_010_registered(self) -> None:
        _, _, _, errors, _ = _ensure_emitted()
        assert "ERR-EU-010" in errors.entries

    def test_err_eu_010_type(self) -> None:
        _, _, _, errors, _ = _ensure_emitted()
        err = errors.get("ERR-EU-010")
        assert err.type == "classification_error"

    def test_err_eu_010_severity(self) -> None:
        _, _, _, errors, _ = _ensure_emitted()
        err = errors.get("ERR-EU-010")
        assert err.severity == "material"

    def test_err_eu_010_test_cases(self) -> None:
        _, _, _, errors, _ = _ensure_emitted()
        err = errors.get("ERR-EU-010")
        assert "TC-10-EU" in err.which_test_cases_should_catch

    def test_err_eu_010_description_mentions_reverse_charge(self) -> None:
        _, _, _, errors, _ = _ensure_emitted()
        err = errors.get("ERR-EU-010")
        assert "reverse charge" in err.description.lower() or "Art. 196" in err.description


# ---------------------------------------------------------------------------
# Gold Standard
# ---------------------------------------------------------------------------


class TestGoldStandard:
    """Verify TC-10-EU_gold.json."""

    def test_gold_json_exists(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "gold_standards" / "TC-10-EU_gold.json"
        assert path.exists()

    def test_gold_test_case(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "gold_standards" / "TC-10-EU_gold.json"
        gold = json.loads(path.read_text())
        assert gold["test_case"] == "TC-10-EU"

    def test_gold_has_transaction_analysis(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "gold_standards" / "TC-10-EU_gold.json"
        gold = json.loads(path.read_text())
        assert "transaction_analysis" in gold["expected_outputs"]

    def test_gold_has_6_flows(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "gold_standards" / "TC-10-EU_gold.json"
        gold = json.loads(path.read_text())
        flows = gold["expected_outputs"]["transaction_analysis"]["flows_analyzed"]
        assert len(flows) == 6

    def test_gold_has_err_eu_010(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "gold_standards" / "TC-10-EU_gold.json"
        gold = json.loads(path.read_text())
        assert "ERR-EU-010" in gold["error_detection"]

    def test_gold_has_canary_verification(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "gold_standards" / "TC-10-EU_gold.json"
        gold = json.loads(path.read_text())
        assert len(gold["canary_verification"]) == 4

    def test_gold_has_scoring_hints(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "gold_standards" / "TC-10-EU_gold.json"
        gold = json.loads(path.read_text())
        for key in ["correctness", "completeness", "format_compliance", "robustness", "communication"]:
            assert key in gold["scoring_hints"]

    def test_gold_scenario_pack(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "gold_standards" / "TC-10-EU_gold.json"
        gold = json.loads(path.read_text())
        assert gold["scenario_pack"] == "cascade_europe_ifrs"

    def test_gold_has_judgment_traps(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "gold_standards" / "TC-10-EU_gold.json"
        gold = json.loads(path.read_text())
        assert len(gold["judgment_traps"]) >= 4

    def test_gold_has_risk_summary(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "gold_standards" / "TC-10-EU_gold.json"
        gold = json.loads(path.read_text())
        risks = gold["expected_outputs"]["risk_summary"]["risks"]
        assert len(risks) >= 3
        risk_ids = {r["risk_id"] for r in risks}
        assert "ERR-EU-010" in risk_ids


# ---------------------------------------------------------------------------
# Prompt and Expected Behavior
# ---------------------------------------------------------------------------


class TestPromptAndExpectedBehavior:
    """Verify prompt.md and expected_behavior.md."""

    def test_prompt_exists(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "test_cases/TC-10-EU/prompt.md"
        assert path.exists()

    def test_prompt_mentions_vat(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "test_cases/TC-10-EU/prompt.md"
        text = path.read_text()
        assert "VAT" in text

    def test_expected_behavior_exists(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "test_cases/TC-10-EU/expected_behavior.md"
        assert path.exists()

    def test_expected_behavior_mentions_err_eu_010(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "test_cases/TC-10-EU/expected_behavior.md"
        text = path.read_text()
        assert "ERR-EU-010" in text

    def test_expected_behavior_mentions_missing_data(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "test_cases/TC-10-EU/expected_behavior.md"
        text = path.read_text()
        assert "proof of dispatch" in text.lower()
        assert "Polish" in text
        assert "Italian" in text


# ---------------------------------------------------------------------------
# Model Tests
# ---------------------------------------------------------------------------


class TestModelData:
    """Verify model-level data integrity."""

    def test_ic_sales_count(self) -> None:
        rows = generate_ic_sales_eu()
        assert len(rows) == 24

    def test_ic_sales_has_error_row(self) -> None:
        rows = generate_ic_sales_eu()
        err_rows = [r for r in rows if r.is_error]
        assert len(err_rows) == 1
        assert err_rows[0].quarter == ERR_EU_010_QUARTER

    def test_vat_registrations_count(self) -> None:
        regs = generate_vat_registrations()
        assert len(regs) == 5

    def test_vat_registrations_has_pending_pl(self) -> None:
        regs = generate_vat_registrations()
        pl_regs = [r for r in regs if r.country == "Poland"]
        assert len(pl_regs) == 1
        assert "Pending" in pl_regs[0].status

    def test_vat_returns_count(self) -> None:
        rets = generate_vat_returns()
        assert len(rets) == 16

    def test_vat_returns_ce_q4_pending(self) -> None:
        rets = generate_vat_returns()
        ce_q4 = [r for r in rets if r.entity_code == "CE" and r.quarter == "Q4"]
        assert len(ce_q4) == 1
        assert "Pending Assessment" in ce_q4[0].filing_status

    def test_canary_keys_count(self) -> None:
        assert len(ALL_CANARY_KEYS_TC10EU) == 4

    def test_canary_keys_sorted(self) -> None:
        assert ALL_CANARY_KEYS_TC10EU == sorted(ALL_CANARY_KEYS_TC10EU)
