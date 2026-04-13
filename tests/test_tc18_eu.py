"""Tests for TC-18-EU — IFRS/ISA Prior-Year Workpaper Rollforward formatter.

Verifies:
- 6 prior-year xlsx workpapers and 4 docx memos emitted
- 5 current-year data files (CSV, xlsx, docx) emitted
- Canary embedding across all 15 files
- Planted error ERR-EU-018 (stale Gewerbesteuer Hebesatz in fixed assets)
- Revenue workpaper content with IFRS 15 data
- Lease schedule with 4 original + 3 new leases (7 total)
- IAS 36 CGU goodwill impairment data
- IFRS/ISA terminology in memos
- Gold standard with correct structure
- Prompt and expected behavior markdown
"""

from __future__ import annotations

import json
import tempfile
from decimal import Decimal
from pathlib import Path

import openpyxl
import pytest
from docx import Document

from generator.canaries import CanaryRegistry
from generator.canaries import build_registry as build_canary_registry
from generator.errors import ErrorRegistry
from generator.formatters.tc18_eu import (
    _CANARY_KEYS,
    _GOODWILL_CGUS,
    _REVENUE_FY2024,
    _TOTAL_GOODWILL,
    emit_tc18_eu,
)
from generator.golds.framework import emit_gold
from generator.manifest import Manifest
from generator.model.build import CascadeModel, build_model

# ---------------------------------------------------------------------------
# Module-level fixture: build the model and run emit_tc18_eu once
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
        _OUTPUT = Path(tempfile.mkdtemp(prefix="tc18eu_test_"))
        _CANARIES = build_canary_registry(_CANARY_KEYS, seed=42)
        _ERRORS = ErrorRegistry()
        _MANIFEST = Manifest(_OUTPUT)
        _MANIFEST.__enter__()
        emit_tc18_eu(_MODEL, _OUTPUT, _CANARIES, _ERRORS, _MANIFEST)
        emit_gold("TC-18-EU", _CANARIES, _ERRORS, _OUTPUT / "gold_standards")
        _MANIFEST.__exit__(None, None, None)
    return _MODEL, _OUTPUT, _CANARIES, _ERRORS, _MANIFEST


_INPUT_DIR = "test_cases/TC-18-EU/input_files"
_PY_DIR = f"{_INPUT_DIR}/prior_year_workpapers"
_CY_DIR = f"{_INPUT_DIR}/current_year_data"

# ---------------------------------------------------------------------------
# File existence and openability
# ---------------------------------------------------------------------------

_PRIOR_YEAR_XLSX = [
    "wp_revenue_fy2024.xlsx",
    "wp_operating_expenses_fy2024.xlsx",
    "wp_balance_sheet_fy2024.xlsx",
    "wp_cash_fy2024.xlsx",
    "wp_fixed_assets_fy2024.xlsx",
    "wp_leases_fy2024.xlsx",
]

_PRIOR_YEAR_DOCX = [
    "memo_planning_fy2024.docx",
    "memo_risk_assessment_fy2024.docx",
    "memo_summary_fy2024.docx",
    "memo_management_letter_fy2024.docx",
]

_CURRENT_YEAR_FILES = [
    ("trial_balance_fy2025.csv", "csv"),
    ("bank_statements_fy2025.csv", "csv"),
    ("lease_schedule_fy2025.xlsx", "xlsx"),
    ("management_projections_fy2025.docx", "docx"),
    ("goodwill_impairment_analysis_ifrs.xlsx", "xlsx"),
]


class TestPriorYearXlsxExist:
    """All 6 prior-year xlsx workpapers must exist and open cleanly."""

    @pytest.mark.parametrize("filename", _PRIOR_YEAR_XLSX)
    def test_exists(self, filename: str) -> None:
        _, out, _, _, _ = _ensure_emitted()
        assert (out / _PY_DIR / filename).exists()

    @pytest.mark.parametrize("filename", _PRIOR_YEAR_XLSX)
    def test_opens_cleanly(self, filename: str) -> None:
        _, out, _, _, _ = _ensure_emitted()
        wb = openpyxl.load_workbook(out / _PY_DIR / filename, data_only=True)
        assert wb.sheetnames


class TestPriorYearDocxExist:
    """All 4 prior-year docx memos must exist and open cleanly."""

    @pytest.mark.parametrize("filename", _PRIOR_YEAR_DOCX)
    def test_exists(self, filename: str) -> None:
        _, out, _, _, _ = _ensure_emitted()
        assert (out / _PY_DIR / filename).exists()

    @pytest.mark.parametrize("filename", _PRIOR_YEAR_DOCX)
    def test_opens_cleanly(self, filename: str) -> None:
        _, out, _, _, _ = _ensure_emitted()
        doc = Document(str(out / _PY_DIR / filename))
        assert len(doc.paragraphs) > 0


class TestCurrentYearFilesExist:
    """All 5 current-year data files must exist and be valid."""

    @pytest.mark.parametrize("filename,fmt", _CURRENT_YEAR_FILES)
    def test_exists(self, filename: str, fmt: str) -> None:
        _, out, _, _, _ = _ensure_emitted()
        assert (out / _CY_DIR / filename).exists()

    @pytest.mark.parametrize("filename,fmt", _CURRENT_YEAR_FILES)
    def test_opens_cleanly(self, filename: str, fmt: str) -> None:
        _, out, _, _, _ = _ensure_emitted()
        path = out / _CY_DIR / filename
        if fmt == "xlsx":
            wb = openpyxl.load_workbook(path, data_only=True)
            assert wb.sheetnames
        elif fmt == "docx":
            doc = Document(str(path))
            assert len(doc.paragraphs) > 0
        elif fmt == "csv":
            text = path.read_text()
            assert len(text) > 10


# ---------------------------------------------------------------------------
# Canary embedding
# ---------------------------------------------------------------------------


class TestCanaryEmbedding:
    """All 15 canary keys should be registered and findable in generated files."""

    def test_all_15_canary_keys_registered(self) -> None:
        _, _, canaries, _, _ = _ensure_emitted()
        for key in _CANARY_KEYS:
            assert key in canaries.entries, f"Missing canary key: {key}"

    def test_canary_values_are_8_chars(self) -> None:
        _, _, canaries, _, _ = _ensure_emitted()
        for key in _CANARY_KEYS:
            canary = canaries.entries[key].canary
            assert len(canary) == 8, f"Canary for {key} is {len(canary)} chars"

    def test_canaries_are_unique(self) -> None:
        _, _, canaries, _, _ = _ensure_emitted()
        values = [canaries.entries[k].canary for k in _CANARY_KEYS]
        assert len(values) == len(set(values)), "Duplicate canary values found"

    @pytest.mark.parametrize("filename", _PRIOR_YEAR_XLSX)
    def test_xlsx_canary_in_properties(self, filename: str) -> None:
        """Xlsx canaries are embedded in workbook description."""
        _, out, canaries, _, _ = _ensure_emitted()
        stem = filename.replace(".xlsx", "")
        key = f"tc18eu_{stem}"
        canary = canaries.entries[key].canary
        wb = openpyxl.load_workbook(out / _PY_DIR / filename, data_only=True)
        props_text = wb.properties.description or ""
        assert canary in props_text, f"Canary {canary} not in {filename} properties"

    @pytest.mark.parametrize("filename", _PRIOR_YEAR_DOCX)
    def test_docx_canary_in_comments(self, filename: str) -> None:
        """Docx canaries are embedded in core_properties.comments."""
        _, out, canaries, _, _ = _ensure_emitted()
        stem = filename.replace(".docx", "")
        key = f"tc18eu_{stem}"
        canary = canaries.entries[key].canary
        doc = Document(str(out / _PY_DIR / filename))
        comments = doc.core_properties.comments or ""
        assert canary in comments, f"Canary {canary} not in {filename} comments"

    @pytest.mark.parametrize("filename", ["trial_balance_fy2025.csv", "bank_statements_fy2025.csv"])
    def test_csv_canary_in_comment_line(self, filename: str) -> None:
        """CSV canaries are embedded as comment lines."""
        _, out, canaries, _, _ = _ensure_emitted()
        stem = filename.replace(".csv", "")
        key = f"tc18eu_cy_{stem}"
        canary = canaries.entries[key].canary
        text = (out / _CY_DIR / filename).read_text()
        assert canary in text, f"Canary {canary} not in {filename}"

    def test_cy_xlsx_canaries(self) -> None:
        """Current-year xlsx files have canaries in properties."""
        _, out, canaries, _, _ = _ensure_emitted()
        for filename, key_suffix in [
            ("lease_schedule_fy2025.xlsx", "cy_lease_schedule_fy2025"),
            ("goodwill_impairment_analysis_ifrs.xlsx", "cy_goodwill_impairment_ifrs_fy2025"),
        ]:
            key = f"tc18eu_{key_suffix}"
            canary = canaries.entries[key].canary
            wb = openpyxl.load_workbook(out / _CY_DIR / filename, data_only=True)
            props_text = wb.properties.description or ""
            assert canary in props_text, f"Canary {canary} not in {filename}"

    def test_cy_docx_canary(self) -> None:
        """Management projections docx has canary in comments."""
        _, out, canaries, _, _ = _ensure_emitted()
        key = "tc18eu_cy_mgmt_projections_fy2025"
        canary = canaries.entries[key].canary
        doc = Document(str(out / _CY_DIR / "management_projections_fy2025.docx"))
        comments = doc.core_properties.comments or ""
        assert canary in comments


# ---------------------------------------------------------------------------
# ERR-EU-018: stale Gewerbesteuer Hebesatz in fixed assets workpaper
# ---------------------------------------------------------------------------


class TestPlantedErrorEU018:
    """ERR-EU-018: stale Munich Gewerbesteuer Hebesatz of 480 instead of 490."""

    def test_error_registered(self) -> None:
        _, _, _, errors, _ = _ensure_emitted()
        assert "ERR-EU-018" in errors.entries, "ERR-EU-018 must be registered"

    def test_error_type_is_stale_data(self) -> None:
        _, _, _, errors, _ = _ensure_emitted()
        err = errors.entries["ERR-EU-018"]
        assert err.type == "stale_data"

    def test_error_file_points_to_fixed_assets(self) -> None:
        _, _, _, errors, _ = _ensure_emitted()
        err = errors.entries["ERR-EU-018"]
        assert "wp_fixed_assets_fy2024.xlsx" in err.file

    def test_error_catches_tc18_eu(self) -> None:
        _, _, _, errors, _ = _ensure_emitted()
        err = errors.entries["ERR-EU-018"]
        assert "TC-18-EU" in err.which_test_cases_should_catch

    def test_tax_rates_sheet_exists(self) -> None:
        _, out, _, _, _ = _ensure_emitted()
        wb = openpyxl.load_workbook(out / _PY_DIR / "wp_fixed_assets_fy2024.xlsx", data_only=True)
        assert "Tax Rates" in wb.sheetnames

    def test_stale_hebesatz_480_in_tax_rates(self) -> None:
        """Tax Rates sheet should have a row with Hebesatz 480 (stale value)."""
        _, out, _, _, _ = _ensure_emitted()
        wb = openpyxl.load_workbook(out / _PY_DIR / "wp_fixed_assets_fy2024.xlsx", data_only=True)
        ws = wb["Tax Rates"]
        found_stale = False
        for row in range(1, ws.max_row + 1):
            rate_val = ws.cell(row=row, column=4).value
            if rate_val and "480" in str(rate_val):
                tax_type = ws.cell(row=row, column=3).value or ""
                if "Gewerbesteuer" in str(tax_type) or "GewSt" in str(tax_type):
                    found_stale = True
                    break
        assert found_stale, "Tax Rates sheet should contain stale Hebesatz 480"


# ---------------------------------------------------------------------------
# Revenue workpaper content
# ---------------------------------------------------------------------------


class TestRevenueWorkpaper:
    """wp_revenue_fy2024.xlsx must have IFRS 15 revenue data."""

    def test_revenue_summary_sheet_exists(self) -> None:
        _, out, _, _, _ = _ensure_emitted()
        wb = openpyxl.load_workbook(out / _PY_DIR / "wp_revenue_fy2024.xlsx", data_only=True)
        assert "Revenue Summary" in wb.sheetnames

    def test_all_revenue_entities_present(self) -> None:
        """All entities from _REVENUE_FY2024 should appear in the workbook."""
        _, out, _, _, _ = _ensure_emitted()
        wb = openpyxl.load_workbook(out / _PY_DIR / "wp_revenue_fy2024.xlsx", data_only=True)
        ws = wb["Revenue Summary"]
        all_values: list[str] = []
        for row in range(1, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                val = ws.cell(row=row, column=col).value
                if val is not None:
                    all_values.append(str(val))
        text = " ".join(all_values)
        entities = {entity for entity, _, _ in _REVENUE_FY2024}
        for entity in entities:
            assert entity in text, f"Missing entity {entity} in revenue workpaper"

    def test_ifrs_15_reference(self) -> None:
        """Revenue workpaper should reference IFRS 15."""
        _, out, _, _, _ = _ensure_emitted()
        wb = openpyxl.load_workbook(out / _PY_DIR / "wp_revenue_fy2024.xlsx", data_only=True)
        ws = wb["Revenue Summary"]
        all_values = []
        for row in range(1, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                val = ws.cell(row=row, column=col).value
                if val:
                    all_values.append(str(val))
        text = " ".join(all_values)
        assert "IFRS 15" in text, "Revenue workpaper should reference IFRS 15"

    def test_no_error_in_revenue_file(self) -> None:
        _, _, _, errors, _ = _ensure_emitted()
        for err_id, err in errors.entries.items():
            assert "wp_revenue" not in err.file, (
                f"Unexpected error {err_id} in revenue file"
            )


# ---------------------------------------------------------------------------
# Lease schedule — FY2025
# ---------------------------------------------------------------------------


class TestLeaseSchedule:
    """FY2025 lease schedule should have 4 original + 3 new = 7 total leases."""

    def test_total_lease_count(self) -> None:
        _, out, _, _, _ = _ensure_emitted()
        wb = openpyxl.load_workbook(
            out / _CY_DIR / "lease_schedule_fy2025.xlsx", data_only=True,
        )
        ws = wb.active
        lease_ids: list[str] = []
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row=row, column=1).value
            if val and str(val).startswith("EU-LS-"):
                lease_ids.append(str(val))
        assert len(lease_ids) == 7, f"Expected 7 leases, got {len(lease_ids)}: {lease_ids}"

    @pytest.mark.parametrize("lease_id", ["EU-LS-005", "EU-LS-006", "EU-LS-007"])
    def test_new_lease_present(self, lease_id: str) -> None:
        _, out, _, _, _ = _ensure_emitted()
        wb = openpyxl.load_workbook(
            out / _CY_DIR / "lease_schedule_fy2025.xlsx", data_only=True,
        )
        ws = wb.active
        found = False
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row=row, column=1).value
            if val and str(val) == lease_id:
                found = True
                break
        assert found, f"New lease {lease_id} not found in FY2025 lease schedule"


# ---------------------------------------------------------------------------
# IAS 36 Goodwill impairment
# ---------------------------------------------------------------------------


class TestGoodwillImpairment:
    """goodwill_impairment_analysis_ifrs.xlsx must have CGU data."""

    def test_cd_cgu_headroom_3_5_pct(self) -> None:
        """Distribution Services (CD) CGU should have 3.5% headroom."""
        _, out, _, _, _ = _ensure_emitted()
        wb = openpyxl.load_workbook(
            out / _CY_DIR / "goodwill_impairment_analysis_ifrs.xlsx", data_only=True,
        )
        ws = wb.active
        found = False
        for row in range(2, ws.max_row + 1):
            for col in range(1, ws.max_column + 1):
                val = ws.cell(row=row, column=col).value
                if val and "3.5%" in str(val):
                    found = True
                    break
            if found:
                break
        assert found, "CD CGU should show 3.5% headroom"

    def test_total_goodwill_matches_constant(self) -> None:
        assert _TOTAL_GOODWILL == sum(c["carrying_amount"] for c in _GOODWILL_CGUS)
        assert _TOTAL_GOODWILL == Decimal("15_800_000")

    def test_three_cgus_present(self) -> None:
        """Should have at least 3 CGUs listed."""
        _, out, _, _, _ = _ensure_emitted()
        wb = openpyxl.load_workbook(
            out / _CY_DIR / "goodwill_impairment_analysis_ifrs.xlsx", data_only=True,
        )
        ws = wb.active
        cgus = []
        for row in range(2, ws.max_row + 1):
            val = ws.cell(row=row, column=1).value
            if val and "(" in str(val):
                cgus.append(str(val))
        assert len(cgus) >= 3, f"Expected at least 3 CGUs, found: {cgus}"


# ---------------------------------------------------------------------------
# IFRS/ISA terminology in memos
# ---------------------------------------------------------------------------


class TestMemoISATerminology:
    """Memos must contain correct IFRS/ISA references."""

    def _read_memo(self, filename: str) -> str:
        _, out, _, _, _ = _ensure_emitted()
        doc = Document(str(out / _PY_DIR / filename))
        return "\n".join(p.text for p in doc.paragraphs)

    def test_planning_memo_has_isa_315(self) -> None:
        text = self._read_memo("memo_planning_fy2024.docx")
        assert "ISA 315" in text

    def test_planning_memo_has_isa_330(self) -> None:
        text = self._read_memo("memo_planning_fy2024.docx")
        assert "ISA 330" in text

    def test_risk_assessment_has_isa_315(self) -> None:
        text = self._read_memo("memo_risk_assessment_fy2024.docx")
        assert "ISA 315" in text

    def test_summary_memo_has_isa_701(self) -> None:
        text = self._read_memo("memo_summary_fy2024.docx")
        assert "ISA 701" in text

    def test_summary_memo_has_isa_600(self) -> None:
        text = self._read_memo("memo_summary_fy2024.docx")
        assert "ISA 600" in text

    def test_management_letter_has_ifrs_content(self) -> None:
        text = self._read_memo("memo_management_letter_fy2024.docx")
        assert "IFRS" in text

    def test_planning_memo_has_ifrs_reference(self) -> None:
        text = self._read_memo("memo_planning_fy2024.docx")
        assert "IFRS" in text


# ---------------------------------------------------------------------------
# Gold standard
# ---------------------------------------------------------------------------


class TestGoldStandard:
    """TC-18-EU gold standard must be registered with correct structure."""

    def test_gold_json_exists(self) -> None:
        _, out, _, _, _ = _ensure_emitted()
        assert (out / "gold_standards/TC-18-EU_gold.json").exists()

    def test_gold_has_new_leases(self) -> None:
        _, out, _, _, _ = _ensure_emitted()
        gold = json.loads((out / "gold_standards/TC-18-EU_gold.json").read_text())
        new_leases = gold["expected_outputs"]["new_leases"]
        assert new_leases == ["EU-LS-005", "EU-LS-006", "EU-LS-007"]

    def test_gold_has_error_detection(self) -> None:
        _, out, _, _, _ = _ensure_emitted()
        gold = json.loads((out / "gold_standards/TC-18-EU_gold.json").read_text())
        assert "ERR-EU-018" in gold["error_detection"]

    def test_gold_has_8_mechanical_updates(self) -> None:
        _, out, _, _, _ = _ensure_emitted()
        gold = json.loads((out / "gold_standards/TC-18-EU_gold.json").read_text())
        mechanical = gold["expected_outputs"]["files_to_update"]["mechanical_updates"]
        assert len(mechanical) == 8

    def test_gold_has_2_requires_manager_judgment(self) -> None:
        _, out, _, _, _ = _ensure_emitted()
        gold = json.loads((out / "gold_standards/TC-18-EU_gold.json").read_text())
        judgment = gold["expected_outputs"]["files_to_update"]["requires_manager_judgment"]
        assert len(judgment) == 2

    def test_gold_has_15_canary_verifications(self) -> None:
        _, out, _, _, _ = _ensure_emitted()
        gold = json.loads((out / "gold_standards/TC-18-EU_gold.json").read_text())
        assert len(gold["canary_verification"]) == 15

    def test_gold_has_total_goodwill(self) -> None:
        _, out, _, _, _ = _ensure_emitted()
        gold = json.loads((out / "gold_standards/TC-18-EU_gold.json").read_text())
        new_audit = gold["expected_outputs"]["new_audit_area"]
        assert new_audit["total_goodwill"] == int(_TOTAL_GOODWILL)


# ---------------------------------------------------------------------------
# Prompt and expected behavior
# ---------------------------------------------------------------------------


class TestPromptAndExpectedBehavior:
    """TC-18-EU must have prompt.md and expected_behavior.md with key phrases."""

    def test_prompt_exists(self) -> None:
        _, out, _, _, _ = _ensure_emitted()
        path = out / "test_cases/TC-18-EU/prompt.md"
        assert path.exists()

    def test_prompt_has_ifrs(self) -> None:
        _, out, _, _, _ = _ensure_emitted()
        text = (out / "test_cases/TC-18-EU/prompt.md").read_text()
        assert "IFRS" in text

    def test_prompt_has_isa(self) -> None:
        _, out, _, _, _ = _ensure_emitted()
        text = (out / "test_cases/TC-18-EU/prompt.md").read_text()
        assert "ISA" in text

    def test_prompt_has_3_new_leases(self) -> None:
        _, out, _, _, _ = _ensure_emitted()
        text = (out / "test_cases/TC-18-EU/prompt.md").read_text()
        assert "3 new leases" in text

    def test_expected_behavior_exists(self) -> None:
        _, out, _, _, _ = _ensure_emitted()
        path = out / "test_cases/TC-18-EU/expected_behavior.md"
        assert path.exists()
        text = path.read_text()
        assert len(text) > 100

    def test_expected_behavior_has_ifrs(self) -> None:
        _, out, _, _, _ = _ensure_emitted()
        text = (out / "test_cases/TC-18-EU/expected_behavior.md").read_text()
        assert "IFRS" in text

    def test_expected_behavior_has_isa(self) -> None:
        _, out, _, _, _ = _ensure_emitted()
        text = (out / "test_cases/TC-18-EU/expected_behavior.md").read_text()
        assert "ISA" in text

    def test_expected_behavior_has_3_new_leases(self) -> None:
        _, out, _, _, _ = _ensure_emitted()
        text = (out / "test_cases/TC-18-EU/expected_behavior.md").read_text()
        assert "3 new leases" in text

    def test_expected_behavior_has_gewerbesteuer_context(self) -> None:
        """Expected behavior references IAS 16 revaluation (context for fixed assets)."""
        _, out, _, _, _ = _ensure_emitted()
        text = (out / "test_cases/TC-18-EU/expected_behavior.md").read_text()
        assert "revaluation" in text.lower()


# ---------------------------------------------------------------------------
# Manifest registration
# ---------------------------------------------------------------------------


class TestManifestRegistration:
    """All TC-18-EU files must appear in the manifest."""

    def test_manifest_has_tc18_eu_entries(self) -> None:
        _, out, _, _, _ = _ensure_emitted()
        manifest_data = json.loads((out / "manifest.json").read_text())
        tc18_entries = [e for e in manifest_data if "TC-18-EU" in e.get("test_cases", [])]
        # 6 xlsx + 4 docx + 5 current-year = 15 files
        assert len(tc18_entries) >= 15, (
            f"Expected at least 15 TC-18-EU manifest entries, got {len(tc18_entries)}"
        )
