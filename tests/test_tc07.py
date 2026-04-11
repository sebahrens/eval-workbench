"""Tests for TC-07 — K-1 Extraction & Consolidation (Tax, Adversarial) formatter.

Verifies:
- 8 K-1 PDFs in k1s/ (3 system-clean, 5 varying layouts)
- entity_org_chart.pdf
- ERR-014 planted error (wrong entity name in K1-003)
- Amended K-1 detection (K1-004: ordinary income $340K→$285K, +$55K GP)
- Canary embedding in all PDF files
- Gold standard structure, per-K-1 detail, consolidated totals, scoring hints
- Prompt and expected behavior markdown files
"""

from __future__ import annotations

import json
import tempfile
from pathlib import Path

from generator.canaries import CanaryRegistry
from generator.canaries import build_registry as build_canary_registry
from generator.errors import ErrorRegistry
from generator.formatters.tc07 import (
    _ALL_CANARY_KEYS,
    _FORM_1120_MAPPING,
    emit_tc07,
)
from generator.golds.framework import emit_gold
from generator.manifest import Manifest
from generator.model.build import CascadeModel, build_model
from generator.model.k1 import (
    K1LayoutType,
    generate_k1_investments,
)

# ---------------------------------------------------------------------------
# Module-level fixture: build the model and run emit_tc07 once
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
        _OUTPUT = Path(tempfile.mkdtemp(prefix="tc07_test_"))
        _CANARIES = build_canary_registry(_ALL_CANARY_KEYS, seed=42)
        _ERRORS = ErrorRegistry()
        _MANIFEST = Manifest(_OUTPUT)
        _MANIFEST.__enter__()

        emit_tc07(_MODEL, _OUTPUT, _CANARIES, _ERRORS, _MANIFEST)

        # Emit gold standard
        emit_gold("TC-07", _CANARIES, _ERRORS, _OUTPUT / "gold_standards", model=_MODEL)

        _MANIFEST.__exit__(None, None, None)
    return _MODEL, _OUTPUT, _CANARIES, _ERRORS, _MANIFEST


_INPUT_DIR = "test_cases/TC-07/input_files"
_K1S_DIR = f"{_INPUT_DIR}/k1s"


# ---------------------------------------------------------------------------
# K-1 PDF files — existence and count
# ---------------------------------------------------------------------------


class TestK1PDFs:
    """Verify 8 K-1 PDFs are emitted in k1s/ subdirectory."""

    def test_k1s_directory_exists(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        assert (output / _K1S_DIR).is_dir()

    def test_8_k1_pdfs_exist(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        for i in range(1, 9):
            path = output / _K1S_DIR / f"K1-{i:03d}.pdf"
            assert path.exists(), f"Missing K1-{i:03d}.pdf"

    def test_no_extra_k1_pdfs(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        pdfs = list((output / _K1S_DIR).glob("K1-*.pdf"))
        assert len(pdfs) == 8

    def test_3_clean_5_varying(self) -> None:
        """K-1 layout types match: 3 system-clean, 5 varying."""
        investments = generate_k1_investments()
        clean = [k for k in investments if k.layout_type == K1LayoutType.SYSTEM_CLEAN]
        varying = [k for k in investments if k.layout_type == K1LayoutType.VARYING]
        assert len(clean) == 3
        assert len(varying) == 5


# ---------------------------------------------------------------------------
# Entity Org Chart PDF
# ---------------------------------------------------------------------------


class TestEntityOrgChart:
    """Verify entity_org_chart.pdf is emitted."""

    def test_file_exists(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / f"{_INPUT_DIR}/entity_org_chart.pdf"
        assert path.exists()

    def test_file_is_pdf(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / f"{_INPUT_DIR}/entity_org_chart.pdf"
        header = path.read_bytes()[:5]
        assert header == b"%PDF-"


# ---------------------------------------------------------------------------
# ERR-014 — Wrong entity name in K1-003
# ---------------------------------------------------------------------------


class TestERR014PlantedError:
    """Verify ERR-014: K1-003 shows wrong entity name."""

    def test_err014_registered(self) -> None:
        _, _, _, errors, _ = _ensure_emitted()
        assert "ERR-014" in errors.entries

    def test_err014_type(self) -> None:
        _, _, _, errors, _ = _ensure_emitted()
        err = errors.entries["ERR-014"]
        assert err.type == "wrong_entity"

    def test_err014_references_k1_003(self) -> None:
        _, _, _, errors, _ = _ensure_emitted()
        err = errors.entries["ERR-014"]
        assert "K1-003" in err.file

    def test_err014_wrong_vs_correct_name(self) -> None:
        _, _, _, errors, _ = _ensure_emitted()
        err = errors.entries["ERR-014"]
        assert "Cascade Precision Components LLC" in err.description
        assert "Cascade Advanced Materials, Inc." in err.description

    def test_err014_location_is_partner_name(self) -> None:
        _, _, _, errors, _ = _ensure_emitted()
        err = errors.entries["ERR-014"]
        assert "Partner" in err.location


# ---------------------------------------------------------------------------
# Amended K-1 detection
# ---------------------------------------------------------------------------


class TestAmendedK1:
    """Verify the amended K-1 (K1-004) is generated correctly."""

    def test_exactly_one_amended(self) -> None:
        investments = generate_k1_investments()
        amended = [k for k in investments if k.is_amended]
        assert len(amended) == 1

    def test_amended_is_k1_004(self) -> None:
        investments = generate_k1_investments()
        amended = [k for k in investments if k.is_amended][0]
        assert amended.k1_id == "K1-004"

    def test_amended_ordinary_income_285k(self) -> None:
        from decimal import Decimal
        investments = generate_k1_investments()
        amended = [k for k in investments if k.is_amended][0]
        assert amended.box_1_ordinary_income == Decimal("285000")

    def test_amended_guaranteed_payments_55k(self) -> None:
        from decimal import Decimal
        investments = generate_k1_investments()
        amended = [k for k in investments if k.is_amended][0]
        assert amended.box_4c_total_guaranteed_payments == Decimal("55000")

    def test_amendment_records_present(self) -> None:
        investments = generate_k1_investments()
        amended = [k for k in investments if k.is_amended][0]
        assert len(amended.amendments) == 2


# ---------------------------------------------------------------------------
# Canary embedding
# ---------------------------------------------------------------------------


class TestCanaryEmbedding:
    """Verify canary codes are embedded in all TC-07 files."""

    def test_all_canary_keys_assigned(self) -> None:
        _, _, canaries, _, _ = _ensure_emitted()
        for key in _ALL_CANARY_KEYS:
            code = canaries.canary_for(key)
            assert len(code) == 8, f"Canary for {key} should be 8 chars"

    def test_k1_pdf_canaries_in_metadata(self) -> None:
        """Each K-1 PDF should have its canary in the PDF metadata."""
        _, output, canaries, _, _ = _ensure_emitted()
        for i in range(1, 9):
            k1_id = f"K1-{i:03d}"
            key = f"tc07_k1_{i:03d}"
            canary = canaries.canary_for(key)
            path = output / _K1S_DIR / f"{k1_id}.pdf"
            content = path.read_bytes()
            assert canary.encode() in content, (
                f"Canary {canary} not found in {k1_id}.pdf"
            )

    def test_org_chart_canary_in_metadata(self) -> None:
        _, output, canaries, _, _ = _ensure_emitted()
        canary = canaries.canary_for("tc07_entity_org_chart")
        path = output / f"{_INPUT_DIR}/entity_org_chart.pdf"
        content = path.read_bytes()
        assert canary.encode() in content, (
            f"Canary {canary} not found in entity_org_chart.pdf"
        )


# ---------------------------------------------------------------------------
# Gold standard
# ---------------------------------------------------------------------------


class TestGoldStandard:
    """Verify gold standard JSON structure."""

    def test_gold_json_exists(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "gold_standards" / "TC-07_gold.json"
        assert path.exists()

    def test_gold_has_expected_outputs(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        gold = json.loads((output / "gold_standards" / "TC-07_gold.json").read_text())
        eo = gold["expected_outputs"]
        assert "k1_count" in eo
        assert eo["k1_count"] == 8

    def test_gold_has_k1_details(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        gold = json.loads((output / "gold_standards" / "TC-07_gold.json").read_text())
        details = gold["expected_outputs"]["k1_details"]
        assert len(details) == 8
        for i in range(1, 9):
            assert f"K1-{i:03d}" in details

    def test_gold_has_consolidated_totals(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        gold = json.loads((output / "gold_standards" / "TC-07_gold.json").read_text())
        totals = gold["expected_outputs"]["consolidated_totals"]
        assert len(totals) > 0

    def test_gold_has_form_1120_mapping(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        gold = json.loads((output / "gold_standards" / "TC-07_gold.json").read_text())
        mapping = gold["expected_outputs"]["form_1120_mapping"]
        assert len(mapping) == len(_FORM_1120_MAPPING)

    def test_gold_has_amended_k1_info(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        gold = json.loads((output / "gold_standards" / "TC-07_gold.json").read_text())
        amended = gold["expected_outputs"]["amended_k1"]
        assert amended["k1_id"] == "K1-004"
        assert len(amended["changes"]) == 2

    def test_gold_has_section_199a_flag(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        gold = json.loads((output / "gold_standards" / "TC-07_gold.json").read_text())
        flag = gold["expected_outputs"]["section_199a_flag"]
        assert "C-corporation" in flag
        assert "NOT applicable" in flag

    def test_gold_has_special_handling(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        gold = json.loads((output / "gold_standards" / "TC-07_gold.json").read_text())
        handling = gold["expected_outputs"]["special_handling"]
        assert len(handling) >= 4

    def test_gold_canary_verification(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        gold = json.loads((output / "gold_standards" / "TC-07_gold.json").read_text())
        cv = gold["canary_verification"]
        # One per K-1 + org chart = 9 entries
        assert len(cv) == 9
        assert "read_org_chart" in cv

    def test_gold_error_detection(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        gold = json.loads((output / "gold_standards" / "TC-07_gold.json").read_text())
        assert "ERR-014" in gold["error_detection"]

    def test_gold_scoring_hints(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        gold = json.loads((output / "gold_standards" / "TC-07_gold.json").read_text())
        hints = gold["scoring_hints"]
        for key in ["correctness", "completeness", "format_compliance", "communication"]:
            assert key in hints


# ---------------------------------------------------------------------------
# Prompt and expected behavior
# ---------------------------------------------------------------------------


class TestPromptAndExpectedBehavior:
    """Verify prompt and expected behavior files are generated."""

    def test_prompt_md_exists(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "test_cases/TC-07/prompt.md"
        assert path.exists()

    def test_prompt_mentions_k1s(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        text = (output / "test_cases/TC-07/prompt.md").read_text()
        assert "K-1" in text
        assert "8 partnership" in text.lower() or "8 Schedule K-1" in text

    def test_prompt_mentions_consolidation(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        text = (output / "test_cases/TC-07/prompt.md").read_text().lower()
        assert "consolidat" in text

    def test_prompt_mentions_amended(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        text = (output / "test_cases/TC-07/prompt.md").read_text().lower()
        assert "amended" in text

    def test_prompt_mentions_section_199a(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        text = (output / "test_cases/TC-07/prompt.md").read_text()
        assert "199A" in text

    def test_expected_behavior_exists(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "test_cases/TC-07/expected_behavior.md"
        assert path.exists()

    def test_expected_behavior_mentions_amended_k1(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        text = (output / "test_cases/TC-07/expected_behavior.md").read_text()
        assert "K1-004" in text
        assert "amended" in text.lower()

    def test_expected_behavior_mentions_c_corp(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        text = (output / "test_cases/TC-07/expected_behavior.md").read_text()
        assert "C-corporation" in text or "C-corp" in text

    def test_expected_behavior_mentions_form_1120(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        text = (output / "test_cases/TC-07/expected_behavior.md").read_text()
        assert "Form 1120" in text


# ---------------------------------------------------------------------------
# Manifest registration
# ---------------------------------------------------------------------------


class TestManifest:
    """Verify files are registered in the manifest."""

    def test_manifest_has_tc07_entries(self) -> None:
        _, _, _, _, manifest = _ensure_emitted()
        tc07_count = sum(
            1 for v in manifest.entries.values()
            if "TC-07" in (v.test_cases or [])
        )
        # 8 K-1 PDFs + 1 org chart = 9
        assert tc07_count >= 9, f"Expected ≥9 TC-07 manifest entries, got {tc07_count}"
