"""Tests for augmentation output validators.

Acceptance criteria from synth-data-8ok.7.3:
- Validators fail with field-level messages
- Required canonical names/dates/amounts checked
- Contradictions against canonical model facts detected
- Canary and planted-error preservation validated
- Usable by DataDesigner or local augmentation pipeline
"""

from __future__ import annotations

from pathlib import Path

import pytest

from generator.augmentation_validators import (
    CanonicalFact,
    ValidationMessage,
    ValidationResult,
    extract_canonical_facts,
    validate_augmentation_output,
    validate_canary_preserved,
    validate_no_contradictions,
    validate_planted_error_preserved,
    validate_required_facts,
)
from generator.config import load_config

# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture()
def config():
    root = Path(__file__).resolve().parent.parent / "config.yaml"
    return load_config(root)


@pytest.fixture()
def facts(config):
    return extract_canonical_facts(config)


# ---------------------------------------------------------------------------
# ValidationResult
# ---------------------------------------------------------------------------

class TestValidationResult:
    def test_ok_when_empty(self) -> None:
        r = ValidationResult()
        assert r.ok is True
        assert r.errors == []
        assert r.warnings == []

    def test_ok_with_warnings_only(self) -> None:
        r = ValidationResult(messages=[
            ValidationMessage("warning", "field.a", "just a heads up"),
        ])
        assert r.ok is True

    def test_not_ok_with_errors(self) -> None:
        r = ValidationResult(messages=[
            ValidationMessage("error", "field.a", "bad"),
        ])
        assert r.ok is False

    def test_merge(self) -> None:
        r1 = ValidationResult(messages=[
            ValidationMessage("error", "a", "msg1"),
        ])
        r2 = ValidationResult(messages=[
            ValidationMessage("warning", "b", "msg2"),
        ])
        r1.merge(r2)
        assert len(r1.messages) == 2

    def test_str_format(self) -> None:
        m = ValidationMessage("error", "company.name", "missing")
        assert "[ERROR]" in str(m)
        assert "company.name" in str(m)


# ---------------------------------------------------------------------------
# extract_canonical_facts
# ---------------------------------------------------------------------------

class TestExtractCanonicalFacts:
    def test_includes_company_name(self, facts: list[CanonicalFact]) -> None:
        names = [f for f in facts if f.field == "company.name"]
        assert len(names) == 1
        assert "Cascade" in names[0].value

    def test_includes_subsidiary_names(self, facts: list[CanonicalFact]) -> None:
        sub_facts = [f for f in facts if f.field.startswith("subsidiary.") and f.field.endswith(".legal_name")]
        assert len(sub_facts) > 0

    def test_includes_financial_figures(self, facts: list[CanonicalFact]) -> None:
        rev = [f for f in facts if f.field == "company.consolidated_revenue"]
        assert len(rev) == 1
        assert rev[0].value == "200000000"

    def test_company_name_is_required(self, facts: list[CanonicalFact]) -> None:
        name_fact = next(f for f in facts if f.field == "company.name")
        assert name_fact.required is True

    def test_subsidiary_names_not_required(self, facts: list[CanonicalFact]) -> None:
        sub_facts = [f for f in facts if f.field.startswith("subsidiary.") and f.field.endswith(".legal_name")]
        for f in sub_facts:
            assert f.required is False


# ---------------------------------------------------------------------------
# validate_required_facts
# ---------------------------------------------------------------------------

class TestValidateRequiredFacts:
    def test_passes_when_all_present(self, facts: list[CanonicalFact]) -> None:
        text = (
            "Cascade Industries, Inc. is a mid-market manufacturer "
            "headquartered in Portland, Oregon with fiscal year ending 12-31."
        )
        result = validate_required_facts(text, facts)
        assert result.ok

    def test_fails_when_company_name_missing(self, facts: list[CanonicalFact]) -> None:
        text = "Some unknown company headquartered in Portland, Oregon."
        result = validate_required_facts(text, facts)
        assert not result.ok
        error_fields = [m.field for m in result.errors]
        assert "company.name" in error_fields

    def test_case_insensitive_for_names(self) -> None:
        facts = [CanonicalFact("company.name", "Cascade Industries, Inc.")]
        text = "cascade industries, inc. is great"
        result = validate_required_facts(text, facts)
        assert result.ok

    def test_skips_non_required_facts(self) -> None:
        facts = [CanonicalFact("subsidiary.PC.legal_name", "Something LLC", required=False)]
        text = "No mention of subsidiaries here."
        result = validate_required_facts(text, facts)
        assert result.ok  # Non-required facts don't cause errors


# ---------------------------------------------------------------------------
# validate_no_contradictions
# ---------------------------------------------------------------------------

class TestValidateNoContradictions:
    def test_passes_when_amounts_match(self) -> None:
        facts = [CanonicalFact("company.consolidated_revenue", "200000000", required=False)]
        text = "The company reported revenue of $200,000,000 this year."
        result = validate_no_contradictions(text, facts)
        assert result.ok

    def test_fails_when_amount_contradicts(self) -> None:
        facts = [CanonicalFact("company.consolidated_revenue", "200000000", required=False)]
        text = "The company reported revenue of $300,000,000 this year."
        result = validate_no_contradictions(text, facts)
        assert not result.ok
        assert any("300,000,000" in m.message for m in result.errors)

    def test_passes_when_no_amounts_mentioned(self) -> None:
        facts = [CanonicalFact("company.consolidated_revenue", "200000000", required=False)]
        text = "The company is a leading manufacturer."
        result = validate_no_contradictions(text, facts)
        assert result.ok

    def test_detects_subsidiary_revenue_contradiction(self) -> None:
        facts = [CanonicalFact("subsidiary.PC.revenue", "95000000", required=False)]
        text = "PC division generated $50,000,000 in revenue."
        result = validate_no_contradictions(text, facts)
        assert not result.ok

    def test_ignores_unrelated_numbers(self) -> None:
        facts = [CanonicalFact("subsidiary.PC.revenue", "95000000", required=False)]
        # A small number not near PC context shouldn't trigger
        text = "The office has 350 employees."
        result = validate_no_contradictions(text, facts)
        assert result.ok


# ---------------------------------------------------------------------------
# validate_canary_preserved
# ---------------------------------------------------------------------------

class TestValidateCanaryPreserved:
    def test_passes_when_canary_present(self) -> None:
        result = validate_canary_preserved("Report abc12345 data", "abc12345", "file_1")
        assert result.ok

    def test_fails_when_canary_missing(self) -> None:
        result = validate_canary_preserved("Report without marker", "abc12345", "file_1")
        assert not result.ok
        assert result.errors[0].field == "canary.file_1"

    def test_empty_canary_passes(self) -> None:
        result = validate_canary_preserved("anything", "", "file_1")
        assert result.ok


# ---------------------------------------------------------------------------
# validate_planted_error_preserved
# ---------------------------------------------------------------------------

class TestValidatePlantedErrorPreserved:
    def test_passes_when_error_present(self) -> None:
        result = validate_planted_error_preserved(
            "Total: $1,234,567 (transposed)", "ERR-001", "$1,234,567",
        )
        assert result.ok

    def test_fails_when_error_masked(self) -> None:
        result = validate_planted_error_preserved(
            "Total: $1,234,567 all correct", "ERR-001", "transposed",
        )
        assert not result.ok
        assert "ERR-001" in result.errors[0].message

    def test_empty_error_passes(self) -> None:
        result = validate_planted_error_preserved("anything", "", "")
        assert result.ok


# ---------------------------------------------------------------------------
# validate_augmentation_output (combined)
# ---------------------------------------------------------------------------

class TestValidateAugmentationOutput:
    def test_all_pass(self, config) -> None:
        text = (
            "Cascade Industries, Inc. is a mid-market manufacturer "
            "headquartered in Portland, Oregon with fiscal year ending 12-31. "
            "Canary: XY12AB34"
        )
        result = validate_augmentation_output(
            text, config, canary="XY12AB34", file_id="test_file",
        )
        assert result.ok

    def test_multiple_failures(self, config) -> None:
        text = "Some company with revenue of $999,999,999. Missing canary."
        result = validate_augmentation_output(
            text, config,
            canary="XY12AB34", file_id="test_file",
            error_id="ERR-001", error_marker="transposed",
        )
        assert not result.ok
        # Should have errors for: missing company name, missing canary, missing error
        assert len(result.errors) >= 3

    def test_field_level_messages(self, config) -> None:
        text = "A document about Portland, Oregon with fiscal year ending 12-31."
        result = validate_augmentation_output(text, config)
        # Missing company name should produce field-specific error
        name_errors = [m for m in result.errors if m.field == "company.name"]
        assert len(name_errors) == 1
        assert "Cascade Industries" in name_errors[0].message
