"""Tests for the validation panel widget and its integration with screens.

Covers: error grouping, field-path extraction, panel state transitions,
and the acceptance-criteria flow: invalid draft → error → corrected → valid.

Bead: synth-data-2u6.6.8
"""

from __future__ import annotations

import textwrap
from pathlib import Path

import pytest

from generator.tui.draft_service import DraftService, ValidationResult
from generator.tui.validation_panel import (
    _extract_field_path,
    _group_errors,
)

# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


def _base_config_yaml() -> str:
    """Minimal valid base config."""
    return textwrap.dedent("""\
        seed: 42
        output_dir: out
        company:
          name: TestCo
          type: C-Corp
          industry: Manufacturing
          headquarters: Portland, OR
          fiscal_year_end: "12-31"
          years: [2025]
          current_year: 2025
          consolidated_revenue: 100000000
          subsidiaries:
            sub_a:
              legal_name: Sub A LLC
              location: Portland, OR
              state: OR
              entity_code: SA
              revenue: 100000000
              type: Manufacturing
              gross_margin: 0.35
              employee_count: 100
          growth_rates:
            fy2023_to_fy2024: 0.06
            fy2024_to_fy2025: 0.09
          intercompany:
            raw_materials_markup: 0.08
            management_fee_pct: 0.015
            intercompany_loan_principal: 5000000
            intercompany_loan_rate: 0.05
          employees:
            total_count: 100
            annual_turnover_rate: 0.08
            remote_states: [CA]
          seasonal_weights:
            Q1: 0.20
            Q2: 0.25
            Q3: 0.25
            Q4: 0.30
    """)


@pytest.fixture()
def base_yaml(tmp_path: Path) -> Path:
    p = tmp_path / "config.yaml"
    p.write_text(_base_config_yaml())
    return p


@pytest.fixture()
def service(base_yaml: Path) -> DraftService:
    return DraftService(base_yaml)


# ---------------------------------------------------------------------------
# _extract_field_path tests
# ---------------------------------------------------------------------------


class TestExtractFieldPath:
    def test_simple_section(self) -> None:
        assert _extract_field_path("company must have a name") == "company"

    def test_dotted_path(self) -> None:
        assert (
            _extract_field_path("difficulty.error_density must be 0.0–1.0, got 2.0")
            == "difficulty.error_density"
        )

    def test_nested_path(self) -> None:
        assert (
            _extract_field_path("Unknown keys in company.subsidiaries.US: ['bogus']")
            == "company.subsidiaries.US"
        )

    def test_no_match(self) -> None:
        assert _extract_field_path("Something went wrong") == ""

    def test_output_section(self) -> None:
        assert _extract_field_path("output.formats is invalid") == "output.formats"


# ---------------------------------------------------------------------------
# _group_errors tests
# ---------------------------------------------------------------------------


class TestGroupErrors:
    def test_empty_errors(self) -> None:
        assert _group_errors([]) == {}

    def test_single_company_error(self) -> None:
        groups = _group_errors(["company.name is required"])
        assert "company" in groups
        assert len(groups["company"]) == 1
        assert groups["company"][0].field_path == "company.name"

    def test_general_bucket(self) -> None:
        groups = _group_errors(["Something unknown happened"])
        assert "general" in groups
        assert groups["general"][0].field_path == ""

    def test_multiple_sections(self) -> None:
        groups = _group_errors([
            "company.name is required",
            "difficulty.error_density must be 0.0–1.0",
            "output.formats is invalid",
        ])
        assert set(groups.keys()) == {"company", "difficulty", "output"}

    def test_multiple_errors_same_section(self) -> None:
        groups = _group_errors([
            "company.name is required",
            "company.type is invalid",
        ])
        assert len(groups["company"]) == 2


# ---------------------------------------------------------------------------
# Acceptance criteria: invalid → error → corrected → valid
# ---------------------------------------------------------------------------


class TestInvalidToValidFlow:
    """End-to-end flow using DraftService validation results
    piped through ValidationPanel's data model."""

    def test_base_config_is_valid(self, service: DraftService) -> None:
        result = service.validate()
        assert result.valid
        assert result.errors == []

    def test_invalid_seasonal_weights_produce_errors(
        self, service: DraftService
    ) -> None:
        # Break the config: seasonal weights must sum to 1.0
        service.set_field("company.seasonal_weights.Q1", 0.99)
        result = service.validate()
        assert not result.valid
        assert len(result.errors) > 0

        # Errors should be groupable
        groups = _group_errors(result.errors)
        assert len(groups) > 0

    def test_corrected_draft_becomes_valid(self, service: DraftService) -> None:
        # 1. Break the config
        service.set_field("company.seasonal_weights.Q1", 0.99)
        bad_result = service.validate()
        assert not bad_result.valid

        # 2. Fix: reset the broken field back to base
        service.reset_field("company.seasonal_weights.Q1")
        good_result = service.validate()
        assert good_result.valid
        assert good_result.errors == []

    def test_invalid_choice_then_valid(self, service: DraftService) -> None:
        # Break via an invalid choice at the service level
        with pytest.raises(ValueError, match="Invalid choice"):
            service.set_field("difficulty.canary_visibility", "invisible")

        # Config hasn't changed — still valid
        result = service.validate()
        assert result.valid

    def test_full_cycle_multiple_errors(self, service: DraftService) -> None:
        """Introduce multiple errors, verify they're grouped, then fix."""
        # Break: seasonal weights out of balance
        service.set_field("company.seasonal_weights.Q1", 0.80)
        result = service.validate()
        assert not result.valid

        groups = _group_errors(result.errors)
        # Should have at least one error group
        total_errors = sum(len(v) for v in groups.values())
        assert total_errors >= 1

        # Fix all overrides
        service.reset_all()
        result = service.validate()
        assert result.valid


class TestValidationResultIntegration:
    """Verify that ValidationResult objects work correctly with panel helpers."""

    def test_valid_result_has_no_errors(self) -> None:
        result = ValidationResult(valid=True)
        groups = _group_errors(result.errors)
        assert groups == {}

    def test_invalid_result_groups_correctly(self) -> None:
        result = ValidationResult(
            valid=False,
            errors=[
                "company.name is required",
                "difficulty.error_density must be 0.0–1.0, got 2.0",
                "Unknown format specification",
            ],
        )
        groups = _group_errors(result.errors)
        assert "company" in groups
        assert "difficulty" in groups
        assert "general" in groups
