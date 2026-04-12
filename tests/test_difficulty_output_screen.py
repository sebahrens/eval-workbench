"""Tests for the difficulty and output profile configuration screen.

Covers: valid and invalid profile updates via the DraftService,
field grouping, widget composition, and unsupported-field rejection.

Bead: synth-data-2u6.6.6
"""

from __future__ import annotations

import textwrap
from pathlib import Path

import pytest

from generator.schema_metadata import FieldGroup, InputType, get_fields_by_group
from generator.tui.draft_service import DraftService

# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


def _base_config_yaml() -> str:
    """Minimal valid base config with difficulty/output sections."""
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
        difficulty:
          error_density: 1.0
          canary_visibility: visible
          judgment_trap_density: 1.0
        output:
          enabled_test_cases: []
          enabled_packs: []
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
# Difficulty field metadata tests
# ---------------------------------------------------------------------------


class TestDifficultyFieldMetadata:
    """Verify difficulty group fields are registered correctly."""

    def test_difficulty_fields_count(self) -> None:
        fields = get_fields_by_group(FieldGroup.DIFFICULTY)
        assert len(fields) == 3

    def test_error_density_meta(self) -> None:
        fields = get_fields_by_group(FieldGroup.DIFFICULTY)
        by_path = {f.path: f for f in fields}
        f = by_path["difficulty.error_density"]
        assert f.input_type == InputType.FLOAT
        assert f.range_min == 0.0
        assert f.range_max == 1.0
        assert not f.required

    def test_canary_visibility_meta(self) -> None:
        fields = get_fields_by_group(FieldGroup.DIFFICULTY)
        by_path = {f.path: f for f in fields}
        f = by_path["difficulty.canary_visibility"]
        assert f.input_type == InputType.CHOICE
        assert set(f.choices) == {"visible", "subtle", "hidden"}

    def test_judgment_trap_density_meta(self) -> None:
        fields = get_fields_by_group(FieldGroup.DIFFICULTY)
        by_path = {f.path: f for f in fields}
        f = by_path["difficulty.judgment_trap_density"]
        assert f.input_type == InputType.FLOAT
        assert f.range_min == 0.0
        assert f.range_max == 1.0


# ---------------------------------------------------------------------------
# Output field metadata tests
# ---------------------------------------------------------------------------


class TestOutputFieldMetadata:
    """Verify output group fields are registered correctly."""

    def test_output_fields_count(self) -> None:
        fields = get_fields_by_group(FieldGroup.OUTPUT)
        assert len(fields) == 2

    def test_enabled_test_cases_meta(self) -> None:
        fields = get_fields_by_group(FieldGroup.OUTPUT)
        by_path = {f.path: f for f in fields}
        f = by_path["output.enabled_test_cases"]
        assert f.input_type == InputType.MULTI_CHOICE
        assert len(f.choices) == 18
        assert f.choices[0] == "TC-01"
        assert f.choices[-1] == "TC-18"

    def test_enabled_packs_meta(self) -> None:
        fields = get_fields_by_group(FieldGroup.OUTPUT)
        by_path = {f.path: f for f in fields}
        f = by_path["output.enabled_packs"]
        assert f.input_type == InputType.LIST_TEXT
        assert not f.required


# ---------------------------------------------------------------------------
# Valid difficulty profile updates
# ---------------------------------------------------------------------------


class TestValidDifficultyUpdates:
    """Service accepts valid difficulty field changes."""

    def test_set_error_density(self, service: DraftService) -> None:
        service.set_field("difficulty.error_density", 0.5)
        assert service.get_field_value("difficulty.error_density") == 0.5
        assert service.is_modified("difficulty.error_density")

    def test_set_error_density_from_string(self, service: DraftService) -> None:
        service.set_field("difficulty.error_density", "0.3")
        assert service.get_field_value("difficulty.error_density") == 0.3

    def test_set_canary_visibility_subtle(self, service: DraftService) -> None:
        service.set_field("difficulty.canary_visibility", "subtle")
        assert service.get_field_value("difficulty.canary_visibility") == "subtle"

    def test_set_canary_visibility_hidden(self, service: DraftService) -> None:
        service.set_field("difficulty.canary_visibility", "hidden")
        assert service.get_field_value("difficulty.canary_visibility") == "hidden"

    def test_set_judgment_trap_density_zero(self, service: DraftService) -> None:
        service.set_field("difficulty.judgment_trap_density", 0.0)
        assert service.get_field_value("difficulty.judgment_trap_density") == 0.0

    def test_set_judgment_trap_density_boundary(self, service: DraftService) -> None:
        service.set_field("difficulty.judgment_trap_density", 1.0)
        assert service.get_field_value("difficulty.judgment_trap_density") == 1.0

    def test_difficulty_diff_tracks_changes(self, service: DraftService) -> None:
        service.set_field("difficulty.error_density", 0.7)
        diffs = service.diff()
        paths = [d.path for d in diffs]
        assert "difficulty.error_density" in paths

    def test_reset_difficulty_field(self, service: DraftService) -> None:
        service.set_field("difficulty.error_density", 0.2)
        service.reset_field("difficulty.error_density")
        assert not service.is_modified("difficulty.error_density")
        assert service.get_field_value("difficulty.error_density") == 1.0


# ---------------------------------------------------------------------------
# Invalid difficulty profile updates
# ---------------------------------------------------------------------------


class TestInvalidDifficultyUpdates:
    """Service rejects invalid difficulty field values."""

    def test_error_density_above_max(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="above maximum"):
            service.set_field("difficulty.error_density", 1.5)

    def test_error_density_below_min(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="below minimum"):
            service.set_field("difficulty.error_density", -0.1)

    def test_canary_visibility_invalid_choice(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="Invalid choice"):
            service.set_field("difficulty.canary_visibility", "invisible")

    def test_judgment_trap_density_above_max(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="above maximum"):
            service.set_field("difficulty.judgment_trap_density", 2.0)

    def test_judgment_trap_density_negative(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="below minimum"):
            service.set_field("difficulty.judgment_trap_density", -0.5)


# ---------------------------------------------------------------------------
# Valid output profile updates
# ---------------------------------------------------------------------------


class TestValidOutputUpdates:
    """Service accepts valid output field changes."""

    def test_set_enabled_test_cases(self, service: DraftService) -> None:
        service.set_field("output.enabled_test_cases", ["TC-01", "TC-06"])
        assert service.get_field_value("output.enabled_test_cases") == ["TC-01", "TC-06"]

    def test_set_empty_test_cases_means_all(self, service: DraftService) -> None:
        service.set_field("output.enabled_test_cases", [])
        assert service.get_field_value("output.enabled_test_cases") == []

    def test_set_all_test_cases(self, service: DraftService) -> None:
        all_tcs = [f"TC-{i:02d}" for i in range(1, 19)]
        service.set_field("output.enabled_test_cases", all_tcs)
        assert service.get_field_value("output.enabled_test_cases") == all_tcs

    def test_set_enabled_packs(self, service: DraftService) -> None:
        service.set_field("output.enabled_packs", ["audit", "tax"])
        assert service.get_field_value("output.enabled_packs") == ["audit", "tax"]

    def test_set_empty_packs(self, service: DraftService) -> None:
        service.set_field("output.enabled_packs", [])
        assert service.get_field_value("output.enabled_packs") == []

    def test_output_diff_tracks_changes(self, service: DraftService) -> None:
        service.set_field("output.enabled_test_cases", ["TC-03"])
        diffs = service.diff()
        paths = [d.path for d in diffs]
        assert "output.enabled_test_cases" in paths

    def test_reset_output_field(self, service: DraftService) -> None:
        service.set_field("output.enabled_test_cases", ["TC-01"])
        service.reset_field("output.enabled_test_cases")
        assert not service.is_modified("output.enabled_test_cases")


# ---------------------------------------------------------------------------
# Invalid output profile updates
# ---------------------------------------------------------------------------


class TestInvalidOutputUpdates:
    """Service rejects invalid output field values."""

    def test_invalid_test_case_choice(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="Invalid choice"):
            service.set_field("output.enabled_test_cases", ["TC-99"])

    def test_test_cases_not_a_list(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="expects a list"):
            service.set_field("output.enabled_test_cases", "TC-01")

    def test_packs_not_a_list(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="expects a list"):
            service.set_field("output.enabled_packs", "audit")


# ---------------------------------------------------------------------------
# Unsupported field rejection
# ---------------------------------------------------------------------------


class TestUnsupportedFieldRejection:
    """Unsupported paths are rejected by the service."""

    def test_canary_assignments_rejected(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="not user-editable"):
            service.set_field("canary_assignments", {"x": "y"})

    def test_error_injections_rejected(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="not user-editable"):
            service.set_field("error_injections", [])

    def test_entity_code_rejected(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="not user-editable"):
            service.set_field("company.subsidiaries.sub_a.entity_code", "XX")


# ---------------------------------------------------------------------------
# Screen-level helper tests
# ---------------------------------------------------------------------------


class TestScreenHelpers:
    """Test the is_unsupported_difficulty_output_path helper."""

    def test_difficulty_path_supported(self) -> None:
        from generator.tui.difficulty_output_screen import (
            is_unsupported_difficulty_output_path,
        )

        assert not is_unsupported_difficulty_output_path("difficulty.error_density")

    def test_output_path_supported(self) -> None:
        from generator.tui.difficulty_output_screen import (
            is_unsupported_difficulty_output_path,
        )

        assert not is_unsupported_difficulty_output_path("output.enabled_test_cases")

    def test_company_path_unsupported_in_this_screen(self) -> None:
        from generator.tui.difficulty_output_screen import (
            is_unsupported_difficulty_output_path,
        )

        assert is_unsupported_difficulty_output_path("company.name")

    def test_computed_path_unsupported(self) -> None:
        from generator.tui.difficulty_output_screen import (
            is_unsupported_difficulty_output_path,
        )

        assert is_unsupported_difficulty_output_path("canary_assignments")

    def test_unknown_path_unsupported(self) -> None:
        from generator.tui.difficulty_output_screen import (
            is_unsupported_difficulty_output_path,
        )

        assert is_unsupported_difficulty_output_path("foo.bar.baz")


# ---------------------------------------------------------------------------
# Combined profile update workflow
# ---------------------------------------------------------------------------


class TestCombinedProfileWorkflow:
    """End-to-end workflow: set difficulty + output, validate, diff, save."""

    def test_set_multiple_then_validate(self, service: DraftService) -> None:
        service.set_field("difficulty.error_density", 0.5)
        service.set_field("difficulty.canary_visibility", "subtle")
        service.set_field("output.enabled_test_cases", ["TC-01", "TC-02"])
        result = service.validate()
        assert result.valid

    def test_set_multiple_then_diff(self, service: DraftService) -> None:
        service.set_field("difficulty.error_density", 0.5)
        service.set_field("output.enabled_packs", ["audit"])
        diffs = service.diff()
        paths = {d.path for d in diffs}
        assert "difficulty.error_density" in paths
        assert "output.enabled_packs" in paths

    def test_set_save_reload_roundtrip(
        self, base_yaml: Path, service: DraftService, tmp_path: Path
    ) -> None:
        service.set_field("difficulty.error_density", 0.3)
        service.set_field("difficulty.canary_visibility", "hidden")
        service.set_field("output.enabled_test_cases", ["TC-05", "TC-10"])

        out = tmp_path / "override.yaml"
        service.save(out)

        svc2 = DraftService(base_yaml)
        svc2.load_draft(out)
        assert svc2.get_field_value("difficulty.error_density") == 0.3
        assert svc2.get_field_value("difficulty.canary_visibility") == "hidden"
        assert svc2.get_field_value("output.enabled_test_cases") == ["TC-05", "TC-10"]

    def test_reset_all_difficulty_output(self, service: DraftService) -> None:
        service.set_field("difficulty.error_density", 0.1)
        service.set_field("output.enabled_packs", ["tax"])

        for meta_path in ["difficulty.error_density", "output.enabled_packs"]:
            service.reset_field(meta_path)

        assert not service.is_modified("difficulty.error_density")
        assert not service.is_modified("output.enabled_packs")
        assert service.get_field_value("difficulty.error_density") == 1.0
