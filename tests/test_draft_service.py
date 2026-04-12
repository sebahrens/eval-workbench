"""Tests for the TUI draft service.

Covers: load, edit, validation failure, diff, and save.
Bead: synth-data-2u6.6.3
"""

from __future__ import annotations

import textwrap
from pathlib import Path

import pytest
import yaml

from generator.tui.draft_service import DraftService

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
# Load tests
# ---------------------------------------------------------------------------


class TestLoad:
    """Service loads base config correctly."""

    def test_load_base_only(self, service: DraftService) -> None:
        assert service.get_field_value("company.name") == "TestCo"
        assert service.get_field_value("seed") == 42

    def test_load_with_overlay(self, base_yaml: Path, tmp_path: Path) -> None:
        overlay = tmp_path / "overlay.yaml"
        overlay.write_text(yaml.dump({"company": {"name": "OverlayCo"}}))

        svc = DraftService(base_yaml, overlays=[overlay])
        assert svc.get_field_value("company.name") == "OverlayCo"

    def test_load_draft_file(self, service: DraftService, tmp_path: Path) -> None:
        draft_file = tmp_path / "draft.yaml"
        draft_file.write_text(yaml.dump({"company": {"name": "DraftCo"}}))

        service.load_draft(draft_file)
        assert service.get_field_value("company.name") == "DraftCo"
        assert service.is_modified("company.name")

    def test_load_empty_draft_file(self, service: DraftService, tmp_path: Path) -> None:
        draft_file = tmp_path / "empty.yaml"
        draft_file.write_text("")

        service.load_draft(draft_file)
        assert service.draft == {}

    def test_load_missing_draft_raises(self, service: DraftService, tmp_path: Path) -> None:
        from generator.config import ConfigError

        with pytest.raises(ConfigError, match="not found"):
            service.load_draft(tmp_path / "nonexistent.yaml")


# ---------------------------------------------------------------------------
# Edit tests
# ---------------------------------------------------------------------------


class TestEdit:
    """Service applies typed field updates to the draft."""

    def test_set_scalar_field(self, service: DraftService) -> None:
        service.set_field("company.name", "NewCo")
        assert service.get_field_value("company.name") == "NewCo"
        assert service.is_modified("company.name")

    def test_set_integer_field_coerces(self, service: DraftService) -> None:
        service.set_field("company.consolidated_revenue", "200000000")
        assert service.get_field_value("company.consolidated_revenue") == 200000000

    def test_set_float_field_coerces(self, service: DraftService) -> None:
        service.set_field("difficulty.error_density", "0.5")
        assert service.get_field_value("difficulty.error_density") == 0.5

    def test_set_choice_field_validates(self, service: DraftService) -> None:
        service.set_field("difficulty.canary_visibility", "subtle")
        assert service.get_field_value("difficulty.canary_visibility") == "subtle"

    def test_set_invalid_choice_raises(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="Invalid choice"):
            service.set_field("difficulty.canary_visibility", "invisible")

    def test_set_range_below_min_raises(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="below minimum"):
            service.set_field("company.consolidated_revenue", -1)

    def test_set_range_above_max_raises(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="above maximum"):
            service.set_field("difficulty.error_density", 1.5)

    def test_set_unsupported_path_raises(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="not user-editable"):
            service.set_field("canary_assignments", {"x": "y"})

    def test_set_list_field(self, service: DraftService) -> None:
        service.set_field("company.employees.remote_states", ["TX", "FL"])
        assert service.get_field_value("company.employees.remote_states") == ["TX", "FL"]

    def test_reset_field(self, service: DraftService) -> None:
        service.set_field("company.name", "TempCo")
        assert service.is_modified("company.name")

        service.reset_field("company.name")
        assert not service.is_modified("company.name")
        # Reverts to base
        assert service.get_field_value("company.name") == "TestCo"

    def test_reset_all(self, service: DraftService) -> None:
        service.set_field("company.name", "TempCo")
        service.set_field("seed", 99)
        service.reset_all()
        assert service.draft == {}
        assert service.get_field_value("company.name") == "TestCo"

    def test_set_unknown_path_allowed(self, service: DraftService) -> None:
        """Paths without metadata are set without coercion (freeform)."""
        service.set_field("company.industry", "Technology")
        assert service.get_field_value("company.industry") == "Technology"


# ---------------------------------------------------------------------------
# Validation tests
# ---------------------------------------------------------------------------


class TestValidation:
    """Service validates merged config through the standard loader."""

    def test_valid_config(self, service: DraftService) -> None:
        result = service.validate()
        assert result.valid
        assert result.config is not None
        assert result.config.company.name == "TestCo"

    def test_valid_after_edit(self, service: DraftService) -> None:
        service.set_field("company.name", "ValidCo")
        result = service.validate()
        assert result.valid
        assert result.config is not None
        assert result.config.company.name == "ValidCo"

    def test_invalid_seasonal_weights(self, service: DraftService) -> None:
        service.set_field("company.seasonal_weights.Q1", 0.99)
        result = service.validate()
        assert not result.valid
        assert any("Seasonal weights" in e for e in result.errors)

    def test_validation_errors_returned(self, service: DraftService) -> None:
        # Force an invalid config by removing a required field via None
        service._draft = {"company": {"name": None}}
        # deep_merge with None deletes the key
        result = service.validate()
        # Config loader should complain about missing required keys or type error
        assert not result.valid
        assert len(result.errors) > 0


# ---------------------------------------------------------------------------
# Diff tests
# ---------------------------------------------------------------------------


class TestDiff:
    """Service computes field-level diffs between base and draft."""

    def test_no_diff_when_no_edits(self, service: DraftService) -> None:
        diffs = service.diff()
        assert diffs == []

    def test_single_field_diff(self, service: DraftService) -> None:
        service.set_field("company.name", "DiffCo")
        diffs = service.diff()
        assert len(diffs) == 1
        assert diffs[0].path == "company.name"
        assert diffs[0].base_value == "TestCo"
        assert diffs[0].draft_value == "DiffCo"

    def test_multiple_field_diffs(self, service: DraftService) -> None:
        service.set_field("company.name", "Multi")
        service.set_field("seed", 99)
        diffs = service.diff()
        paths = [d.path for d in diffs]
        assert "company.name" in paths
        assert "seed" in paths

    def test_same_value_no_diff(self, service: DraftService) -> None:
        """Setting a field to its current base value produces no diff."""
        service.set_field("company.name", "TestCo")
        diffs = service.diff()
        assert diffs == []

    def test_diff_after_reset(self, service: DraftService) -> None:
        service.set_field("company.name", "TempCo")
        service.reset_field("company.name")
        diffs = service.diff()
        assert diffs == []


# ---------------------------------------------------------------------------
# Save tests
# ---------------------------------------------------------------------------


class TestSave:
    """Service performs atomic YAML save of draft overrides."""

    def test_save_creates_file(self, service: DraftService, tmp_path: Path) -> None:
        service.set_field("company.name", "SavedCo")
        out = tmp_path / "saved_override.yaml"
        service.save(out)
        assert out.exists()

    def test_save_contains_only_draft(self, service: DraftService, tmp_path: Path) -> None:
        service.set_field("company.name", "SavedCo")
        out = tmp_path / "saved_override.yaml"
        service.save(out)

        saved = yaml.safe_load(out.read_text())
        # Should only contain the override, not the full config
        assert saved == {"company": {"name": "SavedCo"}}

    def test_save_does_not_modify_base(
        self, base_yaml: Path, service: DraftService, tmp_path: Path
    ) -> None:
        original = base_yaml.read_text()
        service.set_field("company.name", "Modified")
        service.save(tmp_path / "override.yaml")
        assert base_yaml.read_text() == original

    def test_save_empty_draft(self, service: DraftService, tmp_path: Path) -> None:
        out = tmp_path / "empty_override.yaml"
        service.save(out)
        saved = yaml.safe_load(out.read_text())
        assert saved == {} or saved is None

    def test_save_creates_parent_dirs(self, service: DraftService, tmp_path: Path) -> None:
        service.set_field("seed", 99)
        out = tmp_path / "nested" / "dir" / "override.yaml"
        service.save(out)
        assert out.exists()

    def test_save_refuses_overwrite_without_force(
        self, service: DraftService, tmp_path: Path
    ) -> None:
        """Saving to an existing file without force=True raises FileExistsError."""
        out = tmp_path / "override.yaml"
        out.write_text("old content")

        service.set_field("seed", 99)
        with pytest.raises(FileExistsError, match="already exists"):
            service.save(out)

        # Original content is untouched
        assert out.read_text() == "old content"

    def test_save_overwrites_with_force(
        self, service: DraftService, tmp_path: Path
    ) -> None:
        """Overwriting an existing file uses atomic rename when force=True."""
        out = tmp_path / "override.yaml"
        out.write_text("old content")

        service.set_field("seed", 99)
        service.save(out, force=True)

        saved = yaml.safe_load(out.read_text())
        assert saved["seed"] == 99

    def test_save_validates_before_writing(
        self, service: DraftService, tmp_path: Path
    ) -> None:
        """Save rejects invalid merged config before touching disk."""
        from generator.config import ConfigError

        # Break the config with invalid seasonal weights
        service.set_field("company.seasonal_weights.Q1", 0.99)
        out = tmp_path / "invalid_override.yaml"
        with pytest.raises(ConfigError, match="Cannot save"):
            service.save(out)

        assert not out.exists()

    def test_save_skips_validation_when_disabled(
        self, service: DraftService, tmp_path: Path
    ) -> None:
        """validate=False skips pre-save validation."""
        service.set_field("company.seasonal_weights.Q1", 0.99)
        out = tmp_path / "no_validate.yaml"
        service.save(out, validate=False)
        assert out.exists()

    def test_saved_draft_reloads_into_service(
        self, base_yaml: Path, service: DraftService, tmp_path: Path
    ) -> None:
        """A saved draft can be reloaded into a new service instance."""
        service.set_field("company.name", "RoundTrip")
        service.set_field("seed", 77)
        out = tmp_path / "override.yaml"
        service.save(out)

        # New service, load the saved draft
        svc2 = DraftService(base_yaml)
        svc2.load_draft(out)
        assert svc2.get_field_value("company.name") == "RoundTrip"
        assert svc2.get_field_value("seed") == 77

    def test_saved_draft_reloads_through_layered_config(
        self, base_yaml: Path, service: DraftService, tmp_path: Path
    ) -> None:
        """Acceptance criteria: saved YAML reloads through load_layered_config."""
        from generator.config import load_layered_config

        service.set_field("company.name", "LayeredCo")
        service.set_field("seed", 55)
        out = tmp_path / "override.yaml"
        service.save(out)

        # Load through the layered config loader (base + overlay)
        config = load_layered_config(base_yaml, layers=[out])
        assert config.company.name == "LayeredCo"
        assert config.seed == 55
