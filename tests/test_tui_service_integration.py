"""Integration tests for TUI service, state transitions, and smoke paths.

Covers:
- Service state transitions (fresh → edit → validate → save → reload)
- Schema metadata coverage across all config field groups
- Validation/save blocking (invalid config blocks save unless forced)
- Draft import/export round-trips (load_draft from YAML files)
- TUI __main__ argument parsing
- Headless smoke: subprocess invocations of the TUI entry points

Bead: synth-data-2u6.6.12
"""

from __future__ import annotations

import subprocess
import sys
import textwrap
from pathlib import Path

import pytest
import yaml

from generator.config import ConfigError, load_layered_config
from generator.schema_metadata import (
    FieldGroup,
    InputType,
    get_all_fields,
    get_field_by_path,
    get_fields_by_group,
    get_subsidiary_fields,
)
from generator.tui.draft_service import DraftService

# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


def _base_config_yaml() -> str:
    """Minimal valid base config for integration tests."""
    return textwrap.dedent("""\
        seed: 42
        output_dir: out
        company:
          name: TestCo
          type: C-Corp
          industry: Manufacturing
          headquarters: Portland, OR
          fiscal_year_end: "12-31"
          years: [2023, 2024, 2025]
          current_year: 2025
          consolidated_revenue: 100000000
          subsidiaries:
            sub_a:
              legal_name: Sub A LLC
              location: Portland, OR
              state: OR
              entity_code: SA
              revenue: 60000000
              type: Manufacturing
              gross_margin: 0.35
              employee_count: 200
            sub_b:
              legal_name: Sub B Inc.
              location: Austin, TX
              state: TX
              entity_code: SB
              revenue: 40000000
              type: R&D
              gross_margin: 0.50
              employee_count: 100
              rd_spend_pct: 0.10
          growth_rates:
            fy2023_to_fy2024: 0.06
            fy2024_to_fy2025: 0.09
          intercompany:
            raw_materials_markup: 0.08
            management_fee_pct: 0.015
            intercompany_loan_principal: 5000000
            intercompany_loan_rate: 0.05
          employees:
            total_count: 300
            annual_turnover_rate: 0.08
            remote_states: [CA, WA]
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
# State transition tests
# ---------------------------------------------------------------------------


class TestServiceStateTransitions:
    """Test the full lifecycle: fresh → edit → validate → save → reload."""

    def test_fresh_service_has_empty_draft(self, service: DraftService) -> None:
        assert service.draft == {}
        assert service.diff() == []

    def test_fresh_service_validates_clean(self, service: DraftService) -> None:
        result = service.validate()
        assert result.valid
        assert result.config is not None

    def test_edit_then_validate_then_save_roundtrip(
        self, base_yaml: Path, service: DraftService, tmp_path: Path
    ) -> None:
        """Full state machine: edit multiple fields → validate → save → reload."""
        # 1. Edit
        service.set_field("company.name", "Acme Corp")
        service.set_field("difficulty.error_density", 0.5)
        service.set_field("output.enabled_test_cases", ["TC-01", "TC-06"])

        assert service.is_modified("company.name")
        assert service.is_modified("difficulty.error_density")

        # 2. Validate
        result = service.validate()
        assert result.valid
        assert result.config.company.name == "Acme Corp"

        # 3. Diff shows exactly the changed fields
        diffs = service.diff()
        paths = {d.path for d in diffs}
        assert paths == {
            "company.name",
            "difficulty.error_density",
            "output.enabled_test_cases",
        }

        # 4. Save
        out = tmp_path / "override.yaml"
        service.save(out)
        assert out.exists()

        # 5. Reload into a fresh service
        svc2 = DraftService(base_yaml)
        svc2.load_draft(out)
        assert svc2.get_field_value("company.name") == "Acme Corp"
        assert svc2.get_field_value("difficulty.error_density") == 0.5
        assert svc2.get_field_value("output.enabled_test_cases") == ["TC-01", "TC-06"]
        # Base fields unchanged
        assert svc2.get_field_value("seed") == 42

    def test_edit_reset_leaves_no_diff(self, service: DraftService) -> None:
        """Editing then resetting all returns to fresh state."""
        service.set_field("company.name", "TempCo")
        service.set_field("seed", 99)
        service.reset_all()
        assert service.draft == {}
        assert service.diff() == []

    def test_edit_validate_fail_fix_validate_pass(
        self, service: DraftService
    ) -> None:
        """Invalid edit → failed validation → fix → valid."""
        # Break: seasonal weights out of balance
        service.set_field("company.seasonal_weights.Q1", 0.99)
        bad = service.validate()
        assert not bad.valid

        # Fix
        service.reset_field("company.seasonal_weights.Q1")
        good = service.validate()
        assert good.valid

    def test_multiple_edits_to_same_field_keeps_last(
        self, service: DraftService
    ) -> None:
        """Successive edits to the same field keep only the last value."""
        service.set_field("company.name", "First")
        service.set_field("company.name", "Second")
        service.set_field("company.name", "Third")
        assert service.get_field_value("company.name") == "Third"
        diffs = service.diff()
        assert len(diffs) == 1
        assert diffs[0].draft_value == "Third"


# ---------------------------------------------------------------------------
# Metadata coverage tests
# ---------------------------------------------------------------------------


class TestMetadataCoverage:
    """Verify schema metadata covers all config field groups and key paths."""

    def test_every_group_has_at_least_one_field(self) -> None:
        for group in FieldGroup:
            if group == FieldGroup.SUBSIDIARY:
                assert len(get_subsidiary_fields()) > 0
            else:
                fields = get_fields_by_group(group)
                assert len(fields) > 0, f"No fields for {group.name}"

    def test_all_fields_have_valid_input_types(self) -> None:
        for f in get_all_fields():
            assert isinstance(f.input_type, InputType), (
                f"Field {f.path} has invalid input_type: {f.input_type}"
            )

    def test_choice_fields_have_nonempty_choices(self) -> None:
        for f in get_all_fields():
            if f.input_type in (InputType.CHOICE, InputType.MULTI_CHOICE):
                assert len(f.choices) > 0, (
                    f"Field {f.path} is {f.input_type.value} but has no choices"
                )

    def test_range_fields_have_consistent_bounds(self) -> None:
        for f in get_all_fields():
            if f.range_min is not None and f.range_max is not None:
                assert f.range_min <= f.range_max, (
                    f"Field {f.path}: range_min ({f.range_min}) > range_max ({f.range_max})"
                )

    def test_company_group_covers_expected_paths(self) -> None:
        fields = get_fields_by_group(FieldGroup.COMPANY)
        paths = {f.path for f in fields}
        expected = {
            "company.name", "company.type", "company.industry",
            "company.headquarters", "company.fiscal_year_end",
            "company.years", "company.current_year",
            "company.consolidated_revenue",
        }
        assert expected.issubset(paths)

    def test_difficulty_group_covers_expected_paths(self) -> None:
        fields = get_fields_by_group(FieldGroup.DIFFICULTY)
        paths = {f.path for f in fields}
        expected = {
            "difficulty.error_density",
            "difficulty.canary_visibility",
            "difficulty.judgment_trap_density",
        }
        assert paths == expected

    def test_generator_fields_present(self) -> None:
        assert get_field_by_path("seed") is not None
        assert get_field_by_path("output_dir") is not None

    def test_no_duplicate_paths_in_registry(self) -> None:
        paths = [f.path for f in get_all_fields()]
        assert len(paths) == len(set(paths))


# ---------------------------------------------------------------------------
# Validation/save blocking tests
# ---------------------------------------------------------------------------


class TestValidationSaveBlocking:
    """Verify that save is gated on validation by default."""

    def test_save_blocked_when_config_invalid(
        self, service: DraftService, tmp_path: Path
    ) -> None:
        """An invalid merged config blocks save with a clear error."""
        service.set_field("company.seasonal_weights.Q1", 0.99)
        out = tmp_path / "invalid.yaml"
        with pytest.raises(ConfigError, match="Cannot save"):
            service.save(out)
        assert not out.exists()

    def test_save_proceeds_when_validation_disabled(
        self, service: DraftService, tmp_path: Path
    ) -> None:
        """validate=False lets invalid configs through."""
        service.set_field("company.seasonal_weights.Q1", 0.99)
        out = tmp_path / "force_invalid.yaml"
        service.save(out, validate=False)
        assert out.exists()

    def test_save_blocked_by_existing_file_without_force(
        self, service: DraftService, tmp_path: Path
    ) -> None:
        out = tmp_path / "existing.yaml"
        out.write_text("old")
        service.set_field("seed", 99)
        with pytest.raises(FileExistsError, match="already exists"):
            service.save(out)
        assert out.read_text() == "old"

    def test_save_overwrites_existing_with_force(
        self, service: DraftService, tmp_path: Path
    ) -> None:
        out = tmp_path / "existing.yaml"
        out.write_text("old")
        service.set_field("seed", 99)
        service.save(out, force=True)
        saved = yaml.safe_load(out.read_text())
        assert saved["seed"] == 99

    def test_validation_errors_are_strings(self, service: DraftService) -> None:
        """Validation failures return human-readable string messages."""
        service.set_field("company.seasonal_weights.Q1", 0.99)
        result = service.validate()
        assert not result.valid
        for e in result.errors:
            assert isinstance(e, str)
            assert len(e) > 0


# ---------------------------------------------------------------------------
# Draft import tests
# ---------------------------------------------------------------------------


class TestDraftImport:
    """Test loading draft overrides from YAML files."""

    def test_load_draft_applies_overrides(
        self, service: DraftService, tmp_path: Path
    ) -> None:
        draft_file = tmp_path / "draft.yaml"
        draft_file.write_text(yaml.dump({
            "company": {"name": "ImportedCo"},
            "seed": 77,
        }))
        service.load_draft(draft_file)
        assert service.get_field_value("company.name") == "ImportedCo"
        assert service.get_field_value("seed") == 77
        assert service.is_modified("company.name")

    def test_load_draft_replaces_previous_draft(
        self, service: DraftService, tmp_path: Path
    ) -> None:
        """Loading a new draft replaces (not merges with) the previous one."""
        service.set_field("company.name", "EditedCo")
        draft_file = tmp_path / "draft.yaml"
        draft_file.write_text(yaml.dump({"seed": 77}))
        service.load_draft(draft_file)
        # Previous edit was replaced
        assert not service.is_modified("company.name")
        assert service.is_modified("seed")

    def test_load_draft_from_nonexistent_file_raises(
        self, service: DraftService, tmp_path: Path
    ) -> None:
        with pytest.raises(ConfigError, match="not found"):
            service.load_draft(tmp_path / "nope.yaml")

    def test_load_draft_non_dict_raises(
        self, service: DraftService, tmp_path: Path
    ) -> None:
        bad_file = tmp_path / "bad.yaml"
        bad_file.write_text("- item1\n- item2\n")
        with pytest.raises(ConfigError, match="mapping"):
            service.load_draft(bad_file)

    def test_load_empty_draft_clears_state(
        self, service: DraftService, tmp_path: Path
    ) -> None:
        service.set_field("company.name", "EditedCo")
        empty_file = tmp_path / "empty.yaml"
        empty_file.write_text("")
        service.load_draft(empty_file)
        assert service.draft == {}
        assert not service.is_modified("company.name")

    def test_saved_draft_reloads_via_layered_config(
        self, base_yaml: Path, service: DraftService, tmp_path: Path
    ) -> None:
        """Saved draft works as an overlay with load_layered_config."""
        service.set_field("company.name", "LayeredCo")
        service.set_field("seed", 55)
        out = tmp_path / "override.yaml"
        service.save(out)

        config = load_layered_config(base_yaml, layers=[out])
        assert config.company.name == "LayeredCo"
        assert config.seed == 55
        # Non-overridden fields preserved
        assert config.company.headquarters == "Portland, OR"


# ---------------------------------------------------------------------------
# TUI __main__ argument parsing tests
# ---------------------------------------------------------------------------


class TestTuiMainArgParsing:
    """Test the CLI entry point argument parsing without launching the TUI."""

    def test_main_parses_default_config(self) -> None:
        import generator.tui.__main__ as main_mod

        captured = {}

        def mock_configure(**kwargs):
            captured.update(kwargs)

        original = main_mod.configure
        main_mod.configure = mock_configure
        try:
            main_mod.main(["--config", "custom.yaml"])
            assert captured["config_path"] == "custom.yaml"
            assert captured["overlay"] is None
        finally:
            main_mod.configure = original

    def test_main_parses_overlays(self) -> None:
        import generator.tui.__main__ as main_mod

        captured = {}

        def mock_configure(**kwargs):
            captured.update(kwargs)

        original = main_mod.configure
        main_mod.configure = mock_configure
        try:
            main_mod.main(["--config", "base.yaml", "--overlay", "a.yaml", "b.yaml"])
            assert captured["config_path"] == "base.yaml"
            assert captured["overlay"] == ["a.yaml", "b.yaml"]
        finally:
            main_mod.configure = original

    def test_main_default_config_path(self) -> None:
        import generator.tui.__main__ as main_mod

        captured = {}

        def mock_configure(**kwargs):
            captured.update(kwargs)

        original = main_mod.configure
        main_mod.configure = mock_configure
        try:
            main_mod.main([])
            assert captured["config_path"] == "config.yaml"
        finally:
            main_mod.configure = original


# ---------------------------------------------------------------------------
# Headless smoke tests (subprocess)
# ---------------------------------------------------------------------------


class TestHeadlessSmoke:
    """Subprocess smoke tests that verify entry points don't crash."""

    def test_tui_module_importable(self) -> None:
        result = subprocess.run(
            [sys.executable, "-c", "import generator.tui; import generator.tui.draft_service"],
            capture_output=True,
            text=True,
        )
        assert result.returncode == 0, result.stderr

    def test_tui_help_flag(self) -> None:
        result = subprocess.run(
            [sys.executable, "-m", "generator.tui", "--help"],
            capture_output=True,
            text=True,
        )
        assert result.returncode == 0
        assert "scenario configuration" in result.stdout.lower()

    def test_configure_subcommand_help(self) -> None:
        result = subprocess.run(
            [sys.executable, "generate_test_suite.py", "configure", "--help"],
            capture_output=True,
            text=True,
        )
        assert result.returncode == 0
        assert "configuration" in result.stdout.lower()

    def test_draft_service_headless(self, base_yaml: Path, tmp_path: Path) -> None:
        """Run a full draft-service workflow in a subprocess to verify no hidden imports fail."""
        script = textwrap.dedent(f"""\
            import yaml
            from pathlib import Path
            from generator.tui.draft_service import DraftService

            svc = DraftService("{base_yaml}")
            svc.set_field("company.name", "SubprocessCo")
            result = svc.validate()
            assert result.valid, f"Validation failed: {{result.errors}}"
            diffs = svc.diff()
            assert len(diffs) == 1
            out = Path("{tmp_path}/headless.yaml")
            svc.save(out)
            assert out.exists()
            saved = yaml.safe_load(out.read_text())
            assert saved == {{"company": {{"name": "SubprocessCo"}}}}
        """)
        result = subprocess.run(
            [sys.executable, "-c", script],
            capture_output=True,
            text=True,
        )
        assert result.returncode == 0, result.stderr
