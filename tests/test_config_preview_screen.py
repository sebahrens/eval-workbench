"""Tests for the merged config preview and diff screen.

Covers: empty diff, changed-field diff, YAML preview rendering,
summary text, and verifying the preview does not mutate the draft.

Bead: synth-data-2u6.6.9
"""

from __future__ import annotations

import textwrap
from pathlib import Path

import pytest
import yaml

from generator.tui.draft_service import DraftService, FieldDiff

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
# Helper function tests
# ---------------------------------------------------------------------------


class TestFormatValue:
    """Test the _format_value display helper."""

    def test_none(self) -> None:
        from generator.tui.config_preview_screen import _format_value

        assert "(unset)" in _format_value(None)

    def test_empty_list(self) -> None:
        from generator.tui.config_preview_screen import _format_value

        assert "(empty list)" in _format_value([])

    def test_list(self) -> None:
        from generator.tui.config_preview_screen import _format_value

        assert _format_value(["a", "b"]) == "a, b"

    def test_scalar(self) -> None:
        from generator.tui.config_preview_screen import _format_value

        assert _format_value(0.5) == "0.5"


class TestBuildSummary:
    """Test the _build_summary helper."""

    def test_empty_diffs(self) -> None:
        from generator.tui.config_preview_screen import _build_summary

        assert _build_summary([]) == "No changes."

    def test_one_diff(self) -> None:
        from generator.tui.config_preview_screen import _build_summary

        diffs = [FieldDiff(path="x", base_value=1, draft_value=2)]
        assert _build_summary(diffs) == "1 field modified."

    def test_multiple_diffs(self) -> None:
        from generator.tui.config_preview_screen import _build_summary

        diffs = [
            FieldDiff(path="a", base_value=1, draft_value=2),
            FieldDiff(path="b", base_value=3, draft_value=4),
            FieldDiff(path="c", base_value=5, draft_value=6),
        ]
        assert _build_summary(diffs) == "3 fields modified."


class TestBuildDiffText:
    """Test the _build_diff_text helper."""

    def test_no_changes(self) -> None:
        from generator.tui.config_preview_screen import _build_diff_text

        result = _build_diff_text([])
        assert "No overrides" in result

    def test_with_changes(self) -> None:
        from generator.tui.config_preview_screen import _build_diff_text

        diffs = [FieldDiff(path="difficulty.error_density", base_value=1.0, draft_value=0.5)]
        result = _build_diff_text(diffs)
        assert "difficulty.error_density" in result
        assert "1.0" in result
        assert "0.5" in result
        assert "Changed Fields" in result


class TestBuildYamlPreview:
    """Test the _build_yaml_preview helper."""

    def test_empty_draft(self) -> None:
        from generator.tui.config_preview_screen import _build_yaml_preview

        result = _build_yaml_preview({})
        assert "empty" in result

    def test_non_empty_draft(self) -> None:
        from generator.tui.config_preview_screen import _build_yaml_preview

        draft = {"difficulty": {"error_density": 0.5}}
        result = _build_yaml_preview(draft)
        parsed = yaml.safe_load(result)
        assert parsed == {"difficulty": {"error_density": 0.5}}


# ---------------------------------------------------------------------------
# Empty diff scenario
# ---------------------------------------------------------------------------


class TestEmptyDiff:
    """When no fields are modified, preview shows no-change state."""

    def test_empty_diff_returns_no_diffs(self, service: DraftService) -> None:
        diffs = service.diff()
        assert diffs == []

    def test_empty_diff_summary(self, service: DraftService) -> None:
        from generator.tui.config_preview_screen import _build_summary

        assert _build_summary(service.diff()) == "No changes."

    def test_empty_draft_yaml_preview(self, service: DraftService) -> None:
        from generator.tui.config_preview_screen import _build_yaml_preview

        result = _build_yaml_preview(service.draft)
        assert "empty" in result


# ---------------------------------------------------------------------------
# Changed-field diff scenario
# ---------------------------------------------------------------------------


class TestChangedFieldDiff:
    """When fields are modified, preview correctly reflects changes."""

    def test_single_field_diff(self, service: DraftService) -> None:
        service.set_field("difficulty.error_density", 0.5)
        diffs = service.diff()
        assert len(diffs) == 1
        assert diffs[0].path == "difficulty.error_density"
        assert diffs[0].base_value == 1.0
        assert diffs[0].draft_value == 0.5

    def test_multiple_field_diffs(self, service: DraftService) -> None:
        service.set_field("difficulty.error_density", 0.3)
        service.set_field("difficulty.canary_visibility", "hidden")
        service.set_field("output.enabled_test_cases", ["TC-01", "TC-02"])
        diffs = service.diff()
        paths = {d.path for d in diffs}
        assert "difficulty.error_density" in paths
        assert "difficulty.canary_visibility" in paths
        assert "output.enabled_test_cases" in paths

    def test_diff_text_includes_all_changed_paths(self, service: DraftService) -> None:
        from generator.tui.config_preview_screen import _build_diff_text

        service.set_field("difficulty.error_density", 0.2)
        service.set_field("difficulty.canary_visibility", "subtle")
        text = _build_diff_text(service.diff())
        assert "difficulty.error_density" in text
        assert "difficulty.canary_visibility" in text

    def test_yaml_preview_matches_draft(self, service: DraftService) -> None:
        from generator.tui.config_preview_screen import _build_yaml_preview

        service.set_field("difficulty.error_density", 0.7)
        draft = service.draft
        rendered = _build_yaml_preview(draft)
        parsed = yaml.safe_load(rendered)
        assert parsed == draft

    def test_summary_with_changes(self, service: DraftService) -> None:
        from generator.tui.config_preview_screen import _build_summary

        service.set_field("difficulty.error_density", 0.5)
        service.set_field("difficulty.canary_visibility", "hidden")
        assert _build_summary(service.diff()) == "2 fields modified."


# ---------------------------------------------------------------------------
# Preview does not mutate draft
# ---------------------------------------------------------------------------


class TestPreviewDoesNotMutate:
    """Verify that building the preview never modifies the draft state."""

    def test_diff_is_read_only(self, service: DraftService) -> None:
        service.set_field("difficulty.error_density", 0.5)
        draft_before = service.draft

        # Call diff and build text — should not mutate
        diffs = service.diff()
        from generator.tui.config_preview_screen import _build_diff_text

        _build_diff_text(diffs)

        draft_after = service.draft
        assert draft_before == draft_after

    def test_yaml_preview_is_read_only(self, service: DraftService) -> None:
        service.set_field("difficulty.error_density", 0.5)
        draft_before = service.draft

        from generator.tui.config_preview_screen import _build_yaml_preview

        _build_yaml_preview(service.draft)

        draft_after = service.draft
        assert draft_before == draft_after

    def test_full_preview_pipeline_is_read_only(self, service: DraftService) -> None:
        """Simulate what ConfigPreviewScreen.compose() does internally."""
        service.set_field("difficulty.error_density", 0.3)
        service.set_field("difficulty.canary_visibility", "subtle")
        service.set_field("output.enabled_test_cases", ["TC-01"])

        draft_before = service.draft
        merged_before = service.merged_raw

        # Simulate compose() reads
        from generator.tui.config_preview_screen import (
            _build_diff_text,
            _build_summary,
            _build_yaml_preview,
        )

        diffs = service.diff()
        draft = service.draft
        _build_summary(diffs)
        _build_diff_text(diffs)
        _build_yaml_preview(draft)

        assert service.draft == draft_before
        assert service.merged_raw == merged_before
