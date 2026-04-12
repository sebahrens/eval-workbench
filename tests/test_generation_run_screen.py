"""Tests for generation validation and run flow.

Covers command construction logic without running actual generation.
Tests verify that build_generate_command produces correct CLI argument
lists for various configuration combinations.

Bead: synth-data-2u6.6.11
"""

from __future__ import annotations

import sys
import textwrap
from pathlib import Path

import pytest

from generator.tui.draft_service import DraftService
from generator.tui.generation_run_screen import (
    build_generate_command,
    format_command,
)

# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


def _base_config_yaml() -> str:
    """Minimal valid base config for tests."""
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
# build_generate_command tests
# ---------------------------------------------------------------------------


class TestBuildGenerateCommand:
    """Test CLI command construction for various config combinations."""

    def test_default_command(self) -> None:
        """Default args produce minimal command."""
        cmd = build_generate_command()
        assert cmd[0] == sys.executable
        assert cmd[1] == "generate_test_suite.py"
        assert "--output" in cmd
        assert "/tmp/test_suite" in cmd
        # No --config when using default
        assert "--config" not in cmd

    def test_custom_config_path(self) -> None:
        cmd = build_generate_command(config_path="custom.yaml")
        idx = cmd.index("--config")
        assert cmd[idx + 1] == "custom.yaml"

    def test_default_config_path_omitted(self) -> None:
        """config.yaml (the default) is not passed explicitly."""
        cmd = build_generate_command(config_path="config.yaml")
        assert "--config" not in cmd

    def test_overlay_included(self) -> None:
        cmd = build_generate_command(overlay="overrides/my.yaml")
        idx = cmd.index("--overlay")
        assert cmd[idx + 1] == "overrides/my.yaml"

    def test_no_overlay_when_none(self) -> None:
        cmd = build_generate_command(overlay=None)
        assert "--overlay" not in cmd

    def test_custom_output_dir(self) -> None:
        cmd = build_generate_command(output_dir="/tmp/custom_out")
        idx = cmd.index("--output")
        assert cmd[idx + 1] == "/tmp/custom_out"

    def test_packs_included(self) -> None:
        cmd = build_generate_command(
            packs=["cascade_accounting_core", "cascade_legal_hr_diligence"],
        )
        idx = cmd.index("--packs")
        assert cmd[idx + 1] == "cascade_accounting_core"
        assert cmd[idx + 2] == "cascade_legal_hr_diligence"

    def test_no_packs_when_none(self) -> None:
        cmd = build_generate_command(packs=None)
        assert "--packs" not in cmd

    def test_empty_packs_omitted(self) -> None:
        cmd = build_generate_command(packs=[])
        assert "--packs" not in cmd

    def test_all_options_combined(self) -> None:
        cmd = build_generate_command(
            config_path="base.yaml",
            overlay="override.yaml",
            output_dir="/tmp/full_run",
            packs=["all"],
        )
        assert "--config" in cmd
        assert "--overlay" in cmd
        assert "--output" in cmd
        assert "--packs" in cmd
        # Verify ordering: config, overlay, output, packs
        assert cmd.index("--config") < cmd.index("--overlay")
        assert cmd.index("--overlay") < cmd.index("--output")
        assert cmd.index("--output") < cmd.index("--packs")


# ---------------------------------------------------------------------------
# format_command tests
# ---------------------------------------------------------------------------


class TestFormatCommand:
    """Test shell-safe command formatting."""

    def test_simple_command(self) -> None:
        cmd = [sys.executable, "generate_test_suite.py", "--output", "/tmp/out"]
        result = format_command(cmd)
        # Should be a valid shell string
        assert "generate_test_suite.py" in result
        assert "--output" in result

    def test_paths_with_spaces_are_quoted(self) -> None:
        cmd = [sys.executable, "generate_test_suite.py", "--output", "/tmp/my output"]
        result = format_command(cmd)
        assert "'/tmp/my output'" in result

    def test_roundtrip_with_shlex(self) -> None:
        """Formatted command can be parsed back to the original list."""
        import shlex
        cmd = [
            sys.executable, "generate_test_suite.py",
            "--config", "my config.yaml",
            "--overlay", "override.yaml",
            "--output", "/tmp/out",
        ]
        result = format_command(cmd)
        parsed = shlex.split(result)
        assert parsed == cmd


# ---------------------------------------------------------------------------
# Integration with DraftService
# ---------------------------------------------------------------------------


class TestCommandFromService:
    """Test command construction using DraftService state."""

    def test_command_uses_service_base_path(self, service: DraftService) -> None:
        """Command includes the service's base config path."""
        cmd = build_generate_command(
            config_path=str(service._base_path),
            output_dir="/tmp/out",
        )
        assert str(service._base_path) in cmd

    def test_command_with_saved_overlay(
        self, service: DraftService, tmp_path: Path
    ) -> None:
        """After saving a draft, the overlay path appears in the command."""
        service.set_field("company.name", "CommandTestCo")
        overlay_path = tmp_path / "override.yaml"
        service.save(overlay_path)

        cmd = build_generate_command(
            config_path=str(service._base_path),
            overlay=str(overlay_path),
            output_dir="/tmp/gen_out",
        )
        assert "--overlay" in cmd
        assert str(overlay_path) in cmd
        assert "--output" in cmd
        assert "/tmp/gen_out" in cmd

    def test_validation_status_before_command(
        self, service: DraftService
    ) -> None:
        """Valid config lets generation proceed; invalid config is flagged."""
        result = service.validate()
        assert result.valid

        # Break the config
        service.set_field("company.seasonal_weights.Q1", 0.99)
        result = service.validate()
        assert not result.valid
