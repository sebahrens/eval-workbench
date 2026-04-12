"""Tests for CLI --overlay and --set flags (synth-data-2u6.3).

Covers parse_set_overrides, _coerce_value, and main() integration
with overlay and set flags.
"""

from __future__ import annotations

import textwrap
from pathlib import Path

import pytest

from generate_test_suite import _coerce_value, parse_set_overrides
from generator.config import ConfigError

# ---------------------------------------------------------------------------
# parse_set_overrides
# ---------------------------------------------------------------------------

class TestParseSetOverrides:
    def test_simple_scalar(self) -> None:
        assert parse_set_overrides(["seed=99"]) == {"seed": 99}

    def test_string_value(self) -> None:
        assert parse_set_overrides(["company.name=Acme"]) == {
            "company": {"name": "Acme"},
        }

    def test_deeply_nested(self) -> None:
        result = parse_set_overrides(["company.employees.total_count=500"])
        assert result == {"company": {"employees": {"total_count": 500}}}

    def test_multiple_pairs(self) -> None:
        result = parse_set_overrides(["seed=99", "company.name=Acme"])
        assert result == {"seed": 99, "company": {"name": "Acme"}}

    def test_float_value(self) -> None:
        result = parse_set_overrides(["difficulty.error_density=0.5"])
        assert result == {"difficulty": {"error_density": 0.5}}

    def test_bool_true(self) -> None:
        result = parse_set_overrides(["augmentation.enabled=true"])
        assert result == {"augmentation": {"enabled": True}}

    def test_bool_false(self) -> None:
        result = parse_set_overrides(["augmentation.enabled=false"])
        assert result == {"augmentation": {"enabled": False}}

    def test_null_value(self) -> None:
        result = parse_set_overrides(["company.subsidiaries.sub_b=null"])
        assert result == {"company": {"subsidiaries": {"sub_b": None}}}

    def test_value_with_equals_sign(self) -> None:
        """Values containing = are preserved (only split on first =)."""
        result = parse_set_overrides(["company.name=A=B"])
        assert result == {"company": {"name": "A=B"}}

    def test_missing_equals_raises(self) -> None:
        with pytest.raises(ConfigError, match="key=value"):
            parse_set_overrides(["no_equals_here"])

    def test_empty_key_raises(self) -> None:
        with pytest.raises(ConfigError, match="key must not be empty"):
            parse_set_overrides(["=value"])

    def test_empty_value_is_string(self) -> None:
        result = parse_set_overrides(["company.name="])
        assert result == {"company": {"name": ""}}

    def test_multiple_pairs_same_parent(self) -> None:
        """Multiple overrides under the same parent merge correctly."""
        result = parse_set_overrides([
            "company.name=Acme",
            "company.industry=Tech",
        ])
        assert result == {"company": {"name": "Acme", "industry": "Tech"}}


# ---------------------------------------------------------------------------
# _coerce_value
# ---------------------------------------------------------------------------

class TestCoerceValue:
    def test_int(self) -> None:
        assert _coerce_value("42") == 42

    def test_negative_int(self) -> None:
        assert _coerce_value("-5") == -5

    def test_float(self) -> None:
        assert _coerce_value("3.14") == 3.14

    def test_true(self) -> None:
        assert _coerce_value("true") is True
        assert _coerce_value("True") is True

    def test_false(self) -> None:
        assert _coerce_value("false") is False

    def test_null(self) -> None:
        assert _coerce_value("null") is None

    def test_plain_string(self) -> None:
        assert _coerce_value("hello") == "hello"


# ---------------------------------------------------------------------------
# CLI integration: --overlay and --set with main()
# ---------------------------------------------------------------------------

def _base_config_yaml() -> str:
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


def test_cli_overlay_changes_name(base_yaml: Path, tmp_path: Path) -> None:
    """--overlay applies file overlays via load_layered_config."""
    from generator.config import load_layered_config

    overlay = tmp_path / "overlay.yaml"
    overlay.write_text("company:\n  name: OverlayCo\n")
    cfg = load_layered_config(base_yaml, layers=[overlay])
    assert cfg.company.name == "OverlayCo"
    assert cfg.seed == 42  # unchanged


def test_cli_set_changes_seed(base_yaml: Path) -> None:
    """--set overrides are parsed and applied via load_layered_config."""
    from generator.config import load_layered_config

    set_dict = parse_set_overrides(["seed=99"])
    cfg = load_layered_config(base_yaml, set_overrides=set_dict)
    assert cfg.seed == 99


def test_cli_set_after_overlay(base_yaml: Path, tmp_path: Path) -> None:
    """--set takes precedence over --overlay (applied last)."""
    from generator.config import load_layered_config

    overlay = tmp_path / "overlay.yaml"
    overlay.write_text("company:\n  name: FromOverlay\n")
    set_dict = parse_set_overrides(["company.name=FromSet"])
    cfg = load_layered_config(base_yaml, layers=[overlay], set_overrides=set_dict)
    assert cfg.company.name == "FromSet"


def test_cli_no_overlay_no_set_uses_load_config(base_yaml: Path) -> None:
    """Without --overlay or --set, main uses load_config (backward compat)."""
    from generator.config import load_config

    cfg = load_config(base_yaml)
    assert cfg.company.name == "TestCo"
    assert cfg.seed == 42
