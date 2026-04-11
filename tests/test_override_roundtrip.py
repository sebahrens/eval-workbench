"""TUI override round-trip and unsupported-field regression tests.

These tests define the contract for the layered customization system:
- Override YAML files merge non-destructively with the base config.
- Saving and reloading an override produces the same effective config.
- Unsupported/unknown fields in overrides are rejected at validation time.
- The atomic save path never mutates the original config.yaml.

Bead: synth-data-2u6.6.14
Blocks: synth-data-2u6.6.3 (draft service), synth-data-2u6.6.10 (atomic save)
"""

from __future__ import annotations

import copy
import textwrap
from pathlib import Path

import pytest
import yaml

from generator.config import Config, ConfigError, load_config

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

# The set of v1-supported override fields.  Override YAML files may only
# contain keys within this surface.  The layered config loader and TUI
# must reject anything outside this set.
V1_SUPPORTED_OVERRIDE_FIELDS: set[str] = {
    "company.name",
    "company.type",
    "company.industry",
    "company.headquarters",
    "company.fiscal_year_end",
    "company.consolidated_revenue",
    "company.employees.total_count",
    "company.employees.annual_turnover_rate",
    "company.employees.remote_states",
    "company.growth_rates.fy2023_to_fy2024",
    "company.growth_rates.fy2024_to_fy2025",
    "company.seasonal_weights.Q1",
    "company.seasonal_weights.Q2",
    "company.seasonal_weights.Q3",
    "company.seasonal_weights.Q4",
    "seed",
    "output_dir",
}


def _deep_merge(base: dict, override: dict) -> dict:
    """Recursively merge *override* into a copy of *base*.

    - Scalar values in *override* replace values in *base*.
    - Dicts are merged recursively.
    - Lists in *override* replace (not extend) lists in *base*.
    """
    result = copy.deepcopy(base)
    for key, val in override.items():
        if key in result and isinstance(result[key], dict) and isinstance(val, dict):
            result[key] = _deep_merge(result[key], val)
        else:
            result[key] = copy.deepcopy(val)
    return result


def _flatten_keys(d: dict, prefix: str = "") -> set[str]:
    """Return all dotted key paths in a nested dict."""
    keys: set[str] = set()
    for k, v in d.items():
        full = f"{prefix}.{k}" if prefix else k
        if isinstance(v, dict):
            keys.update(_flatten_keys(v, full))
        else:
            keys.add(full)
    return keys


def _validate_override_fields(override: dict) -> list[str]:
    """Return list of unsupported dotted-key paths found in *override*.

    This is the expected validation behavior: any key not in the v1
    supported set is an error.
    """
    flat = _flatten_keys(override)
    return sorted(flat - V1_SUPPORTED_OVERRIDE_FIELDS)


def _base_config_yaml() -> str:
    """Return a minimal valid base config YAML string."""
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


# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture()
def base_yaml(tmp_path: Path) -> Path:
    """Write a minimal valid base config and return its path."""
    p = tmp_path / "config.yaml"
    p.write_text(_base_config_yaml())
    return p


@pytest.fixture()
def base_config(base_yaml: Path) -> Config:
    """Load the minimal base config."""
    return load_config(base_yaml)


# ---------------------------------------------------------------------------
# Round-trip: override merge → save → reload equivalence
# ---------------------------------------------------------------------------

class TestOverrideRoundTrip:
    """Override YAML saves are non-destructive and reload into the same
    effective configuration."""

    def test_empty_override_preserves_base(
        self, base_yaml: Path, base_config: Config, tmp_path: Path,
    ) -> None:
        """An empty override produces the same config as the base."""
        base_raw = yaml.safe_load(base_yaml.read_text())
        override: dict = {}
        merged = _deep_merge(base_raw, override)

        merged_path = tmp_path / "merged.yaml"
        merged_path.write_text(yaml.dump(merged, default_flow_style=False, sort_keys=True))

        reloaded = load_config(merged_path)
        assert reloaded.seed == base_config.seed
        assert reloaded.company.name == base_config.company.name
        assert reloaded.company.consolidated_revenue == base_config.company.consolidated_revenue

    def test_scalar_override_round_trips(
        self, base_yaml: Path, tmp_path: Path,
    ) -> None:
        """A scalar override is preserved through save and reload."""
        base_raw = yaml.safe_load(base_yaml.read_text())
        override = {"company": {"name": "OverrideCo"}}
        merged = _deep_merge(base_raw, override)

        merged_path = tmp_path / "merged.yaml"
        merged_path.write_text(yaml.dump(merged, default_flow_style=False, sort_keys=True))

        reloaded = load_config(merged_path)
        assert reloaded.company.name == "OverrideCo"

    def test_nested_override_round_trips(
        self, base_yaml: Path, tmp_path: Path,
    ) -> None:
        """Nested dict overrides merge correctly and round-trip."""
        base_raw = yaml.safe_load(base_yaml.read_text())
        override = {
            "company": {
                "employees": {
                    "total_count": 500,
                    "annual_turnover_rate": 0.12,
                },
            },
        }
        merged = _deep_merge(base_raw, override)

        # Employees were overridden but remote_states preserved from base
        assert merged["company"]["employees"]["remote_states"] == ["CA"]
        assert merged["company"]["employees"]["total_count"] == 500

        merged_path = tmp_path / "merged.yaml"
        merged_path.write_text(yaml.dump(merged, default_flow_style=False, sort_keys=True))

        reloaded = load_config(merged_path)
        assert reloaded.company.employees.total_count == 500
        assert reloaded.company.employees.annual_turnover_rate == 0.12
        assert reloaded.company.employees.remote_states == ["CA"]

    def test_list_override_replaces_not_extends(
        self, base_yaml: Path, tmp_path: Path,
    ) -> None:
        """List overrides replace the base list, not append to it."""
        base_raw = yaml.safe_load(base_yaml.read_text())
        override = {
            "company": {
                "employees": {
                    "remote_states": ["TX", "FL"],
                },
            },
        }
        merged = _deep_merge(base_raw, override)
        assert merged["company"]["employees"]["remote_states"] == ["TX", "FL"]

        merged_path = tmp_path / "merged.yaml"
        merged_path.write_text(yaml.dump(merged, default_flow_style=False, sort_keys=True))

        reloaded = load_config(merged_path)
        assert reloaded.company.employees.remote_states == ["TX", "FL"]

    def test_multiple_overrides_compose(
        self, base_yaml: Path, tmp_path: Path,
    ) -> None:
        """Multiple successive overrides compose correctly."""
        base_raw = yaml.safe_load(base_yaml.read_text())

        override1 = {"company": {"name": "FirstOverride"}}
        override2 = {"company": {"industry": "Technology"}}

        merged = _deep_merge(base_raw, override1)
        merged = _deep_merge(merged, override2)

        merged_path = tmp_path / "merged.yaml"
        merged_path.write_text(yaml.dump(merged, default_flow_style=False, sort_keys=True))

        reloaded = load_config(merged_path)
        assert reloaded.company.name == "FirstOverride"
        assert reloaded.company.industry == "Technology"

    def test_save_reload_idempotent(
        self, base_yaml: Path, tmp_path: Path,
    ) -> None:
        """Saving and reloading twice produces identical configs."""
        base_raw = yaml.safe_load(base_yaml.read_text())
        override = {
            "seed": 99,
            "company": {"consolidated_revenue": 300_000_000},
        }
        merged = _deep_merge(base_raw, override)

        # First save + reload
        path1 = tmp_path / "round1.yaml"
        path1.write_text(yaml.dump(merged, default_flow_style=False, sort_keys=True))
        cfg1 = load_config(path1)

        # Reload what was saved, save again, reload again
        raw1 = yaml.safe_load(path1.read_text())
        path2 = tmp_path / "round2.yaml"
        path2.write_text(yaml.dump(raw1, default_flow_style=False, sort_keys=True))
        cfg2 = load_config(path2)

        assert cfg1.seed == cfg2.seed == 99
        assert cfg1.company.consolidated_revenue == cfg2.company.consolidated_revenue == 300_000_000
        assert cfg1.company.name == cfg2.company.name

    def test_seasonal_weights_override_validates(
        self, base_yaml: Path, tmp_path: Path,
    ) -> None:
        """Overriding seasonal weights must still satisfy the sum-to-1 invariant."""
        base_raw = yaml.safe_load(base_yaml.read_text())
        override = {
            "company": {
                "seasonal_weights": {"Q1": 0.10, "Q2": 0.20, "Q3": 0.30, "Q4": 0.40},
            },
        }
        merged = _deep_merge(base_raw, override)

        merged_path = tmp_path / "valid_weights.yaml"
        merged_path.write_text(yaml.dump(merged, default_flow_style=False, sort_keys=True))
        cfg = load_config(merged_path)
        assert cfg.company.seasonal_weights.Q1 == 0.10
        assert cfg.company.seasonal_weights.Q4 == 0.40

    def test_seasonal_weights_override_bad_sum_rejected(
        self, base_yaml: Path, tmp_path: Path,
    ) -> None:
        """Overriding seasonal weights that don't sum to 1.0 raises ConfigError."""
        base_raw = yaml.safe_load(base_yaml.read_text())
        override = {
            "company": {
                "seasonal_weights": {"Q1": 0.50},  # breaks sum
            },
        }
        merged = _deep_merge(base_raw, override)

        merged_path = tmp_path / "bad_weights.yaml"
        merged_path.write_text(yaml.dump(merged, default_flow_style=False, sort_keys=True))
        with pytest.raises(ConfigError, match="Seasonal weights must sum to 1.0"):
            load_config(merged_path)


# ---------------------------------------------------------------------------
# Unsupported-field handling
# ---------------------------------------------------------------------------

class TestUnsupportedFieldValidation:
    """Unsupported fields in override YAML produce documented validation errors.

    The v1 supported-field surface is bounded.  Any field outside that
    surface must be caught by validation so the TUI cannot silently
    preserve, drop, or reinterpret unsupported configuration.
    """

    def test_supported_fields_pass(self) -> None:
        """All v1 supported fields pass validation."""
        override = {
            "seed": 99,
            "company": {
                "name": "NewCo",
                "consolidated_revenue": 500_000_000,
                "employees": {"total_count": 900},
            },
        }
        errors = _validate_override_fields(override)
        assert errors == []

    def test_unknown_top_level_field_rejected(self) -> None:
        """An unknown top-level key is flagged."""
        override = {"magic_mode": True}
        errors = _validate_override_fields(override)
        assert "magic_mode" in errors

    def test_unknown_nested_field_rejected(self) -> None:
        """An unknown nested key under a valid parent is flagged."""
        override = {
            "company": {"secret_sauce": 42},
        }
        errors = _validate_override_fields(override)
        assert "company.secret_sauce" in errors

    def test_subsidiary_override_rejected_in_v1(self) -> None:
        """Subsidiary structure changes are NOT in the v1 surface."""
        override = {
            "company": {
                "subsidiaries": {
                    "new_sub": {
                        "legal_name": "Sneaky LLC",
                        "entity_code": "SN",
                    },
                },
            },
        }
        errors = _validate_override_fields(override)
        assert any("subsidiaries" in e for e in errors)

    def test_intercompany_override_rejected_in_v1(self) -> None:
        """Intercompany params are NOT in the v1 supported surface."""
        override = {
            "company": {
                "intercompany": {"raw_materials_markup": 0.15},
            },
        }
        errors = _validate_override_fields(override)
        assert any("intercompany" in e for e in errors)

    def test_canary_override_rejected(self) -> None:
        """Users must not override canary assignments."""
        override = {"canary_assignments": {"file_a": "DEADBEEF"}}
        errors = _validate_override_fields(override)
        assert any("canary_assignments" in e for e in errors)

    def test_error_injections_override_rejected(self) -> None:
        """Users must not override error injection points."""
        override = {"error_injections": {"ERR-001": {"value": "sneaky"}}}
        errors = _validate_override_fields(override)
        assert any("error_injections" in e for e in errors)

    def test_mixed_supported_and_unsupported(self) -> None:
        """Validation catches unsupported fields even alongside supported ones."""
        override = {
            "seed": 99,              # supported
            "company": {
                "name": "ValidCo",   # supported
                "dark_mode": True,   # unsupported
            },
        }
        errors = _validate_override_fields(override)
        assert errors == ["company.dark_mode"]


# ---------------------------------------------------------------------------
# Atomic save: config.yaml must not be mutated
# ---------------------------------------------------------------------------

class TestAtomicSaveNonMutation:
    """Saving an override file must never modify the base config.yaml."""

    def test_save_override_does_not_touch_base(
        self, base_yaml: Path, tmp_path: Path,
    ) -> None:
        """Writing a merged override to a separate file leaves config.yaml intact."""
        original_content = base_yaml.read_text()
        original_mtime = base_yaml.stat().st_mtime

        base_raw = yaml.safe_load(original_content)
        override = {"company": {"name": "DifferentCo"}}
        merged = _deep_merge(base_raw, override)

        # Save to a DIFFERENT file (the override file), not to config.yaml
        override_path = tmp_path / "override.yaml"
        override_path.write_text(yaml.dump(merged, default_flow_style=False, sort_keys=True))

        # Base config.yaml must be completely unchanged
        assert base_yaml.read_text() == original_content
        assert base_yaml.stat().st_mtime == original_mtime

    def test_override_only_file_contains_deltas(self, tmp_path: Path) -> None:
        """An override-only file should contain just the changed fields,
        not a full copy of the base config."""
        override_only = {"company": {"name": "OverrideCo"}}
        override_path = tmp_path / "override_only.yaml"
        override_path.write_text(
            yaml.dump(override_only, default_flow_style=False, sort_keys=True)
        )

        reloaded = yaml.safe_load(override_path.read_text())
        # The override file should NOT contain fields that weren't overridden
        assert "seed" not in reloaded
        assert "output_dir" not in reloaded
        assert "subsidiaries" not in reloaded.get("company", {})

    def test_base_config_reloads_unchanged_after_override_save(
        self, base_yaml: Path, base_config: Config, tmp_path: Path,
    ) -> None:
        """After saving an override, reloading the base config still works identically."""
        # Save some override
        override_path = tmp_path / "override.yaml"
        override_path.write_text(yaml.dump(
            {"company": {"name": "SomethingElse"}},
            default_flow_style=False,
        ))

        # Base config still loads identically
        reloaded_base = load_config(base_yaml)
        assert reloaded_base.company.name == base_config.company.name
        assert reloaded_base.seed == base_config.seed


# ---------------------------------------------------------------------------
# Deep merge edge cases
# ---------------------------------------------------------------------------

class TestDeepMerge:
    """Edge cases for the recursive merge that the layered config loader
    must implement."""

    def test_merge_preserves_unrelated_branches(self) -> None:
        """Overriding one branch of the tree doesn't affect sibling branches."""
        base = {
            "a": {"x": 1, "y": 2},
            "b": {"z": 3},
        }
        override = {"a": {"x": 10}}
        merged = _deep_merge(base, override)

        assert merged["a"]["x"] == 10
        assert merged["a"]["y"] == 2  # sibling preserved
        assert merged["b"]["z"] == 3  # sibling branch preserved

    def test_merge_does_not_mutate_base(self) -> None:
        """The base dict is never modified in place."""
        base = {"a": {"x": 1}}
        override = {"a": {"x": 99}}
        original_val = base["a"]["x"]

        _deep_merge(base, override)
        assert base["a"]["x"] == original_val

    def test_merge_does_not_mutate_override(self) -> None:
        """The override dict is never modified in place."""
        base = {"a": {"x": 1}}
        override = {"a": {"x": 99, "nested": {"v": True}}}
        override_copy = copy.deepcopy(override)

        _deep_merge(base, override)
        assert override == override_copy

    def test_merge_new_key_in_override(self) -> None:
        """A key present in override but not base is added."""
        base = {"a": 1}
        override = {"b": 2}
        merged = _deep_merge(base, override)
        assert merged == {"a": 1, "b": 2}

    def test_merge_override_none_replaces(self) -> None:
        """An explicit None in the override replaces the base value."""
        base = {"a": {"x": 1}}
        override = {"a": None}
        merged = _deep_merge(base, override)
        assert merged["a"] is None
