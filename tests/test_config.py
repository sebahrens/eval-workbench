"""Tests for generator.config — loading, validation, and round-trip."""

from __future__ import annotations

import textwrap
from pathlib import Path

import pytest

from generator.config import Config, ConfigError, load_config
from generator.model.entities import (
    ENTITIES,
    SUBSIDIARIES,
    entities_from_config,
)

# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture()
def default_config(tmp_path: Path) -> Config:
    """Load the real config.yaml from the project root."""
    root = Path(__file__).resolve().parent.parent / "config.yaml"
    return load_config(root)


@pytest.fixture()
def minimal_yaml(tmp_path: Path) -> Path:
    """Write a minimal valid config and return its path."""
    p = tmp_path / "config.yaml"
    p.write_text(textwrap.dedent("""\
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
    """))
    return p


# ---------------------------------------------------------------------------
# Happy-path tests
# ---------------------------------------------------------------------------

def test_load_default_config(default_config: Config) -> None:
    """The real config.yaml loads without error and has expected shape."""
    assert default_config.seed == 42
    assert default_config.output_dir == "test_suite"
    assert default_config.company.name == "Cascade Industries, Inc."
    assert len(default_config.company.subsidiaries) == 3
    assert default_config.company.current_year == 2025


def test_subsidiary_fields(default_config: Config) -> None:
    pc = default_config.company.subsidiaries["precision_components"]
    assert pc.entity_code == "PC"
    assert pc.gross_margin == 0.35
    assert pc.rd_spend_pct == 0.0

    am = default_config.company.subsidiaries["advanced_materials"]
    assert am.rd_spend_pct == 0.12


def test_growth_rates(default_config: Config) -> None:
    gr = default_config.company.growth_rates
    assert gr.fy2023_to_fy2024 == 0.06
    assert gr.fy2024_to_fy2025 == 0.09


def test_seasonal_weights_sum(default_config: Config) -> None:
    sw = default_config.company.seasonal_weights
    assert abs(sw.Q1 + sw.Q2 + sw.Q3 + sw.Q4 - 1.0) < 1e-6


def test_empty_canary_and_error_dicts(default_config: Config) -> None:
    assert default_config.canary_assignments == {}
    assert default_config.error_injections == {}


def test_minimal_config_loads(minimal_yaml: Path) -> None:
    cfg = load_config(minimal_yaml)
    assert cfg.seed == 42
    assert len(cfg.company.subsidiaries) == 1


def test_config_round_trip(default_config: Config) -> None:
    """Verify the loaded config carries all expected intercompany params."""
    ic = default_config.company.intercompany
    assert ic.raw_materials_markup == 0.08
    assert ic.management_fee_pct == 0.015
    assert ic.intercompany_loan_principal == 5_000_000
    assert ic.intercompany_loan_rate == 0.05


# ---------------------------------------------------------------------------
# Error-path tests
# ---------------------------------------------------------------------------

def test_missing_file_raises(tmp_path: Path) -> None:
    with pytest.raises(ConfigError, match="not found"):
        load_config(tmp_path / "nope.yaml")


def test_invalid_yaml_raises(tmp_path: Path) -> None:
    p = tmp_path / "bad.yaml"
    p.write_text("seed: [unterminated")
    with pytest.raises(ConfigError, match="Invalid YAML"):
        load_config(p)


def test_non_mapping_raises(tmp_path: Path) -> None:
    p = tmp_path / "list.yaml"
    p.write_text("- a\n- b\n")
    with pytest.raises(ConfigError, match="must be a mapping"):
        load_config(p)


def test_missing_top_key_raises(tmp_path: Path) -> None:
    p = tmp_path / "partial.yaml"
    p.write_text("seed: 42\noutput_dir: out\n")
    with pytest.raises(ConfigError, match="company"):
        load_config(p)


def test_missing_subsidiary_key_raises(minimal_yaml: Path) -> None:
    import yaml

    raw = yaml.safe_load(minimal_yaml.read_text())
    del raw["company"]["subsidiaries"]["sub_a"]["entity_code"]
    minimal_yaml.write_text(yaml.dump(raw))
    with pytest.raises(ConfigError, match="entity_code"):
        load_config(minimal_yaml)


def test_bad_seasonal_weights_raises(minimal_yaml: Path) -> None:
    text = minimal_yaml.read_text().replace("Q4: 0.30", "Q4: 0.99")
    minimal_yaml.write_text(text)
    with pytest.raises(ConfigError, match="Seasonal weights must sum to 1.0"):
        load_config(minimal_yaml)


# ---------------------------------------------------------------------------
# v1 compatibility — unknown-field rejection (synth-data-2u6.7)
# ---------------------------------------------------------------------------

def test_unknown_top_level_key_rejected(minimal_yaml: Path) -> None:
    """Unknown top-level keys are rejected, not silently ignored."""
    text = minimal_yaml.read_text() + "difficulty: hard\n"
    minimal_yaml.write_text(text)
    with pytest.raises(ConfigError, match="Unknown key.*config root.*difficulty"):
        load_config(minimal_yaml)


def test_unknown_company_key_rejected(minimal_yaml: Path) -> None:
    """Unknown keys inside company section are rejected."""
    text = minimal_yaml.read_text().replace(
        "  name: TestCo",
        "  name: TestCo\n  product_lines: [Widgets]",
    )
    minimal_yaml.write_text(text)
    with pytest.raises(ConfigError, match="Unknown key.*company.*product_lines"):
        load_config(minimal_yaml)


def test_unknown_subsidiary_key_rejected(minimal_yaml: Path) -> None:
    """Unknown keys inside a subsidiary are rejected."""
    text = minimal_yaml.read_text().replace(
        "      employee_count: 100",
        "      employee_count: 100\n      salary_range: [50000, 150000]",
    )
    minimal_yaml.write_text(text)
    with pytest.raises(ConfigError, match="Unknown key.*subsidiaries.sub_a.*salary_range"):
        load_config(minimal_yaml)


def test_multiple_unknown_top_keys_all_reported(minimal_yaml: Path) -> None:
    """All unknown top-level keys appear in the error, not just the first."""
    text = minimal_yaml.read_text() + "difficulty: hard\npacks: [accounting]\n"
    minimal_yaml.write_text(text)
    with pytest.raises(ConfigError, match="Unknown key.*config root") as exc_info:
        load_config(minimal_yaml)
    msg = str(exc_info.value)
    assert "difficulty" in msg
    assert "packs" in msg


def test_build_time_fields_accepted(minimal_yaml: Path) -> None:
    """canary_assignments and error_injections are build-time fields, not user-editable,
    but must be accepted in config.yaml because the generator populates them."""
    text = minimal_yaml.read_text() + "canary_assignments: {}\nerror_injections: {}\n"
    minimal_yaml.write_text(text)
    cfg = load_config(minimal_yaml)
    assert cfg.canary_assignments == {}
    assert cfg.error_injections == {}


def test_augmentation_field_accepted(minimal_yaml: Path) -> None:
    """augmentation is an allowed top-level key."""
    text = minimal_yaml.read_text() + "augmentation:\n  enabled: false\n"
    minimal_yaml.write_text(text)
    cfg = load_config(minimal_yaml)
    assert cfg.augmentation.enabled is False


def test_rd_spend_pct_accepted_in_subsidiary(minimal_yaml: Path) -> None:
    """rd_spend_pct is an optional subsidiary field and must not be rejected."""
    text = minimal_yaml.read_text().replace(
        "      employee_count: 100",
        "      employee_count: 100\n      rd_spend_pct: 0.05",
    )
    minimal_yaml.write_text(text)
    cfg = load_config(minimal_yaml)
    assert cfg.company.subsidiaries["sub_a"].rd_spend_pct == 0.05


# ---------------------------------------------------------------------------
# v1 compatibility — legacy config backward compatibility (synth-data-2u6.7)
# ---------------------------------------------------------------------------

def test_legacy_config_loads_without_mutation() -> None:
    """The real config.yaml (the legacy format) loads cleanly and produces the
    expected default values. This is the backward-compatibility contract: existing
    config.yaml must remain a valid v1 config with no changes."""
    root = Path(__file__).resolve().parent.parent / "config.yaml"
    cfg = load_config(root)

    # Core defaults preserved
    assert cfg.seed == 42
    assert cfg.output_dir == "test_suite"
    assert cfg.company.name == "Cascade Industries, Inc."
    assert cfg.company.consolidated_revenue == 200_000_000
    assert len(cfg.company.subsidiaries) == 3
    assert set(cfg.company.subsidiaries) == {
        "precision_components", "advanced_materials", "distribution_services",
    }

    # Seasonal weights unchanged
    sw = cfg.company.seasonal_weights
    assert (sw.Q1, sw.Q2, sw.Q3, sw.Q4) == (0.20, 0.25, 0.25, 0.30)

    # Build-time fields default to empty dicts
    assert cfg.canary_assignments == {}
    assert cfg.error_injections == {}


def test_config_without_version_treated_as_v1(minimal_yaml: Path) -> None:
    """Configs without a config_version field are treated as v1. No error
    is raised — the version marker is deferred to v2."""
    cfg = load_config(minimal_yaml)
    assert cfg.seed == 42  # Just verify it loads — no version attribute in v1


def test_v1_custom_config_does_not_mutate_defaults(minimal_yaml: Path) -> None:
    """Loading a custom config does not affect the default config values.
    This verifies no shared mutable state between config loads."""
    root = Path(__file__).resolve().parent.parent / "config.yaml"
    default_before = load_config(root)

    # Load a different config in between
    custom = load_config(minimal_yaml)
    assert custom.company.name == "TestCo"

    # Re-load defaults — must be identical
    default_after = load_config(root)
    assert default_after.seed == default_before.seed
    assert default_after.company.name == default_before.company.name
    assert default_after.company.consolidated_revenue == default_before.company.consolidated_revenue
    assert len(default_after.company.subsidiaries) == len(default_before.company.subsidiaries)


# ---------------------------------------------------------------------------
# entities_from_config adapter tests
# ---------------------------------------------------------------------------

def test_default_config_matches_hardcoded_entities(default_config: Config) -> None:
    """Default Cascade config must produce the same entities as the module constants."""
    all_ents, subs = entities_from_config(default_config.company)

    # Same entity codes
    assert set(all_ents) == set(ENTITIES)
    assert set(subs) == set(SUBSIDIARIES)

    # Each hardcoded entity matches its config-derived counterpart
    for code, expected in ENTITIES.items():
        derived = all_ents[code]
        assert derived.code == expected.code
        assert derived.name == expected.name
        assert derived.location == expected.location
        assert derived.state == expected.state
        assert derived.revenue_target == expected.revenue_target
        assert derived.gross_margin == expected.gross_margin
        assert derived.is_parent == expected.is_parent


def test_custom_config_produces_custom_entities(minimal_yaml: Path) -> None:
    """A minimal custom config produces entities with the right codes/names."""
    cfg = load_config(minimal_yaml)
    all_ents, subs = entities_from_config(cfg.company)

    # One subsidiary (SA) plus auto-derived parent
    assert len(subs) == 1
    assert "SA" in subs
    assert subs["SA"].name == "Sub A LLC"
    assert subs["SA"].revenue_target == 100_000_000
    assert subs["SA"].gross_margin == 0.35

    # Parent exists
    parent_codes = [c for c, e in all_ents.items() if e.is_parent]
    assert len(parent_codes) == 1
    parent = all_ents[parent_codes[0]]
    assert parent.name == "TestCo"
    assert parent.revenue_target == 100_000_000
    assert parent.is_parent is True


def test_build_model_uses_config_entities(default_config: Config) -> None:
    """build_model with config populates model.entities and model.subsidiaries."""
    from generator.model.build import build_model

    model = build_model(default_config)
    assert set(model.entities) == set(ENTITIES)
    assert set(model.subsidiaries) == set(SUBSIDIARIES)
    assert model.entities["CI"].is_parent is True
    assert "CI" not in model.subsidiaries


def test_build_model_without_config_uses_hardcoded() -> None:
    """build_model(seed=42) falls back to hardcoded entity constants."""
    from generator.model.build import build_model

    model = build_model(seed=42)
    assert model.entities is ENTITIES
    assert model.subsidiaries is SUBSIDIARIES
