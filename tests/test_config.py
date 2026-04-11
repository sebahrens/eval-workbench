"""Tests for generator.config — loading, validation, and round-trip."""

from __future__ import annotations

import textwrap
from pathlib import Path

import pytest

from generator.config import (
    Config,
    ConfigError,
    DifficultyProfile,
    ErrorProfile,
    OutputProfile,
    deep_merge,
    load_config,
    load_layered_config,
)
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
    text = minimal_yaml.read_text() + "packs: [accounting]\n"
    minimal_yaml.write_text(text)
    with pytest.raises(ConfigError, match="Unknown key.*config root.*packs"):
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
    text = minimal_yaml.read_text() + "packs: [accounting]\nworkflow: draft\n"
    minimal_yaml.write_text(text)
    with pytest.raises(ConfigError, match="Unknown key.*config root") as exc_info:
        load_config(minimal_yaml)
    msg = str(exc_info.value)
    assert "packs" in msg
    assert "workflow" in msg


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


# ---------------------------------------------------------------------------
# v1 customization profiles (synth-data-2u6.2)
# ---------------------------------------------------------------------------

class TestDifficultyProfile:
    def test_defaults(self) -> None:
        d = DifficultyProfile()
        assert d.error_density == 1.0
        assert d.canary_visibility == "visible"
        assert d.judgment_trap_density == 1.0

    def test_valid_values(self) -> None:
        d = DifficultyProfile(error_density=0.3, canary_visibility="hidden", judgment_trap_density=0.5)
        assert d.error_density == 0.3
        assert d.canary_visibility == "hidden"

    def test_error_density_out_of_range(self) -> None:
        with pytest.raises(ConfigError, match="error_density must be 0.0"):
            DifficultyProfile(error_density=1.5)

    def test_error_density_negative(self) -> None:
        with pytest.raises(ConfigError, match="error_density must be 0.0"):
            DifficultyProfile(error_density=-0.1)

    def test_invalid_canary_visibility(self) -> None:
        with pytest.raises(ConfigError, match="canary_visibility"):
            DifficultyProfile(canary_visibility="loud")

    def test_judgment_trap_density_out_of_range(self) -> None:
        with pytest.raises(ConfigError, match="judgment_trap_density"):
            DifficultyProfile(judgment_trap_density=2.0)


class TestOutputProfile:
    def test_defaults(self) -> None:
        o = OutputProfile()
        assert o.enabled_test_cases == []
        assert o.enabled_packs == []

    def test_with_values(self) -> None:
        o = OutputProfile(enabled_test_cases=["TC-01", "TC-06"], enabled_packs=["accounting"])
        assert o.enabled_test_cases == ["TC-01", "TC-06"]
        assert o.enabled_packs == ["accounting"]


class TestErrorProfile:
    def test_defaults(self) -> None:
        e = ErrorProfile()
        assert e.include == []
        assert e.exclude == []
        assert e.density_override is None

    def test_valid(self) -> None:
        e = ErrorProfile(include=["ERR-001"], exclude=["ERR-002"], density_override=0.5)
        assert e.include == ["ERR-001"]
        assert e.density_override == 0.5

    def test_density_out_of_range(self) -> None:
        with pytest.raises(ConfigError, match="density_override"):
            ErrorProfile(density_override=1.5)

    def test_overlap_rejected(self) -> None:
        with pytest.raises(ConfigError, match="overlap"):
            ErrorProfile(include=["ERR-001"], exclude=["ERR-001"])


# ---------------------------------------------------------------------------
# deep_merge tests (synth-data-2u6.2)
# ---------------------------------------------------------------------------

class TestDeepMerge:
    def test_scalar_override(self) -> None:
        assert deep_merge({"a": 1}, {"a": 2}) == {"a": 2}

    def test_add_new_key(self) -> None:
        assert deep_merge({"a": 1}, {"b": 2}) == {"a": 1, "b": 2}

    def test_nested_dict_merge(self) -> None:
        base = {"company": {"name": "Cascade", "revenue": 200}}
        overlay = {"company": {"name": "Acme"}}
        result = deep_merge(base, overlay)
        assert result == {"company": {"name": "Acme", "revenue": 200}}

    def test_list_replace(self) -> None:
        base = {"years": [2023, 2024, 2025]}
        overlay = {"years": [2024, 2025]}
        result = deep_merge(base, overlay)
        assert result == {"years": [2024, 2025]}

    def test_null_deletes(self) -> None:
        base = {"a": 1, "b": 2}
        overlay = {"b": None}
        result = deep_merge(base, overlay)
        assert result == {"a": 1}

    def test_null_deletes_nested(self) -> None:
        base = {"company": {"subs": {"a": 1, "b": 2}}}
        overlay = {"company": {"subs": {"b": None}}}
        result = deep_merge(base, overlay)
        assert result == {"company": {"subs": {"a": 1}}}

    def test_base_unchanged(self) -> None:
        base = {"a": 1, "b": {"c": 3}}
        overlay = {"b": {"c": 99}}
        deep_merge(base, overlay)
        assert base == {"a": 1, "b": {"c": 3}}

    def test_deeply_nested(self) -> None:
        base = {"l1": {"l2": {"l3": {"val": 1, "keep": True}}}}
        overlay = {"l1": {"l2": {"l3": {"val": 2}}}}
        result = deep_merge(base, overlay)
        assert result == {"l1": {"l2": {"l3": {"val": 2, "keep": True}}}}


# ---------------------------------------------------------------------------
# Config loading with profiles (synth-data-2u6.2)
# ---------------------------------------------------------------------------

def test_config_with_difficulty_section(minimal_yaml: Path) -> None:
    text = minimal_yaml.read_text() + textwrap.dedent("""\
        difficulty:
          error_density: 0.3
          canary_visibility: subtle
    """)
    minimal_yaml.write_text(text)
    cfg = load_config(minimal_yaml)
    assert cfg.difficulty.error_density == 0.3
    assert cfg.difficulty.canary_visibility == "subtle"
    assert cfg.difficulty.judgment_trap_density == 1.0  # default


def test_config_with_output_section(minimal_yaml: Path) -> None:
    text = minimal_yaml.read_text() + textwrap.dedent("""\
        output:
          enabled_test_cases: [TC-01, TC-06]
          enabled_packs: [accounting]
    """)
    minimal_yaml.write_text(text)
    cfg = load_config(minimal_yaml)
    assert cfg.output.enabled_test_cases == ["TC-01", "TC-06"]
    assert cfg.output.enabled_packs == ["accounting"]


def test_config_with_errors_section(minimal_yaml: Path) -> None:
    text = minimal_yaml.read_text() + textwrap.dedent("""\
        errors:
          include: [ERR-001]
          exclude: [ERR-002]
          density_override: 0.5
    """)
    minimal_yaml.write_text(text)
    cfg = load_config(minimal_yaml)
    assert cfg.errors.include == ["ERR-001"]
    assert cfg.errors.exclude == ["ERR-002"]
    assert cfg.errors.density_override == 0.5


def test_config_without_profiles_has_defaults(minimal_yaml: Path) -> None:
    cfg = load_config(minimal_yaml)
    assert cfg.difficulty == DifficultyProfile()
    assert cfg.output == OutputProfile()
    assert cfg.errors == ErrorProfile()


def test_difficulty_scalar_rejected(minimal_yaml: Path) -> None:
    text = minimal_yaml.read_text() + "difficulty: hard\n"
    minimal_yaml.write_text(text)
    with pytest.raises(ConfigError, match="difficulty.*must be a mapping"):
        load_config(minimal_yaml)


def test_unknown_difficulty_key_rejected(minimal_yaml: Path) -> None:
    text = minimal_yaml.read_text() + textwrap.dedent("""\
        difficulty:
          error_density: 0.5
          level: hard
    """)
    minimal_yaml.write_text(text)
    with pytest.raises(ConfigError, match="Unknown key.*difficulty.*level"):
        load_config(minimal_yaml)


def test_invalid_difficulty_value_rejected(minimal_yaml: Path) -> None:
    text = minimal_yaml.read_text() + textwrap.dedent("""\
        difficulty:
          error_density: 2.0
    """)
    minimal_yaml.write_text(text)
    with pytest.raises(ConfigError, match="error_density must be 0.0"):
        load_config(minimal_yaml)


# ---------------------------------------------------------------------------
# Layered config loading (synth-data-2u6.2)
# ---------------------------------------------------------------------------

def test_layered_no_overlays_matches_base(minimal_yaml: Path) -> None:
    base = load_config(minimal_yaml)
    layered = load_layered_config(minimal_yaml)
    assert layered.seed == base.seed
    assert layered.company.name == base.company.name


def test_layered_company_name_override(minimal_yaml: Path, tmp_path: Path) -> None:
    overlay = tmp_path / "overlay.yaml"
    overlay.write_text("company:\n  name: Acme Corp\n")
    cfg = load_layered_config(minimal_yaml, layers=[overlay])
    assert cfg.company.name == "Acme Corp"
    # Other company fields preserved from base
    assert cfg.company.type == "C-Corp"
    assert len(cfg.company.subsidiaries) == 1


def test_layered_seed_override(minimal_yaml: Path, tmp_path: Path) -> None:
    overlay = tmp_path / "overlay.yaml"
    overlay.write_text("seed: 99\n")
    cfg = load_layered_config(minimal_yaml, layers=[overlay])
    assert cfg.seed == 99


def test_layered_difficulty_overlay(minimal_yaml: Path, tmp_path: Path) -> None:
    overlay = tmp_path / "overlay.yaml"
    overlay.write_text(textwrap.dedent("""\
        difficulty:
          error_density: 0.3
          canary_visibility: hidden
    """))
    cfg = load_layered_config(minimal_yaml, layers=[overlay])
    assert cfg.difficulty.error_density == 0.3
    assert cfg.difficulty.canary_visibility == "hidden"


def test_layered_multiple_overlays(minimal_yaml: Path, tmp_path: Path) -> None:
    layer1 = tmp_path / "layer1.yaml"
    layer1.write_text("company:\n  name: Layer1 Corp\nseed: 10\n")
    layer2 = tmp_path / "layer2.yaml"
    layer2.write_text("seed: 20\n")  # Override seed again
    cfg = load_layered_config(minimal_yaml, layers=[layer1, layer2])
    assert cfg.company.name == "Layer1 Corp"  # From layer1
    assert cfg.seed == 20  # From layer2 (last wins)


def test_layered_null_deletes_subsidiary(minimal_yaml: Path, tmp_path: Path) -> None:
    # First add a second subsidiary to the base
    base_text = minimal_yaml.read_text().replace(
        "          employee_count: 100",
        "          employee_count: 100\n"
        "    sub_b:\n"
        "      legal_name: Sub B LLC\n"
        "      location: Austin, TX\n"
        "      state: TX\n"
        "      entity_code: SB\n"
        "      revenue: 50000000\n"
        "      type: Services\n"
        "      gross_margin: 0.40\n"
        "      employee_count: 50",
    )
    minimal_yaml.write_text(base_text)

    # Overlay that removes sub_b
    overlay = tmp_path / "overlay.yaml"
    overlay.write_text("company:\n  subsidiaries:\n    sub_b: null\n")
    cfg = load_layered_config(minimal_yaml, layers=[overlay])
    assert "sub_a" in cfg.company.subsidiaries
    assert "sub_b" not in cfg.company.subsidiaries


def test_layered_rejects_unknown_in_overlay(minimal_yaml: Path, tmp_path: Path) -> None:
    overlay = tmp_path / "overlay.yaml"
    overlay.write_text("packs: [accounting]\n")
    with pytest.raises(ConfigError, match="Unknown key.*config root"):
        load_layered_config(minimal_yaml, layers=[overlay])
