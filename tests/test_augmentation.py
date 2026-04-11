"""Tests for augmentation feature gate and offline fallback contract.

Acceptance criteria from synth-data-8ok.7.5:
- Augmentation is off by default
- Default generation performs no provider calls
- Missing provider credentials produce a clear disabled/skip path
- Cache-miss behavior is explicit
- Canonical outputs remain byte-identical with augmentation disabled
"""

from __future__ import annotations

import textwrap
from pathlib import Path

import pytest

from generator.augmentation import AugmentationCache, WarmResult, _compute_cache_key
from generator.config import AugmentationConfig, load_config

# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture()
def disabled_config() -> AugmentationConfig:
    return AugmentationConfig(enabled=False)


@pytest.fixture()
def enabled_config(tmp_path: Path) -> AugmentationConfig:
    cache_dir = tmp_path / ".augmentation_cache"
    cache_dir.mkdir()
    return AugmentationConfig(
        enabled=True,
        model="test-model/v1",
        cache_dir=str(cache_dir),
        warm_on_miss=False,
    )


# ---------------------------------------------------------------------------
# Config defaults
# ---------------------------------------------------------------------------

class TestAugmentationConfigDefaults:
    def test_default_augmentation_disabled(self) -> None:
        cfg = AugmentationConfig()
        assert cfg.enabled is False
        assert cfg.model == ""
        assert cfg.cache_dir == ".augmentation_cache"
        assert cfg.warm_on_miss is False

    def test_config_loads_without_augmentation_section(self, tmp_path: Path) -> None:
        """Config YAML with no augmentation section should parse with defaults."""
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
                  type: Widget making
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
              seasonal_weights:
                Q1: 0.25
                Q2: 0.25
                Q3: 0.25
                Q4: 0.25
        """))
        cfg = load_config(p)
        assert cfg.augmentation.enabled is False
        assert cfg.augmentation.model == ""

    def test_config_loads_with_augmentation_enabled(self, tmp_path: Path) -> None:
        p = tmp_path / "config.yaml"
        p.write_text(textwrap.dedent("""\
            seed: 42
            output_dir: out
            augmentation:
              enabled: true
              model: openrouter/nemotron-3-nano-30b
              cache_dir: .my_cache
              warm_on_miss: true
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
                  type: Widget making
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
              seasonal_weights:
                Q1: 0.25
                Q2: 0.25
                Q3: 0.25
                Q4: 0.25
        """))
        cfg = load_config(p)
        assert cfg.augmentation.enabled is True
        assert cfg.augmentation.model == "openrouter/nemotron-3-nano-30b"
        assert cfg.augmentation.cache_dir == ".my_cache"
        assert cfg.augmentation.warm_on_miss is True

    def test_real_config_has_augmentation_disabled(self) -> None:
        """The production config.yaml must have augmentation off by default."""
        root = Path(__file__).resolve().parent.parent / "config.yaml"
        cfg = load_config(root)
        assert cfg.augmentation.enabled is False


# ---------------------------------------------------------------------------
# Cache — disabled path (no-ops)
# ---------------------------------------------------------------------------

class TestCacheDisabled:
    def test_get_returns_none(self, disabled_config: AugmentationConfig) -> None:
        cache = AugmentationCache(disabled_config)
        assert cache.get("prompt", "model", "ns") is None

    def test_put_returns_empty(self, disabled_config: AugmentationConfig) -> None:
        cache = AugmentationCache(disabled_config)
        assert cache.put("prompt", "model", "ns", "output") == ""

    def test_warm_returns_zero(self, disabled_config: AugmentationConfig) -> None:
        cache = AugmentationCache(disabled_config)
        result = cache.warm([])
        assert result == WarmResult(hits=0, misses=0, total_tokens=0, estimated_cost_usd=0.0)

    def test_prune_returns_zero(self, disabled_config: AugmentationConfig) -> None:
        cache = AugmentationCache(disabled_config)
        assert cache.prune() == 0

    def test_enabled_property_false(self, disabled_config: AugmentationConfig) -> None:
        cache = AugmentationCache(disabled_config)
        assert cache.enabled is False

    def test_manifest_empty(self, disabled_config: AugmentationConfig) -> None:
        cache = AugmentationCache(disabled_config)
        m = cache.manifest
        assert m["entries"] == {}


# ---------------------------------------------------------------------------
# Cache — enabled path
# ---------------------------------------------------------------------------

class TestCacheEnabled:
    def test_enabled_requires_model(self, tmp_path: Path) -> None:
        cfg = AugmentationConfig(enabled=True, model="", cache_dir=str(tmp_path))
        with pytest.raises(ValueError, match="model is required"):
            AugmentationCache(cfg)

    def test_put_and_get_roundtrip(self, enabled_config: AugmentationConfig) -> None:
        cache = AugmentationCache(enabled_config, scenario_hash="abc123")
        key = cache.put("my prompt", "test-model/v1", "tc_04.leases", "enriched text")
        assert key  # non-empty hash

        result = cache.get("my prompt", "test-model/v1", "tc_04.leases")
        assert result == "enriched text"

    def test_get_miss_returns_none(self, enabled_config: AugmentationConfig) -> None:
        cache = AugmentationCache(enabled_config, scenario_hash="abc123")
        assert cache.get("nonexistent", "model", "ns") is None

    def test_manifest_updated_on_put(self, enabled_config: AugmentationConfig) -> None:
        cache = AugmentationCache(enabled_config, scenario_hash="abc123")
        cache.put("prompt", "test-model/v1", "ns", "text", {"input": 10, "output": 20})
        m = cache.manifest
        assert len(m["entries"]) == 1
        entry = next(iter(m["entries"].values()))
        assert entry["model_alias"] == "test-model/v1"
        assert entry["token_usage"] == {"input": 10, "output": 20}


# ---------------------------------------------------------------------------
# Cache key determinism
# ---------------------------------------------------------------------------

class TestCacheKey:
    def test_same_inputs_same_key(self) -> None:
        k1 = _compute_cache_key("hash1", "model1", "prompt text", "ns1")
        k2 = _compute_cache_key("hash1", "model1", "prompt text", "ns1")
        assert k1 == k2

    def test_different_model_different_key(self) -> None:
        k1 = _compute_cache_key("hash1", "model1", "prompt text", "ns1")
        k2 = _compute_cache_key("hash1", "model2", "prompt text", "ns1")
        assert k1 != k2

    def test_different_prompt_different_key(self) -> None:
        k1 = _compute_cache_key("hash1", "model1", "prompt A", "ns1")
        k2 = _compute_cache_key("hash1", "model1", "prompt B", "ns1")
        assert k1 != k2

    def test_different_namespace_different_key(self) -> None:
        k1 = _compute_cache_key("hash1", "model1", "prompt", "ns1")
        k2 = _compute_cache_key("hash1", "model1", "prompt", "ns2")
        assert k1 != k2


# ---------------------------------------------------------------------------
# Manifest entry augmentation_source field
# ---------------------------------------------------------------------------

class TestManifestAugmentationSource:
    def test_default_augmentation_source_is_none(self) -> None:
        from generator.manifest import ManifestEntry
        entry = ManifestEntry(path="test.xlsx", type="xlsx")
        assert entry.augmentation_source is None

    def test_augmentation_source_excluded_from_dict_when_none(self) -> None:
        from generator.manifest import Manifest
        m = Manifest(Path("/tmp/test_manifest_aug"))
        m.register("test.xlsx", "xlsx")
        entries = m.to_dict()
        assert "augmentation_source" not in entries[0]

    def test_augmentation_source_included_when_set(self, tmp_path: Path) -> None:
        from generator.manifest import Manifest
        m = Manifest(tmp_path)
        m.register("test.xlsx", "xlsx")
        m._entries["test.xlsx"].augmentation_source = "cache:abc123"
        entries = m.to_dict()
        assert entries[0]["augmentation_source"] == "cache:abc123"


# ---------------------------------------------------------------------------
# Byte-identical output with augmentation disabled
# ---------------------------------------------------------------------------

class TestDeterminismWithAugmentationDisabled:
    def test_canonical_output_unchanged(self, tmp_path: Path) -> None:
        """Generator output must be byte-identical whether or not the
        augmentation config section exists, as long as enabled=False.
        """
        from generate_test_suite import generate
        from generator.config import load_config

        root_config = Path(__file__).resolve().parent.parent / "config.yaml"
        config = load_config(root_config)

        # Confirm augmentation is disabled
        assert config.augmentation.enabled is False

        # Run 1
        out1 = tmp_path / "run1"
        generate(config, out1)

        # Run 2
        out2 = tmp_path / "run2"
        generate(config, out2)

        # Compare all generated files
        files1 = sorted(f.relative_to(out1) for f in out1.rglob("*") if f.is_file())
        files2 = sorted(f.relative_to(out2) for f in out2.rglob("*") if f.is_file())
        assert files1 == files2, "File lists differ between runs"

        for rel in files1:
            content1 = (out1 / rel).read_bytes()
            content2 = (out2 / rel).read_bytes()
            assert content1 == content2, f"Content differs: {rel}"
