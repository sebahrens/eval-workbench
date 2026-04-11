"""Tests for generator.scenario_context — deterministic seed namespaces."""

from __future__ import annotations

import numpy as np

from generator.scenario_context import ScenarioContext


class TestStability:
    """Same name + same seed always produces the same stream."""

    def test_named_rng_reproducible(self) -> None:
        ctx1 = ScenarioContext(seed=42)
        ctx2 = ScenarioContext(seed=42)

        vals1 = ctx1.named_rng("general_ledger").random(10)
        vals2 = ctx2.named_rng("general_ledger").random(10)

        np.testing.assert_array_equal(vals1, vals2)

    def test_child_seed_reproducible(self) -> None:
        ctx1 = ScenarioContext(seed=42)
        ctx2 = ScenarioContext(seed=42)

        assert ctx1.child_seed("employees") == ctx2.child_seed("employees")

    def test_child_context_reproducible(self) -> None:
        ctx1 = ScenarioContext(seed=42).child("tc_08")
        ctx2 = ScenarioContext(seed=42).child("tc_08")

        vals1 = ctx1.named_rng("leases").random(5)
        vals2 = ctx2.named_rng("leases").random(5)

        np.testing.assert_array_equal(vals1, vals2)


class TestIsolation:
    """New namespaces do not perturb existing streams."""

    def test_adding_namespace_does_not_change_existing(self) -> None:
        ctx = ScenarioContext(seed=42)

        # Get stream for "gl" without any other calls
        vals_gl_alone = ScenarioContext(seed=42).named_rng("gl").random(10)

        # Now get "ar" first, then "gl"
        _ar_rng = ctx.named_rng("ar")
        vals_gl_after_ar = ctx.named_rng("gl").random(10)

        np.testing.assert_array_equal(vals_gl_alone, vals_gl_after_ar)

    def test_different_names_produce_different_streams(self) -> None:
        ctx = ScenarioContext(seed=42)
        vals_a = ctx.named_rng("alpha").random(10)
        vals_b = ctx.named_rng("beta").random(10)

        assert not np.array_equal(vals_a, vals_b)

    def test_different_seeds_produce_different_streams(self) -> None:
        vals1 = ScenarioContext(seed=1).named_rng("x").random(10)
        vals2 = ScenarioContext(seed=2).named_rng("x").random(10)

        assert not np.array_equal(vals1, vals2)

    def test_child_isolated_from_parent(self) -> None:
        ctx = ScenarioContext(seed=42)
        child = ctx.child("sub")

        parent_vals = ctx.named_rng("data").random(10)
        child_vals = child.named_rng("data").random(10)

        assert not np.array_equal(parent_vals, child_vals)


class TestBaseSeed:
    """Property access works."""

    def test_base_seed_accessible(self) -> None:
        ctx = ScenarioContext(seed=99)
        assert ctx.base_seed == 99

    def test_child_base_seed_differs_from_parent(self) -> None:
        ctx = ScenarioContext(seed=42)
        child = ctx.child("sub")
        assert child.base_seed != ctx.base_seed


class TestEdgeCases:
    """Boundary and edge-case inputs."""

    def test_empty_name(self) -> None:
        ctx = ScenarioContext(seed=42)
        # Should not raise — empty string is a valid namespace
        rng = ctx.named_rng("")
        assert rng.random() >= 0.0

    def test_unicode_name(self) -> None:
        ctx = ScenarioContext(seed=42)
        rng = ctx.named_rng("日本語テスト")
        assert rng.random() >= 0.0

    def test_zero_seed(self) -> None:
        ctx = ScenarioContext(seed=0)
        rng = ctx.named_rng("test")
        assert rng.random() >= 0.0

    def test_negative_seed(self) -> None:
        ctx = ScenarioContext(seed=-1)
        rng = ctx.named_rng("test")
        assert rng.random() >= 0.0
