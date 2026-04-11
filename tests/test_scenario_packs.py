"""Tests for the scenario pack registry — lookup, ordering, dependencies, validation."""

from __future__ import annotations

import pytest

from generator.config import ConfigError
from generator.packs import (
    REGISTRY,
    ScenarioPack,
    collect_canary_keys,
    collect_test_case_count,
    get_pack,
    list_packs,
    register_pack,
    resolve_packs,
)

# ---------------------------------------------------------------------------
# Default accounting-core registration
# ---------------------------------------------------------------------------


class TestDefaultRegistration:
    """The accounting-core pack auto-registers on import."""

    def test_accounting_core_registered(self):
        assert "cascade_accounting_core" in REGISTRY

    def test_accounting_core_test_cases(self):
        pack = REGISTRY["cascade_accounting_core"]
        # TC-01 through TC-18
        assert len(pack.test_cases) == 18
        assert pack.test_cases[0] == "TC-01"
        assert pack.test_cases[-1] == "TC-18"

    def test_accounting_core_has_emitters(self):
        pack = REGISTRY["cascade_accounting_core"]
        assert len(pack.emitters) > 0

    def test_accounting_core_has_canary_keys(self):
        pack = REGISTRY["cascade_accounting_core"]
        assert len(pack.canary_file_keys) > 0


# ---------------------------------------------------------------------------
# Pack lookup
# ---------------------------------------------------------------------------


class TestGetPack:
    def test_valid_lookup(self):
        pack = get_pack("cascade_accounting_core")
        assert pack.pack_id == "cascade_accounting_core"

    def test_unknown_pack_raises(self):
        with pytest.raises(ConfigError, match="unknown pack.*nonexistent_pack"):
            get_pack("nonexistent_pack")

    def test_unknown_pack_lists_valid(self):
        with pytest.raises(ConfigError, match="valid packs:"):
            get_pack("no_such_pack")


# ---------------------------------------------------------------------------
# list_packs
# ---------------------------------------------------------------------------


class TestListPacks:
    def test_returns_all_registered(self):
        packs = list_packs()
        ids = [p.pack_id for p in packs]
        assert "cascade_accounting_core" in ids

    def test_insertion_order_preserved(self):
        packs = list_packs()
        # First registered is accounting_core
        assert packs[0].pack_id == "cascade_accounting_core"


# ---------------------------------------------------------------------------
# resolve_packs
# ---------------------------------------------------------------------------


class TestResolvePacks:
    def test_none_returns_defaults(self):
        packs = resolve_packs(None)
        assert len(packs) == 1
        assert packs[0].pack_id == "cascade_accounting_core"

    def test_explicit_pack_id(self):
        packs = resolve_packs(["cascade_accounting_core"])
        assert len(packs) == 1
        assert packs[0].pack_id == "cascade_accounting_core"

    def test_all_keyword(self):
        packs = resolve_packs(["all"])
        ids = [p.pack_id for p in packs]
        assert "cascade_accounting_core" in ids

    def test_unknown_pack_in_list_raises(self):
        with pytest.raises(ConfigError, match="unknown pack"):
            resolve_packs(["cascade_accounting_core", "bogus_pack"])


# ---------------------------------------------------------------------------
# Duplicate pack ID detection
# ---------------------------------------------------------------------------


class TestDuplicateRegistration:
    def test_duplicate_raises(self):
        dup = ScenarioPack(
            pack_id="cascade_accounting_core",
            display_name="Duplicate",
            test_cases=[],
            canary_file_keys=[],
            emitters=[],
        )
        with pytest.raises(ConfigError, match="duplicate pack_id"):
            register_pack(dup)


# ---------------------------------------------------------------------------
# Dependency validation
# ---------------------------------------------------------------------------


class TestDependencyValidation:
    def test_missing_dependency_raises(self):
        """Selecting a pack whose dependency is not also selected should fail."""
        # Temporarily register a pack that depends on a non-selected pack
        dep_pack = ScenarioPack(
            pack_id="_test_dep_child",
            display_name="Test Child",
            test_cases=["TC-99"],
            canary_file_keys=[],
            emitters=[],
            dependencies=["_test_dep_parent"],
        )
        parent_pack = ScenarioPack(
            pack_id="_test_dep_parent",
            display_name="Test Parent",
            test_cases=["TC-98"],
            canary_file_keys=[],
            emitters=[],
        )
        # Register both
        REGISTRY["_test_dep_child"] = dep_pack
        REGISTRY["_test_dep_parent"] = parent_pack
        try:
            # Select child without parent → should fail
            with pytest.raises(ConfigError, match="depends on.*_test_dep_parent"):
                resolve_packs(["_test_dep_child"])
        finally:
            del REGISTRY["_test_dep_child"]
            del REGISTRY["_test_dep_parent"]

    def test_satisfied_dependency_resolves(self):
        """Selecting both parent and child succeeds and orders correctly."""
        dep_pack = ScenarioPack(
            pack_id="_test_dep_child2",
            display_name="Test Child",
            test_cases=["TC-99"],
            canary_file_keys=[],
            emitters=[],
            dependencies=["_test_dep_parent2"],
        )
        parent_pack = ScenarioPack(
            pack_id="_test_dep_parent2",
            display_name="Test Parent",
            test_cases=["TC-98"],
            canary_file_keys=[],
            emitters=[],
        )
        REGISTRY["_test_dep_child2"] = dep_pack
        REGISTRY["_test_dep_parent2"] = parent_pack
        try:
            packs = resolve_packs(["_test_dep_child2", "_test_dep_parent2"])
            ids = [p.pack_id for p in packs]
            assert ids.index("_test_dep_parent2") < ids.index("_test_dep_child2")
        finally:
            del REGISTRY["_test_dep_child2"]
            del REGISTRY["_test_dep_parent2"]

    def test_circular_dependency_raises(self):
        """Circular deps between selected packs should fail."""
        pack_a = ScenarioPack(
            pack_id="_test_circ_a",
            display_name="A",
            test_cases=[],
            canary_file_keys=[],
            emitters=[],
            dependencies=["_test_circ_b"],
        )
        pack_b = ScenarioPack(
            pack_id="_test_circ_b",
            display_name="B",
            test_cases=[],
            canary_file_keys=[],
            emitters=[],
            dependencies=["_test_circ_a"],
        )
        REGISTRY["_test_circ_a"] = pack_a
        REGISTRY["_test_circ_b"] = pack_b
        try:
            with pytest.raises(ConfigError, match="circular dependency"):
                resolve_packs(["_test_circ_a", "_test_circ_b"])
        finally:
            del REGISTRY["_test_circ_a"]
            del REGISTRY["_test_circ_b"]


# ---------------------------------------------------------------------------
# Deterministic ordering
# ---------------------------------------------------------------------------


class TestDeterministicOrdering:
    def test_resolve_order_stable(self):
        """Multiple calls to resolve_packs return the same order."""
        r1 = [p.pack_id for p in resolve_packs(None)]
        r2 = [p.pack_id for p in resolve_packs(None)]
        assert r1 == r2

    def test_canary_keys_sorted(self):
        """collect_canary_keys returns sorted keys."""
        packs = resolve_packs(None)
        keys = collect_canary_keys(packs)
        assert keys == sorted(keys)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


class TestHelpers:
    def test_collect_test_case_count(self):
        packs = resolve_packs(None)
        count = collect_test_case_count(packs)
        assert count == 18

    def test_collect_canary_keys_nonempty(self):
        packs = resolve_packs(None)
        keys = collect_canary_keys(packs)
        assert len(keys) > 0
