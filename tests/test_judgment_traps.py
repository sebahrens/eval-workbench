"""Tests for generator.judgment_traps — judgment trap registry.

Ref: bead synth-data-ups.10, spec §judgment traps.
"""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from generator.judgment_traps import (
    _VALID_RESPONSES,
    _VALID_TRAP_TYPES,
    JudgmentTrap,
    JudgmentTrapRegistry,
)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _make_trap(**overrides) -> JudgmentTrap:
    defaults = dict(
        trap_id="JT-001",
        test_case="TC-19",
        trap_type="summary_contradiction",
        description="Summary says 'no change of control risk' but Section 12.3 has a CoC clause",
        source_refs=("tc19_contract_acme", "tc19_acme_amendment_2025"),
        expected_response="flag",
        rationale="Agent must cross-reference summary against source clauses",
    )
    defaults.update(overrides)
    return JudgmentTrap(**defaults)


# ---------------------------------------------------------------------------
# JudgmentTrap dataclass
# ---------------------------------------------------------------------------

class TestJudgmentTrap:
    def test_required_fields(self):
        trap = _make_trap()
        assert trap.trap_id == "JT-001"
        assert trap.test_case == "TC-19"
        assert trap.trap_type == "summary_contradiction"
        assert trap.expected_response == "flag"
        assert trap.source_refs == ("tc19_contract_acme", "tc19_acme_amendment_2025")

    def test_frozen(self):
        trap = _make_trap()
        with pytest.raises(AttributeError):
            trap.trap_id = "JT-999"  # type: ignore[misc]


# ---------------------------------------------------------------------------
# JudgmentTrapRegistry — add / validation
# ---------------------------------------------------------------------------

class TestRegistryAdd:
    def test_add_valid(self):
        reg = JudgmentTrapRegistry()
        reg.add(_make_trap())
        assert len(reg.entries) == 1

    def test_reject_invalid_trap_type(self):
        reg = JudgmentTrapRegistry()
        with pytest.raises(ValueError, match="Invalid trap type"):
            reg.add(_make_trap(trap_type="bogus"))

    def test_reject_invalid_response(self):
        reg = JudgmentTrapRegistry()
        with pytest.raises(ValueError, match="Invalid expected_response"):
            reg.add(_make_trap(expected_response="ignore"))

    def test_reject_duplicate_id(self):
        reg = JudgmentTrapRegistry()
        reg.add(_make_trap())
        with pytest.raises(ValueError, match="Duplicate trap_id"):
            reg.add(_make_trap())

    def test_all_trap_types_accepted(self):
        reg = JudgmentTrapRegistry()
        for i, tt in enumerate(sorted(_VALID_TRAP_TYPES)):
            reg.add(_make_trap(trap_id=f"JT-{i:03d}", trap_type=tt))
        assert len(reg.entries) == len(_VALID_TRAP_TYPES)

    def test_all_responses_accepted(self):
        reg = JudgmentTrapRegistry()
        for i, resp in enumerate(sorted(_VALID_RESPONSES)):
            reg.add(_make_trap(trap_id=f"JT-{i:03d}", expected_response=resp))
        assert len(reg.entries) == len(_VALID_RESPONSES)


# ---------------------------------------------------------------------------
# JudgmentTrapRegistry — filtering
# ---------------------------------------------------------------------------

class TestRegistryFilters:
    @pytest.fixture()
    def populated_registry(self) -> JudgmentTrapRegistry:
        reg = JudgmentTrapRegistry()
        reg.add(_make_trap(trap_id="JT-001", test_case="TC-19",
                           trap_type="summary_contradiction",
                           expected_response="flag"))
        reg.add(_make_trap(trap_id="JT-002", test_case="TC-19",
                           trap_type="missing_evidence",
                           expected_response="do_not_assert"))
        reg.add(_make_trap(trap_id="JT-003", test_case="TC-20",
                           trap_type="stale_document",
                           expected_response="flag"))
        reg.add(_make_trap(trap_id="JT-004", test_case="TC-21",
                           trap_type="scope_boundary",
                           expected_response="deprioritize"))
        reg.add(_make_trap(trap_id="JT-005", test_case="TC-21",
                           trap_type="overconfident_conclusion",
                           expected_response="caveat"))
        return reg

    def test_by_test_case(self, populated_registry: JudgmentTrapRegistry):
        tc19 = populated_registry.by_test_case("TC-19")
        assert [t.trap_id for t in tc19] == ["JT-001", "JT-002"]

    def test_by_test_case_empty(self, populated_registry: JudgmentTrapRegistry):
        assert populated_registry.by_test_case("TC-99") == []

    def test_by_type(self, populated_registry: JudgmentTrapRegistry):
        stale = populated_registry.by_type("stale_document")
        assert [t.trap_id for t in stale] == ["JT-003"]

    def test_by_response(self, populated_registry: JudgmentTrapRegistry):
        flags = populated_registry.by_response("flag")
        assert [t.trap_id for t in flags] == ["JT-001", "JT-003"]

    def test_get(self, populated_registry: JudgmentTrapRegistry):
        assert populated_registry.get("JT-004").trap_type == "scope_boundary"

    def test_get_missing_raises(self, populated_registry: JudgmentTrapRegistry):
        with pytest.raises(KeyError):
            populated_registry.get("JT-999")


# ---------------------------------------------------------------------------
# JudgmentTrapRegistry — serialisation
# ---------------------------------------------------------------------------

class TestRegistrySerialization:
    def test_to_dict_deterministic_order(self):
        """Entries are sorted by trap_id regardless of insertion order."""
        reg = JudgmentTrapRegistry()
        reg.add(_make_trap(trap_id="JT-003"))
        reg.add(_make_trap(trap_id="JT-001"))
        reg.add(_make_trap(trap_id="JT-002"))
        result = reg.to_dict()
        assert [d["trap_id"] for d in result] == ["JT-001", "JT-002", "JT-003"]

    def test_to_dict_source_refs_are_lists(self):
        """Tuples in dataclass are serialised as JSON lists."""
        reg = JudgmentTrapRegistry()
        reg.add(_make_trap())
        result = reg.to_dict()
        assert isinstance(result[0]["source_refs"], list)

    def test_write_json_roundtrip(self, tmp_path: Path):
        reg = JudgmentTrapRegistry()
        reg.add(_make_trap(trap_id="JT-001"))
        reg.add(_make_trap(trap_id="JT-002", expected_response="caveat"))

        out = tmp_path / "judgment_traps.json"
        reg.write_json(out)

        data = json.loads(out.read_text())
        assert len(data) == 2
        assert data[0]["trap_id"] == "JT-001"
        assert data[1]["trap_id"] == "JT-002"

    def test_write_json_sorted_keys(self, tmp_path: Path):
        """JSON keys within each entry are sorted for determinism."""
        reg = JudgmentTrapRegistry()
        reg.add(_make_trap())

        out = tmp_path / "judgment_traps.json"
        reg.write_json(out)

        text = out.read_text()
        data = json.loads(text)
        keys = list(data[0].keys())
        assert keys == sorted(keys)

    def test_write_json_creates_parent_dirs(self, tmp_path: Path):
        reg = JudgmentTrapRegistry()
        reg.add(_make_trap())

        out = tmp_path / "nested" / "dir" / "traps.json"
        reg.write_json(out)
        assert out.exists()
