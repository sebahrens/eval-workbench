"""Tests for generator.golds.framework — gold standard emission framework."""

import json
import tempfile
from pathlib import Path
from typing import Any

from generator.canaries import CanaryEntry, CanaryRegistry
from generator.errors import ErrorRegistry, PlantedError
from generator.golds.framework import (
    _REGISTRY,
    GoldStandard,
    _sort_recursive,
    emit_all_golds,
    emit_gold,
    read_gold_json,
    register_gold,
    registered_test_cases,
    verify_round_trip,
)

# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

def _make_canary_registry() -> CanaryRegistry:
    reg = CanaryRegistry()
    reg.entries["cascade_tb_fy2025"] = CanaryEntry(
        file_key="cascade_tb_fy2025", canary="XK7P2M9Q",
    )
    reg.entries["bank_stmt_dec2025"] = CanaryEntry(
        file_key="bank_stmt_dec2025", canary="LM3N8R2T",
    )
    return reg


def _make_error_registry() -> ErrorRegistry:
    reg = ErrorRegistry()
    reg.add(PlantedError(
        error_id="ERR-001",
        file="cascade_tb_fy2025.xlsx",
        location="Sheet 'Trial Balance', Cell G47",
        type="transposed_digits",
        description="AR balance transposed digits: $9,000 error",
        severity="material",
        which_test_cases_should_catch=["TC-01"],
    ))
    return reg


# ---------------------------------------------------------------------------
# GoldStandard dataclass
# ---------------------------------------------------------------------------

class TestGoldStandard:

    def test_to_dict_keys(self) -> None:
        gold = GoldStandard(
            test_case="TC-01",
            expected_outputs={"total": 42},
            canary_verification={"read_tb": "XK7P2M9Q"},
            error_detection={"ERR-001": "transposed digits"},
        )
        d = gold.to_dict()
        assert d["test_case"] == "TC-01"
        assert "expected_outputs" in d
        assert "canary_verification" in d
        assert "error_detection" in d
        # scoring_hints omitted when empty
        assert "scoring_hints" not in d

    def test_to_dict_includes_scoring_hints_when_present(self) -> None:
        gold = GoldStandard(
            test_case="TC-02",
            scoring_hints={"correctness": "exact match required"},
        )
        d = gold.to_dict()
        assert "scoring_hints" in d

    def test_from_dict_round_trip(self) -> None:
        gold = GoldStandard(
            test_case="TC-01",
            expected_outputs={"sheets": ["Mapping", "Variance"]},
            canary_verification={"a": "12345678"},
            error_detection={"ERR-001": "desc"},
            scoring_hints={"correctness": "strict"},
        )
        d = gold.to_dict()
        restored = GoldStandard.from_dict(d)
        assert restored.to_dict() == d

    def test_deterministic_key_order(self) -> None:
        """Keys in expected_outputs must be sorted regardless of insertion order."""
        gold = GoldStandard(
            test_case="TC-01",
            expected_outputs={"z_field": 1, "a_field": 2, "m_field": 3},
        )
        d = gold.to_dict()
        keys = list(d["expected_outputs"].keys())
        assert keys == sorted(keys)


# ---------------------------------------------------------------------------
# Registration
# ---------------------------------------------------------------------------

class TestRegistration:

    def setup_method(self) -> None:
        """Clear the registry between tests to avoid leaking state."""
        self._saved = dict(_REGISTRY)
        _REGISTRY.clear()

    def teardown_method(self) -> None:
        _REGISTRY.clear()
        _REGISTRY.update(self._saved)

    def test_register_and_list(self) -> None:
        @register_gold("TC-99")
        def tc99(canaries: Any, errors: Any, **kw: Any) -> GoldStandard:
            return GoldStandard(test_case="TC-99")

        assert "TC-99" in registered_test_cases()

    def test_duplicate_registration_raises(self) -> None:
        @register_gold("TC-98")
        def tc98a(canaries: Any, errors: Any, **kw: Any) -> GoldStandard:
            return GoldStandard(test_case="TC-98")

        import pytest
        with pytest.raises(ValueError, match="Duplicate"):
            @register_gold("TC-98")
            def tc98b(canaries: Any, errors: Any, **kw: Any) -> GoldStandard:
                return GoldStandard(test_case="TC-98")


# ---------------------------------------------------------------------------
# Emission
# ---------------------------------------------------------------------------

class TestEmission:

    def setup_method(self) -> None:
        self._saved = dict(_REGISTRY)
        _REGISTRY.clear()

    def teardown_method(self) -> None:
        _REGISTRY.clear()
        _REGISTRY.update(self._saved)

    def test_emit_gold_writes_json(self) -> None:
        @register_gold("TC-50")
        def tc50(canaries: CanaryRegistry, errors: ErrorRegistry, **kw: Any) -> GoldStandard:
            return GoldStandard(
                test_case="TC-50",
                expected_outputs={"value": 123},
                canary_verification={"file_a": canaries.canary_for("cascade_tb_fy2025")},
            )

        canaries = _make_canary_registry()
        errors = _make_error_registry()

        with tempfile.TemporaryDirectory() as tmpdir:
            out = Path(tmpdir) / "gold_standards"
            gold = emit_gold("TC-50", canaries, errors, out)
            assert gold.test_case == "TC-50"

            path = out / "TC-50_gold.json"
            assert path.exists()

            with open(path) as f:
                data = json.load(f)
            assert data["test_case"] == "TC-50"
            assert data["expected_outputs"]["value"] == 123
            assert data["canary_verification"]["file_a"] == "XK7P2M9Q"

    def test_emit_gold_validates_test_case_id(self) -> None:
        @register_gold("TC-51")
        def tc51(canaries: Any, errors: Any, **kw: Any) -> GoldStandard:
            return GoldStandard(test_case="TC-WRONG")

        import pytest
        with tempfile.TemporaryDirectory() as tmpdir:
            with pytest.raises(ValueError, match="TC-WRONG"):
                emit_gold("TC-51", CanaryRegistry(), ErrorRegistry(), Path(tmpdir))

    def test_emit_all_golds(self) -> None:
        for i in range(3):
            tc_id = f"TC-6{i}"

            # Use default args to capture tc_id properly
            @register_gold(tc_id)
            def make_gold(canaries: Any, errors: Any, _id: str = tc_id, **kw: Any) -> GoldStandard:
                return GoldStandard(test_case=_id, expected_outputs={"idx": int(_id[-1])})

        with tempfile.TemporaryDirectory() as tmpdir:
            out = Path(tmpdir) / "golds"
            golds = emit_all_golds(CanaryRegistry(), ErrorRegistry(), out)
            assert len(golds) == 3
            assert [g.test_case for g in golds] == ["TC-60", "TC-61", "TC-62"]

            for i in range(3):
                assert (out / f"TC-6{i}_gold.json").exists()

    def test_model_kwargs_forwarded(self) -> None:
        """Gold functions receive extra keyword arguments from the model."""
        received: dict[str, Any] = {}

        @register_gold("TC-70")
        def tc70(canaries: Any, errors: Any, **kw: Any) -> GoldStandard:
            received.update(kw)
            return GoldStandard(test_case="TC-70")

        with tempfile.TemporaryDirectory() as tmpdir:
            emit_gold("TC-70", CanaryRegistry(), ErrorRegistry(), Path(tmpdir),
                       ledger="fake_ledger", config="fake_config")

        assert received["ledger"] == "fake_ledger"
        assert received["config"] == "fake_config"


# ---------------------------------------------------------------------------
# Round-trip self-test
# ---------------------------------------------------------------------------

class TestRoundTrip:

    def test_simple_round_trip(self) -> None:
        gold = GoldStandard(
            test_case="TC-01",
            expected_outputs={
                "mapping": {"total_accounts_mapped": 118, "new_accounts_flagged": 3},
                "variance_analysis": {"flagged": ["4110", "5220"]},
            },
            canary_verification={"read_tb": "XK7P2M9Q"},
            error_detection={"ERR-001": "transposed digits in AR"},
        )
        with tempfile.TemporaryDirectory() as tmpdir:
            assert verify_round_trip(gold, Path(tmpdir)) is True

    def test_round_trip_with_scoring_hints(self) -> None:
        gold = GoldStandard(
            test_case="TC-05",
            scoring_hints={
                "correctness": "qualitative — no exact numbers",
                "completeness": "must use template sections",
            },
        )
        with tempfile.TemporaryDirectory() as tmpdir:
            assert verify_round_trip(gold, Path(tmpdir)) is True

    def test_round_trip_complex_nested(self) -> None:
        gold = GoldStandard(
            test_case="TC-06",
            expected_outputs={
                "provision": {
                    "current_federal": 1234567,
                    "current_state": 234567,
                    "deferred": -45000,
                    "effective_rate": 24.8,
                },
                "required_sheets": ["Current Provision", "Deferred Rollforward",
                                    "Rate Reconciliation", "Summary"],
            },
            canary_verification={"tb": "AB12CD34", "prior_wp": "EF56GH78"},
            error_detection={},
            scoring_hints={"correctness": "±0.5% tolerance on computed values"},
        )
        with tempfile.TemporaryDirectory() as tmpdir:
            assert verify_round_trip(gold, Path(tmpdir)) is True

    def test_read_gold_json(self) -> None:
        gold = GoldStandard(
            test_case="TC-03",
            expected_outputs={"growth_rate": 9.2},
        )
        with tempfile.TemporaryDirectory() as tmpdir:
            out = Path(tmpdir)
            from generator.golds.framework import _write_gold_json
            path = _write_gold_json(gold, out)
            reloaded = read_gold_json(path)
            assert reloaded.test_case == "TC-03"
            assert reloaded.expected_outputs["growth_rate"] == 9.2


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

class TestSortRecursive:

    def test_nested_dicts_sorted(self) -> None:
        result = _sort_recursive({"z": {"b": 1, "a": 2}, "a": 3})
        assert list(result.keys()) == ["a", "z"]
        assert list(result["z"].keys()) == ["a", "b"]

    def test_lists_preserved(self) -> None:
        result = _sort_recursive({"items": [3, 1, 2]})
        assert result["items"] == [3, 1, 2]

    def test_scalar_passthrough(self) -> None:
        assert _sort_recursive(42) == 42
        assert _sort_recursive("hello") == "hello"
