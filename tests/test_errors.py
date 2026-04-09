"""Tests for generator.errors — planted error framework."""

from __future__ import annotations

import json
from pathlib import Path

import pytest

from generator.errors import (
    ErrorRegistry,
    PlantedError,
    classification_error,
    date_inconsistency,
    formula_error,
    mismatch_total,
    missing_data,
    rounding_discrepancy,
    stale_data,
    transpose_digits,
    wrong_entity,
)

# ---------------------------------------------------------------------------
# PlantedError dataclass
# ---------------------------------------------------------------------------

class TestPlantedError:
    def test_fields(self):
        e = PlantedError(
            error_id="ERR-001",
            file="cascade_tb_fy2025.xlsx",
            location="Sheet 'Trial Balance', Cell G47",
            type="transposed_digits",
            description="AR balance transposed",
            severity="material",
            which_test_cases_should_catch=["TC-01", "TC-02"],
        )
        assert e.error_id == "ERR-001"
        assert e.which_test_cases_should_catch == ["TC-01", "TC-02"]

    def test_default_test_cases_empty(self):
        e = PlantedError(
            error_id="ERR-099",
            file="test.xlsx",
            location="A1",
            type="stale_data",
            description="test",
            severity="immaterial",
        )
        assert e.which_test_cases_should_catch == []


# ---------------------------------------------------------------------------
# ErrorRegistry
# ---------------------------------------------------------------------------

def _make_error(error_id: str, etype: str = "stale_data", file: str = "a.xlsx",
                test_cases: list[str] | None = None) -> PlantedError:
    return PlantedError(
        error_id=error_id,
        file=file,
        location="A1",
        type=etype,
        description=f"Test error {error_id}",
        severity="material",
        which_test_cases_should_catch=test_cases or [],
    )


class TestErrorRegistry:
    def test_add_and_get(self):
        reg = ErrorRegistry()
        e = _make_error("ERR-001")
        reg.add(e)
        assert reg.get("ERR-001") is e

    def test_reject_invalid_type(self):
        reg = ErrorRegistry()
        e = _make_error("ERR-001", etype="bogus_type")
        with pytest.raises(ValueError, match="Invalid error type"):
            reg.add(e)

    def test_reject_duplicate_id(self):
        reg = ErrorRegistry()
        reg.add(_make_error("ERR-001"))
        with pytest.raises(ValueError, match="Duplicate error_id"):
            reg.add(_make_error("ERR-001", file="b.xlsx"))

    def test_by_file(self):
        reg = ErrorRegistry()
        reg.add(_make_error("ERR-001", file="a.xlsx"))
        reg.add(_make_error("ERR-002", file="b.xlsx"))
        reg.add(_make_error("ERR-003", file="a.xlsx"))
        result = reg.by_file("a.xlsx")
        assert [e.error_id for e in result] == ["ERR-001", "ERR-003"]

    def test_by_type(self):
        reg = ErrorRegistry()
        reg.add(_make_error("ERR-001", etype="stale_data"))
        reg.add(_make_error("ERR-002", etype="missing_data"))
        reg.add(_make_error("ERR-003", etype="stale_data"))
        result = reg.by_type("stale_data")
        assert [e.error_id for e in result] == ["ERR-001", "ERR-003"]

    def test_by_test_case(self):
        reg = ErrorRegistry()
        reg.add(_make_error("ERR-001", test_cases=["TC-01", "TC-02"]))
        reg.add(_make_error("ERR-002", test_cases=["TC-03"]))
        reg.add(_make_error("ERR-003", test_cases=["TC-01"]))
        result = reg.by_test_case("TC-01")
        assert [e.error_id for e in result] == ["ERR-001", "ERR-003"]


class TestRegistryJson:
    def test_write_and_read(self, tmp_path: Path):
        reg = ErrorRegistry()
        reg.add(_make_error("ERR-002"))
        reg.add(_make_error("ERR-001", file="b.xlsx"))

        out = tmp_path / "error_registry.json"
        reg.write_json(out)

        data = json.loads(out.read_text())
        assert isinstance(data, list)
        assert len(data) == 2
        # Sorted by error_id
        assert [d["error_id"] for d in data] == ["ERR-001", "ERR-002"]

    def test_deterministic_json(self, tmp_path: Path):
        """Two identical registries produce identical JSON."""
        def build() -> ErrorRegistry:
            reg = ErrorRegistry()
            reg.add(_make_error("ERR-003", file="c.xlsx"))
            reg.add(_make_error("ERR-001", file="a.xlsx"))
            reg.add(_make_error("ERR-002", file="b.xlsx"))
            return reg

        p1 = tmp_path / "a.json"
        p2 = tmp_path / "b.json"
        build().write_json(p1)
        build().write_json(p2)
        assert p1.read_text() == p2.read_text()


# ---------------------------------------------------------------------------
# Transformation functions
# ---------------------------------------------------------------------------

class TestTransposeDigits:
    def test_basic_int(self):
        # 18423109 → swap pos -3 and -2 → 18423190? Let's be explicit.
        # digits: 1 8 4 2 3 1 0 9
        # pos -3 = index 5 = '1', pos -2 = index 6 = '0'
        # swapped: 1 8 4 2 3 0 1 9 = 18423019
        result = transpose_digits(18423109, pos1=-3, pos2=-2)
        assert result == 18423019

    def test_basic_float(self):
        result = transpose_digits(12345.67, pos1=0, pos2=1)
        # digits of int part "12345": swap index 0,1 → "21345"
        assert result == 21345.67

    def test_negative(self):
        result = transpose_digits(-12345, pos1=0, pos2=1)
        assert result == -21345

    def test_returns_same_type(self):
        assert isinstance(transpose_digits(12345, pos1=0, pos2=1), int)
        assert isinstance(transpose_digits(12345.0, pos1=0, pos2=1), float)

    def test_same_position_raises(self):
        with pytest.raises(ValueError, match="same index"):
            transpose_digits(12345, pos1=2, pos2=2)

    def test_identical_digits_raises(self):
        # 11234 — positions 0 and 1 are both '1'
        with pytest.raises(ValueError, match="identical"):
            transpose_digits(11234, pos1=0, pos2=1)


class TestMismatchTotal:
    def test_positive_delta(self):
        assert mismatch_total(100_000, 5_000) == 105_000

    def test_negative_delta(self):
        assert mismatch_total(100_000, -5_000) == 95_000

    def test_float(self):
        result = mismatch_total(100.50, 0.25)
        assert result == pytest.approx(100.75)

    def test_preserves_type(self):
        assert isinstance(mismatch_total(100, 5), int)
        assert isinstance(mismatch_total(100.0, 5.0), float)


class TestStaleData:
    def test_returns_input(self):
        assert stale_data(42) == 42
        assert stale_data("old value") == "old value"


class TestFormulaError:
    def test_returns_wrong_value(self):
        assert formula_error(correct_value=1000, wrong_value=900) == 900


class TestWrongEntity:
    def test_returns_wrong_name(self):
        assert wrong_entity("Cascade Precision", "Cascade Advanced") == "Cascade Advanced"


class TestDateInconsistency:
    def test_returns_wrong_date(self):
        assert date_inconsistency("2025-01-01", "2024-01-01") == "2024-01-01"


class TestClassificationError:
    def test_returns_wrong_account(self):
        assert classification_error("6100-Travel", "6200-Meals") == "6200-Meals"


class TestMissingData:
    def test_default_none(self):
        assert missing_data() is None

    def test_custom_placeholder(self):
        assert missing_data("") == ""
        assert missing_data(0) == 0


class TestRoundingDiscrepancy:
    def test_round_up(self):
        # 1234.5 rounded to 0 dp with ROUND_HALF_UP → 1235
        result = rounding_discrepancy(1234.5, decimal_places=0, direction="up")
        assert result == 1235.0

    def test_truncate_down(self):
        # 1234.9 truncated to 0 dp → 1234
        result = rounding_discrepancy(1234.9, decimal_places=0, direction="down")
        assert result == 1234.0

    def test_no_discrepancy_raises(self):
        # 1234.0 rounded to 0 dp is still 1234.0
        with pytest.raises(ValueError, match="same value"):
            rounding_discrepancy(1234.0, decimal_places=0, direction="up")

    def test_invalid_direction_raises(self):
        with pytest.raises(ValueError, match="direction must be"):
            rounding_discrepancy(1234.5, direction="sideways")

    def test_negative_truncate(self):
        # -1234.9 truncated toward zero → -1234
        result = rounding_discrepancy(-1234.9, decimal_places=0, direction="down")
        assert result == -1234.0

    def test_decimal_places(self):
        # 100.456 rounded to 2dp up → 100.46
        result = rounding_discrepancy(100.456, decimal_places=2, direction="up")
        assert result == pytest.approx(100.46)
