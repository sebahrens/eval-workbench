"""Planted error framework (section 1.7 of prompt.md).

Provides a PlantedError dataclass, transformation functions for each error
type defined in the spec, and an ErrorRegistry that serialises to
error_registry.json.

Error types (25 total):
- transposed_digits (3)
- mismatched_total (4)
- stale_data (3)
- formula_error (2)
- wrong_entity (2)
- date_inconsistency (3)
- classification_error (3)
- missing_data (3)
- rounding_discrepancy (2)
"""

from __future__ import annotations

import json
import math
from dataclasses import asdict, dataclass, field
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path
from typing import Any

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class PlantedError:
    """One entry in the error registry."""

    error_id: str                     # e.g. "ERR-001"
    file: str                         # File where the error lives
    location: str                     # e.g. "Sheet 'Trial Balance', Cell G47"
    type: str                         # One of the _VALID_TYPES
    description: str                  # Human-readable description
    severity: str                     # "material" | "immaterial" | "significant"
    which_test_cases_should_catch: list[str] = field(default_factory=list)


_VALID_TYPES = frozenset({
    "transposed_digits",
    "mismatched_total",
    "stale_data",
    "formula_error",
    "wrong_entity",
    "date_inconsistency",
    "classification_error",
    "missing_data",
    "rounding_discrepancy",
})


# ---------------------------------------------------------------------------
# Registry
# ---------------------------------------------------------------------------

@dataclass
class ErrorRegistry:
    """Complete registry of planted errors for the test suite."""

    entries: dict[str, PlantedError] = field(default_factory=dict)

    def add(self, error: PlantedError) -> None:
        if error.type not in _VALID_TYPES:
            raise ValueError(
                f"Invalid error type '{error.type}'. "
                f"Must be one of: {sorted(_VALID_TYPES)}"
            )
        if error.error_id in self.entries:
            raise ValueError(f"Duplicate error_id: {error.error_id}")
        self.entries[error.error_id] = error

    def get(self, error_id: str) -> PlantedError:
        return self.entries[error_id]

    def by_file(self, file: str) -> list[PlantedError]:
        """Return all errors planted in *file*, sorted by error_id."""
        return sorted(
            (e for e in self.entries.values() if e.file == file),
            key=lambda e: e.error_id,
        )

    def by_type(self, error_type: str) -> list[PlantedError]:
        """Return all errors of *error_type*, sorted by error_id."""
        return sorted(
            (e for e in self.entries.values() if e.type == error_type),
            key=lambda e: e.error_id,
        )

    def by_test_case(self, tc: str) -> list[PlantedError]:
        """Return all errors that *tc* should catch, sorted by error_id."""
        return sorted(
            (e for e in self.entries.values() if tc in e.which_test_cases_should_catch),
            key=lambda e: e.error_id,
        )

    # -- serialisation --------------------------------------------------------

    def to_dict(self) -> list[dict]:
        """Return a sorted list of entry dicts (deterministic order)."""
        return [asdict(self.entries[k]) for k in sorted(self.entries)]

    def write_json(self, path: str | Path) -> None:
        """Write the registry to *path* as formatted JSON."""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w") as f:
            json.dump(self.to_dict(), f, indent=2, sort_keys=True)
            f.write("\n")


# ---------------------------------------------------------------------------
# Transformation functions
#
# Each takes a clean value and returns the corrupted value.
# The caller is responsible for placing the corrupted value in the right cell.
# ---------------------------------------------------------------------------

def transpose_digits(value: int | float, pos1: int = -3, pos2: int = -2) -> int | float:
    """Swap two digits in *value*.

    *pos1* and *pos2* are indices into the digit string (negative indices
    count from the right, like Python slicing).

    Returns the same numeric type as the input.
    """
    is_float = isinstance(value, float)
    if is_float:
        # Work on the integer part only; preserve decimals
        int_part = int(value)
        frac = value - int_part
    else:
        int_part = value
        frac = 0

    negative = int_part < 0
    digits = list(str(abs(int_part)))

    # Normalise negative indices
    n = len(digits)
    i = pos1 % n
    j = pos2 % n
    if i == j:
        raise ValueError(f"pos1 and pos2 resolve to the same index ({i})")
    if digits[i] == digits[j]:
        raise ValueError(
            f"Digits at positions {pos1},{pos2} are identical ('{digits[i]}') — "
            "transposition would be invisible"
        )
    digits[i], digits[j] = digits[j], digits[i]

    result = int("".join(digits))
    if negative:
        result = -result
    if is_float:
        return float(result) + frac
    return result


def mismatch_total(
    correct_total: int | float,
    delta: int | float,
) -> int | float:
    """Return a total that differs from *correct_total* by *delta*.

    The delta should be chosen to be plausible but detectable — e.g. omitting
    one line item, double-counting, or an off-by-one error.
    """
    return type(correct_total)(correct_total + delta)


def stale_data(prior_year_value: Any) -> Any:
    """Return *prior_year_value* unchanged — the error is that this value
    should have been updated to the current year but wasn't.

    This function exists for symmetry with other transformations and to make
    the intent explicit in calling code.
    """
    return prior_year_value


def formula_error(correct_value: int | float, wrong_value: int | float) -> int | float:
    """Return *wrong_value* — the result of an incorrect Excel formula.

    The caller specifies both the correct and wrong value.  The wrong value
    is typically computed from a plausible but incorrect formula (e.g.
    SUM of wrong range, missing a row, referencing the wrong column).
    """
    return wrong_value


def wrong_entity(correct_name: str, wrong_name: str) -> str:
    """Return *wrong_name* instead of *correct_name*.

    Typically a subsidiary or client name that was copy-pasted from another
    engagement or entity.
    """
    return wrong_name


def date_inconsistency(correct_date: str, wrong_date: str) -> str:
    """Return *wrong_date* instead of *correct_date*.

    Dates are strings because they may appear in various formats across
    different file types (xlsx cells, docx text, PDF text).
    """
    return wrong_date


def classification_error(correct_account: str, wrong_account: str) -> str:
    """Return *wrong_account* — an expense or item booked to the wrong account.

    The caller provides both the correct and incorrect account codes/names.
    """
    return wrong_account


def missing_data(placeholder: Any = None) -> Any:
    """Return *placeholder* (default ``None``) — representing a blank cell
    or missing field that should have a value.

    The caller is responsible for writing this as an empty cell / blank field.
    """
    return placeholder


def rounding_discrepancy(
    correct_value: float,
    decimal_places: int = 0,
    direction: str = "up",
) -> float:
    """Return a value that differs from *correct_value* by a rounding error.

    Simulates the common case where one system rounds and another truncates,
    or where intermediate rounding causes a 1-cent / 1-dollar difference.

    Parameters
    ----------
    correct_value:
        The correct value.
    decimal_places:
        Number of decimal places to round to.
    direction:
        "up" rounds half-up (ROUND_HALF_UP), "down" truncates toward zero.
    """
    d = Decimal(str(correct_value))
    quantize_to = Decimal(10) ** -decimal_places

    if direction == "up":
        rounded = float(d.quantize(quantize_to, rounding=ROUND_HALF_UP))
    elif direction == "down":
        # Truncate toward zero (ROUND_DOWN always truncates toward zero)
        rounded = float(d.quantize(quantize_to, rounding="ROUND_DOWN"))
    else:
        raise ValueError(f"direction must be 'up' or 'down', got '{direction}'")

    if math.isclose(rounded, correct_value, rel_tol=0, abs_tol=1e-12):
        raise ValueError(
            f"Rounding {correct_value} to {decimal_places} decimal places "
            f"({direction}) produces the same value — no discrepancy"
        )

    return rounded
