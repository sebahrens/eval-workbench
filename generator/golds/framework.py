"""Gold standard emission framework (generic emitter + per-TC registration).

Each test case registers a thin gold function that receives the canonical
data model and returns a GoldStandard dataclass.  The framework serialises
every registered gold to ``gold_standards/TC-XX_gold.json`` and supports a
round-trip self-test (emit → re-read → equal).

Gold JSON schema (per prompt.md §7):

    {
      "test_case": "TC-XX",
      "expected_outputs": { ... },         # TC-specific numerical/structural expectations
      "canary_verification": { ... },      # {label: canary_code} the agent must read
      "error_detection": { ... },          # {error_id: description} the agent should catch
      "scoring_hints": { ... }             # Optional: per-dimension scoring guidance
    }
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

from generator.canaries import CanaryRegistry
from generator.errors import ErrorRegistry

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class GoldStandard:
    """One test case's gold standard definition.

    Attributes
    ----------
    test_case:
        Identifier such as ``"TC-01"``.
    expected_outputs:
        TC-specific dict of expected numerical values, required sheets/sections,
        structural requirements, etc.
    canary_verification:
        Mapping of human-readable labels (e.g. ``"read_correct_tb"``) to the
        8-char canary codes the agent should encounter when reading the right files.
    error_detection:
        Mapping of error IDs (e.g. ``"ERR-001"``) to a short description of
        what the agent should flag.
    scoring_hints:
        Optional per-dimension guidance for the auto-grader / human grader.
    """

    test_case: str
    expected_outputs: dict[str, Any] = field(default_factory=dict)
    canary_verification: dict[str, str] = field(default_factory=dict)
    error_detection: dict[str, str] = field(default_factory=dict)
    scoring_hints: dict[str, Any] = field(default_factory=dict)

    # -- serialisation --------------------------------------------------------

    def to_dict(self) -> dict[str, Any]:
        """Return a JSON-serialisable dict with deterministic key order."""
        d: dict[str, Any] = {
            "test_case": self.test_case,
            "expected_outputs": _sort_recursive(self.expected_outputs),
            "canary_verification": dict(sorted(self.canary_verification.items())),
            "error_detection": dict(sorted(self.error_detection.items())),
        }
        if self.scoring_hints:
            d["scoring_hints"] = _sort_recursive(self.scoring_hints)
        return d

    # -- deserialisation ------------------------------------------------------

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> GoldStandard:
        """Reconstruct a GoldStandard from a dict (e.g. loaded from JSON)."""
        return cls(
            test_case=data["test_case"],
            expected_outputs=data.get("expected_outputs", {}),
            canary_verification=data.get("canary_verification", {}),
            error_detection=data.get("error_detection", {}),
            scoring_hints=data.get("scoring_hints", {}),
        )


# ---------------------------------------------------------------------------
# Registry of per-TC gold functions
# ---------------------------------------------------------------------------

# Type alias: a gold function takes (canary_registry, error_registry, **model_kwargs)
# and returns a GoldStandard.  The model_kwargs will carry whatever data the
# TC needs from the canonical model (ledger, employees, config, etc.).
GoldFunc = Callable[..., GoldStandard]

_REGISTRY: dict[str, GoldFunc] = {}


def register_gold(test_case_id: str) -> Callable[[GoldFunc], GoldFunc]:
    """Decorator to register a gold function for *test_case_id*.

    Usage::

        @register_gold("TC-01")
        def tc01_gold(
            canaries: CanaryRegistry,
            errors: ErrorRegistry,
            **model: Any,
        ) -> GoldStandard:
            ...
    """
    def decorator(fn: GoldFunc) -> GoldFunc:
        if test_case_id in _REGISTRY:
            raise ValueError(
                f"Duplicate gold registration for {test_case_id}"
            )
        _REGISTRY[test_case_id] = fn
        return fn
    return decorator


def registered_test_cases() -> list[str]:
    """Return sorted list of test case IDs that have gold functions."""
    return sorted(_REGISTRY)


def get_gold_func(test_case_id: str) -> GoldFunc:
    """Return the gold function for *test_case_id*, raising KeyError if absent."""
    return _REGISTRY[test_case_id]


# ---------------------------------------------------------------------------
# Emission
# ---------------------------------------------------------------------------

def emit_gold(
    test_case_id: str,
    canaries: CanaryRegistry,
    errors: ErrorRegistry,
    output_dir: Path,
    **model_kwargs: Any,
) -> GoldStandard:
    """Run the gold function for *test_case_id* and write ``TC-XX_gold.json``.

    Returns the GoldStandard so callers can inspect or aggregate.
    """
    fn = get_gold_func(test_case_id)
    gold = fn(canaries, errors, **model_kwargs)

    # Validate test_case matches
    if gold.test_case != test_case_id:
        raise ValueError(
            f"Gold function for {test_case_id} returned test_case="
            f"'{gold.test_case}'"
        )

    _write_gold_json(gold, output_dir)
    return gold


def emit_all_golds(
    canaries: CanaryRegistry,
    errors: ErrorRegistry,
    output_dir: Path,
    **model_kwargs: Any,
) -> list[GoldStandard]:
    """Emit gold JSONs for every registered test case.

    Returns the list of GoldStandard objects (sorted by test_case).
    """
    golds: list[GoldStandard] = []
    for tc_id in registered_test_cases():
        gold = emit_gold(tc_id, canaries, errors, output_dir, **model_kwargs)
        golds.append(gold)
    return golds


# ---------------------------------------------------------------------------
# I/O helpers
# ---------------------------------------------------------------------------

def _write_gold_json(gold: GoldStandard, output_dir: Path) -> Path:
    """Write a single gold standard JSON and return the path."""
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / f"{gold.test_case}_gold.json"
    with open(path, "w") as f:
        json.dump(gold.to_dict(), f, indent=2, sort_keys=False)
        f.write("\n")
    return path


def read_gold_json(path: Path) -> GoldStandard:
    """Read a gold standard JSON back into a GoldStandard object."""
    with open(path) as f:
        data = json.load(f)
    return GoldStandard.from_dict(data)


# ---------------------------------------------------------------------------
# Round-trip self-test
# ---------------------------------------------------------------------------

def verify_round_trip(gold: GoldStandard, output_dir: Path) -> bool:
    """Emit *gold* to JSON, read it back, and verify equality.

    Returns True if the round-trip produces an identical GoldStandard.
    Raises AssertionError with details on mismatch.
    """
    path = _write_gold_json(gold, output_dir)
    reloaded = read_gold_json(path)

    original_dict = gold.to_dict()
    reloaded_dict = reloaded.to_dict()

    if original_dict != reloaded_dict:
        # Find the diff for diagnostics
        import difflib
        orig_lines = json.dumps(original_dict, indent=2).splitlines(keepends=True)
        reload_lines = json.dumps(reloaded_dict, indent=2).splitlines(keepends=True)
        diff = "".join(difflib.unified_diff(orig_lines, reload_lines,
                                             fromfile="original", tofile="reloaded"))
        raise AssertionError(f"Round-trip mismatch for {gold.test_case}:\n{diff}")

    return True


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _sort_recursive(obj: Any) -> Any:
    """Recursively sort dict keys for deterministic JSON output."""
    if isinstance(obj, dict):
        return {k: _sort_recursive(v) for k, v in sorted(obj.items())}
    if isinstance(obj, list):
        return [_sort_recursive(item) for item in obj]
    return obj
