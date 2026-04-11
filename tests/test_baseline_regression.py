"""Baseline regression: generated suite must match committed SHA-256 checksums.

Unlike test_determinism.py (which checks run-to-run equality), this test catches
silent drift where both runs agree with each other but differ from the last known-good
baseline.  This is the failure mode where Config/ScenarioContext plumbing changes shift
outputs consistently — both runs match, but the suite has drifted.

The baseline file ``tests/baseline_checksums.json`` is a committed JSON mapping of
``{relative_path: sha256_hex}``.  When the generator legitimately changes output
(new test case, new canary, schema change), regenerate the baseline::

    uv run python -c "
    from tests.test_baseline_regression import regenerate_baseline
    regenerate_baseline()
    "

Ref: bead synth-data-ln9.9.
"""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

import pytest

from generate_test_suite import generate
from generator.config import load_config

BASELINE_PATH = Path(__file__).parent / "baseline_checksums.json"


def _checksum_tree(root: Path) -> dict[str, str]:
    """Return ``{relative_path: sha256_hex}`` for every file under *root*."""
    result: dict[str, str] = {}
    for p in sorted(root.rglob("*")):
        if p.is_file():
            result[str(p.relative_to(root))] = hashlib.sha256(
                p.read_bytes(),
            ).hexdigest()
    return result


def _load_baseline() -> dict[str, str]:
    if not BASELINE_PATH.exists():
        pytest.fail(
            f"Baseline file missing: {BASELINE_PATH}\n"
            "Regenerate with: uv run python -c "
            '"from tests.test_baseline_regression import regenerate_baseline; '
            'regenerate_baseline()"',
        )
    return json.loads(BASELINE_PATH.read_text())


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------


def test_suite_matches_baseline(tmp_path: Path) -> None:
    """Generated suite must be byte-identical to the committed baseline."""
    config = load_config("config.yaml")
    out = tmp_path / "suite"
    generate(config, out)

    actual = _checksum_tree(out)
    expected = _load_baseline()

    # Check file listing first — gives a clear message on added/removed files.
    actual_keys = set(actual)
    expected_keys = set(expected)

    missing = expected_keys - actual_keys
    extra = actual_keys - expected_keys

    if missing or extra:
        parts: list[str] = ["File listing differs from baseline."]
        if missing:
            parts.append(f"  Missing ({len(missing)}): {sorted(missing)[:10]}")
        if extra:
            parts.append(f"  Extra   ({len(extra)}): {sorted(extra)[:10]}")
        pytest.fail("\n".join(parts))

    # Check content — report the first drifted path for fast debugging.
    drifted: list[str] = [
        rel for rel in sorted(expected) if actual[rel] != expected[rel]
    ]
    if drifted:
        first = drifted[0]
        pytest.fail(
            f"Baseline drift detected in {len(drifted)} file(s). "
            f"First drifted: {first}\n"
            f"  expected: {expected[first]}\n"
            f"  actual:   {actual[first]}\n"
            "If the change is intentional, regenerate the baseline:\n"
            '  uv run python -c "from tests.test_baseline_regression '
            'import regenerate_baseline; regenerate_baseline()"',
        )


# ---------------------------------------------------------------------------
# Baseline regeneration helper
# ---------------------------------------------------------------------------


def regenerate_baseline() -> None:
    """Regenerate ``tests/baseline_checksums.json`` from a fresh generation."""
    config = load_config("config.yaml")
    from tempfile import TemporaryDirectory

    with TemporaryDirectory() as td:
        out = Path(td) / "suite"
        generate(config, out)
        checksums = _checksum_tree(out)

    BASELINE_PATH.write_text(json.dumps(checksums, sort_keys=True, indent=2) + "\n")
    print(f"Wrote {len(checksums)} checksums to {BASELINE_PATH}")
