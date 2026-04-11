"""End-to-end tests for the legal/HR diligence scenario pack (synth-data-ups.21).

Validates:
  1. TC-19, TC-20, TC-21 directories and gold standards are generated when
     the legal/HR pack is enabled via ``["all"]``.
  2. Self-test passes for the full suite (accounting-core + legal/HR).
  3. Determinism holds: two identical runs produce byte-identical output.
  4. Default output (no pack selection) is byte-identical to an accounting-core-
     only run — enabling the legal/HR pack does not alter the core output.
"""

from __future__ import annotations

import filecmp
import json
from pathlib import Path

import pytest

from generate_test_suite import generate
from generator.config import load_config
from generator.packs import REGISTRY, resolve_packs
from scoring.auto_grader import _run_self_test

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _collect_files(root: Path) -> list[Path]:
    """Return sorted list of all files under *root*, relative to *root*."""
    return sorted(p.relative_to(root) for p in root.rglob("*") if p.is_file())


def _assert_trees_identical(dir_a: Path, dir_b: Path) -> None:
    files_a = _collect_files(dir_a)
    files_b = _collect_files(dir_b)
    assert files_a == files_b, (
        f"File listings differ.\n"
        f"  Only in run1: {set(files_a) - set(files_b)}\n"
        f"  Only in run2: {set(files_b) - set(files_a)}"
    )
    for rel in files_a:
        assert filecmp.cmp(dir_a / rel, dir_b / rel, shallow=False), (
            f"Content differs: {rel}"
        )


@pytest.fixture()
def config():
    return load_config("config.yaml")


# ---------------------------------------------------------------------------
# Pack registration
# ---------------------------------------------------------------------------


class TestPackRegistration:
    """The legal/HR pack is registered and discoverable."""

    def test_legal_hr_pack_in_registry(self):
        assert "cascade_legal_hr_diligence" in REGISTRY

    def test_legal_hr_test_cases(self):
        pack = REGISTRY["cascade_legal_hr_diligence"]
        assert pack.test_cases == ["TC-19", "TC-20", "TC-21"]

    def test_legal_hr_depends_on_accounting_core(self):
        pack = REGISTRY["cascade_legal_hr_diligence"]
        assert "cascade_accounting_core" in pack.dependencies

    def test_resolve_all_includes_legal_hr(self):
        packs = resolve_packs(["all"])
        ids = [p.pack_id for p in packs]
        assert "cascade_legal_hr_diligence" in ids
        # Accounting core must come first (dependency ordering)
        assert ids.index("cascade_accounting_core") < ids.index(
            "cascade_legal_hr_diligence"
        )

    def test_legal_hr_alone_fails_without_dependency(self):
        """Selecting only the legal/HR pack without its dependency must fail."""
        from generator.config import ConfigError

        with pytest.raises(ConfigError, match="depends on"):
            resolve_packs(["cascade_legal_hr_diligence"])


# ---------------------------------------------------------------------------
# Generation — TC-19/20/21 presence
# ---------------------------------------------------------------------------


class TestLegalHRGeneration:
    """Full generation with all packs produces TC-19 through TC-21."""

    def test_tc19_tc20_tc21_directories_exist(self, tmp_path, config):
        out = tmp_path / "suite"
        generate(config, out, pack_ids=["all"])

        for tc in ("TC-19", "TC-20", "TC-21"):
            tc_dir = out / "test_cases" / tc
            assert tc_dir.is_dir(), f"{tc} directory not created"
            inputs = tc_dir / "input_files"
            assert inputs.is_dir(), f"{tc}/input_files not created"
            # At least one input file should exist
            input_files = list(inputs.rglob("*"))
            assert len([f for f in input_files if f.is_file()]) > 0, (
                f"{tc} has no input files"
            )

    def test_gold_standards_exist(self, tmp_path, config):
        out = tmp_path / "suite"
        generate(config, out, pack_ids=["all"])

        gold_dir = out / "gold_standards"
        for tc in ("TC-19", "TC-20", "TC-21"):
            gold_path = gold_dir / f"{tc}_gold.json"
            assert gold_path.exists(), f"Missing gold standard: {gold_path.name}"
            data = json.loads(gold_path.read_text())
            assert "canary_verification" in data
            assert "expected_outputs" in data

    def test_accounting_core_tcs_still_present(self, tmp_path, config):
        """TC-01 through TC-18 must still exist when all packs are enabled."""
        out = tmp_path / "suite"
        generate(config, out, pack_ids=["all"])

        for i in range(1, 19):
            tc_dir = out / "test_cases" / f"TC-{i:02d}"
            assert tc_dir.is_dir(), f"TC-{i:02d} missing when all packs enabled"


# ---------------------------------------------------------------------------
# Self-test
# ---------------------------------------------------------------------------


class TestSelfTest:
    """Self-test passes for the full suite including legal/HR TCs."""

    def test_self_test_all_packs(self, tmp_path, config):
        out = tmp_path / "suite"
        generate(config, out, pack_ids=["all"])

        ok = _run_self_test(out)
        assert ok is True, "Self-test must pass for all packs including legal/HR"


# ---------------------------------------------------------------------------
# Determinism
# ---------------------------------------------------------------------------


class TestDeterminism:
    """Two runs with all packs produce byte-identical output."""

    def test_all_packs_deterministic(self, tmp_path, config):
        run1 = tmp_path / "run1"
        run2 = tmp_path / "run2"

        generate(config, run1, pack_ids=["all"])
        generate(config, run2, pack_ids=["all"])

        _assert_trees_identical(run1, run2)


# ---------------------------------------------------------------------------
# Default output isolation
# ---------------------------------------------------------------------------


class TestDefaultOutputUnchanged:
    """Enabling the legal/HR pack must not alter default (accounting-core) output."""

    def test_default_output_byte_identical(self, tmp_path, config):
        default_run = tmp_path / "default"
        generate(config, default_run)  # pack_ids=None → default packs only

        core_only = tmp_path / "core_only"
        generate(config, core_only, pack_ids=["cascade_accounting_core"])

        _assert_trees_identical(default_run, core_only)
