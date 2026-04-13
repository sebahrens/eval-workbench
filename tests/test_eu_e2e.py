"""End-to-end self-test and determinism gate for the European variants pack.

Verifies:
1. EU pack generates all expected TC-*-EU test case directories and gold standards
2. Self-test (canary verification) passes for all EU test cases
3. Determinism: two full EU-pack runs produce byte-identical outputs
4. Default generation (no --packs) is unaffected by the EU pack's existence

Ref: bead synth-data-eu.22
"""

from __future__ import annotations

import filecmp
import json
from pathlib import Path

import pytest

from generate_test_suite import generate
from generator.config import load_config
from scoring.auto_grader import _run_self_test

# All TC-*-EU IDs that the cascade_europe_ifrs pack produces
_EU_TC_IDS = [
    "TC-04-EU",
    "TC-06-EU",
    "TC-07-EU",
    "TC-08-EU",
    "TC-09-EU",
    "TC-10-EU",
    "TC-12-EU",
    "TC-16-EU",
    "TC-18-EU",
]

# Packs needed for EU generation (EU depends on accounting_core)
_EU_PACKS = ["cascade_accounting_core", "cascade_europe_ifrs"]


def _collect_files(root: Path) -> list[Path]:
    """Return sorted list of all files under *root*, relative to *root*."""
    return sorted(p.relative_to(root) for p in root.rglob("*") if p.is_file())


def _assert_trees_identical(dir_a: Path, dir_b: Path) -> None:
    """Assert two directory trees contain identical files with identical bytes."""
    files_a = _collect_files(dir_a)
    files_b = _collect_files(dir_b)

    assert files_a == files_b, (
        f"File listings differ.\n"
        f"  Only in run1: {set(files_a) - set(files_b)}\n"
        f"  Only in run2: {set(files_b) - set(files_a)}"
    )

    for rel in files_a:
        fa = dir_a / rel
        fb = dir_b / rel
        assert filecmp.cmp(fa, fb, shallow=False), f"Content differs: {rel}"


# ---------------------------------------------------------------------------
# Fixture: generate the EU pack once per module
# ---------------------------------------------------------------------------

_EU_OUTPUT: Path | None = None
_EU_CONFIG = None


def _ensure_eu_generated(tmp_path_factory: pytest.TempPathFactory) -> Path:
    """Generate EU suite once, cache the output path."""
    global _EU_OUTPUT, _EU_CONFIG  # noqa: PLW0603
    if _EU_OUTPUT is None:
        _EU_CONFIG = load_config("config.yaml")
        _EU_OUTPUT = tmp_path_factory.mktemp("eu_suite")
        generate(_EU_CONFIG, _EU_OUTPUT, pack_ids=_EU_PACKS)
    return _EU_OUTPUT


@pytest.fixture(scope="module")
def eu_suite(tmp_path_factory: pytest.TempPathFactory) -> Path:
    return _ensure_eu_generated(tmp_path_factory)


# ---------------------------------------------------------------------------
# 1. File integrity: all EU test case directories and golds exist
# ---------------------------------------------------------------------------


class TestEUFileIntegrity:
    """Verify that the EU pack produces the expected directory structure."""

    @pytest.mark.parametrize("tc_id", _EU_TC_IDS)
    def test_test_case_dir_exists(self, eu_suite: Path, tc_id: str) -> None:
        tc_dir = eu_suite / "test_cases" / tc_id
        assert tc_dir.is_dir(), f"Missing test case directory: {tc_dir}"

    @pytest.mark.parametrize("tc_id", _EU_TC_IDS)
    def test_input_files_dir_exists(self, eu_suite: Path, tc_id: str) -> None:
        input_dir = eu_suite / "test_cases" / tc_id / "input_files"
        assert input_dir.is_dir(), f"Missing input_files: {input_dir}"

    @pytest.mark.parametrize("tc_id", _EU_TC_IDS)
    def test_input_files_not_empty(self, eu_suite: Path, tc_id: str) -> None:
        input_dir = eu_suite / "test_cases" / tc_id / "input_files"
        files = list(input_dir.iterdir())
        assert len(files) > 0, f"No input files in {input_dir}"

    @pytest.mark.parametrize("tc_id", _EU_TC_IDS)
    def test_gold_standard_exists(self, eu_suite: Path, tc_id: str) -> None:
        gold_path = eu_suite / "gold_standards" / f"{tc_id}_gold.json"
        assert gold_path.exists(), f"Missing gold standard: {gold_path}"

    @pytest.mark.parametrize("tc_id", _EU_TC_IDS)
    def test_gold_standard_valid_json(self, eu_suite: Path, tc_id: str) -> None:
        gold_path = eu_suite / "gold_standards" / f"{tc_id}_gold.json"
        data = json.loads(gold_path.read_text())
        assert "expected_outputs" in data, f"Gold {tc_id} missing expected_outputs"

    @pytest.mark.parametrize("tc_id", _EU_TC_IDS)
    def test_prompt_exists(self, eu_suite: Path, tc_id: str) -> None:
        prompt_path = eu_suite / "test_cases" / tc_id / "prompt.md"
        assert prompt_path.exists(), f"Missing prompt: {prompt_path}"

    @pytest.mark.parametrize("tc_id", _EU_TC_IDS)
    def test_expected_behavior_exists(self, eu_suite: Path, tc_id: str) -> None:
        eb_path = eu_suite / "test_cases" / tc_id / "expected_behavior.md"
        assert eb_path.exists(), f"Missing expected_behavior: {eb_path}"


# ---------------------------------------------------------------------------
# 2. Canary integrity: all EU canaries are present in the registry
# ---------------------------------------------------------------------------


class TestEUCanaryIntegrity:
    """Verify canary registry includes entries for all EU files."""

    def test_canary_registry_exists(self, eu_suite: Path) -> None:
        path = eu_suite / "canary_registry.json"
        assert path.exists()

    def test_canary_registry_has_eu_keys(self, eu_suite: Path) -> None:
        registry = json.loads((eu_suite / "canary_registry.json").read_text())
        eu_keys = [
            entry["file_key"]
            for entry in registry
            if "eu" in entry.get("file_key", "")
        ]
        assert len(eu_keys) > 0, "No EU canary keys in registry"

    @pytest.mark.parametrize("tc_id", _EU_TC_IDS)
    def test_canary_keys_registered_for_tc(self, eu_suite: Path, tc_id: str) -> None:
        """Each EU TC should have at least one canary key in the registry."""
        registry = json.loads((eu_suite / "canary_registry.json").read_text())
        tc_prefix = tc_id.lower().replace("-", "")
        tc_canaries = [
            entry for entry in registry
            if entry.get("file_key", "").startswith(tc_prefix)
        ]
        assert len(tc_canaries) > 0, (
            f"No canary keys with prefix {tc_prefix} in registry"
        )


# ---------------------------------------------------------------------------
# 3. Self-test: grader self-test passes for EU test cases
# ---------------------------------------------------------------------------


class TestEUSelfTest:
    """Verify the grader self-test passes when run against the EU suite."""

    def test_self_test_passes(self, eu_suite: Path) -> None:
        """_run_self_test must return True for the full suite including EU TCs."""
        ok = _run_self_test(eu_suite)
        assert ok, "Grader self-test failed on EU suite"

    @pytest.mark.parametrize("tc_id", _EU_TC_IDS)
    def test_eu_gold_in_self_test_scope(self, eu_suite: Path, tc_id: str) -> None:
        """Each EU gold file must be discovered by the self-test gold glob."""
        gold_dir = eu_suite / "gold_standards"
        gold_files = sorted(gold_dir.glob("TC-*_gold.json"))
        gold_ids = {gf.stem.replace("_gold", "") for gf in gold_files}
        assert tc_id in gold_ids, f"{tc_id} not found in self-test scope"


# ---------------------------------------------------------------------------
# 4. Determinism: two EU-pack runs produce identical outputs
# ---------------------------------------------------------------------------


class TestEUDeterminism:
    """Two generate() calls with the EU pack must produce byte-identical trees."""

    def test_determinism(self, tmp_path: Path) -> None:
        config = load_config("config.yaml")

        run1 = tmp_path / "run1"
        run2 = tmp_path / "run2"

        generate(config, run1, pack_ids=_EU_PACKS)
        generate(config, run2, pack_ids=_EU_PACKS)

        _assert_trees_identical(run1, run2)


# ---------------------------------------------------------------------------
# 5. Default generation unchanged: no EU TCs when packs=None
# ---------------------------------------------------------------------------


class TestDefaultUnchanged:
    """Default generation (packs=None) must NOT include any EU test cases."""

    def test_no_eu_dirs_in_default(self, tmp_path: Path) -> None:
        config = load_config("config.yaml")
        output = tmp_path / "default"
        generate(config, output)

        tc_dir = output / "test_cases"
        if tc_dir.exists():
            eu_dirs = [d.name for d in tc_dir.iterdir() if d.name.endswith("-EU")]
            assert eu_dirs == [], f"EU test cases in default generation: {eu_dirs}"

    def test_no_eu_golds_in_default(self, tmp_path: Path) -> None:
        config = load_config("config.yaml")
        output = tmp_path / "default"
        generate(config, output)

        gold_dir = output / "gold_standards"
        if gold_dir.exists():
            eu_golds = [g.name for g in gold_dir.glob("TC-*-EU_gold.json")]
            assert eu_golds == [], f"EU golds in default generation: {eu_golds}"

    def test_no_eu_canaries_in_default(self, tmp_path: Path) -> None:
        config = load_config("config.yaml")
        output = tmp_path / "default"
        generate(config, output)

        registry_path = output / "canary_registry.json"
        if registry_path.exists():
            registry = json.loads(registry_path.read_text())
            eu_keys = [
                entry["file_key"]
                for entry in registry
                if "eu" in entry.get("file_key", "")
            ]
            assert eu_keys == [], f"EU canary keys in default: {eu_keys}"
