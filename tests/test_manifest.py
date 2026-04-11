"""Tests for generator/manifest.py — manifest emitter."""

from __future__ import annotations

import json
from pathlib import Path

from generator.manifest import Manifest


def test_register_and_to_dict() -> None:
    """Registered entries appear sorted by path in to_dict()."""
    manifest = Manifest(Path("/tmp/fake"))
    manifest.register("z_file.csv", "csv", canary="AAAA1111")
    manifest.register("a_file.xlsx", "xlsx", canary="BBBB2222", test_cases=["TC-03", "TC-01"])

    result = manifest.to_dict()
    assert len(result) == 2
    # Sorted by path
    assert result[0]["path"] == "a_file.xlsx"
    assert result[1]["path"] == "z_file.csv"
    # Test cases are sorted
    assert result[0]["test_cases"] == ["TC-01", "TC-03"]


def test_context_manager_writes_manifest(tmp_path: Path) -> None:
    """Exiting the context manager writes manifest.json with correct sizes."""
    # Create a dummy file so size resolution works
    (tmp_path / "shared_data").mkdir()
    dummy = tmp_path / "shared_data" / "coa.xlsx"
    dummy.write_bytes(b"fake excel content here")

    with Manifest(tmp_path) as manifest:
        manifest.register("shared_data/coa.xlsx", "xlsx", canary="XYZW1234", test_cases=["TC-01"])

    manifest_path = tmp_path / "manifest.json"
    assert manifest_path.exists()

    data = json.loads(manifest_path.read_text())
    assert len(data) == 1
    entry = data[0]
    assert entry["path"] == "shared_data/coa.xlsx"
    assert entry["type"] == "xlsx"
    assert entry["size"] == len(b"fake excel content here")
    assert entry["canary"] == "XYZW1234"
    assert entry["test_cases"] == ["TC-01"]


def test_context_manager_no_write_on_exception(tmp_path: Path) -> None:
    """If an exception occurs, manifest.json should NOT be written."""
    try:
        with Manifest(tmp_path) as manifest:
            manifest.register("test.csv", "csv")
            raise RuntimeError("simulated failure")
    except RuntimeError:
        pass

    assert not (tmp_path / "manifest.json").exists()


def test_size_zero_for_missing_file(tmp_path: Path) -> None:
    """Files that don't exist on disk get size=0."""
    with Manifest(tmp_path) as manifest:
        manifest.register("does_not_exist.csv", "csv")

    data = json.loads((tmp_path / "manifest.json").read_text())
    assert data[0]["size"] == 0


def test_orchestrator_produces_manifest(tmp_path: Path) -> None:
    """The orchestrator's generate() should produce a Manifest (currently empty)."""
    from generate_test_suite import generate
    from generator.config import load_config

    config = load_config("config.yaml")
    output = tmp_path / "suite"

    manifest = generate(config, output)
    assert isinstance(manifest, Manifest)
    # TC-01 formatter is wired — manifest should have entries
    entries = manifest.to_dict()
    assert len(entries) > 0
    # manifest.json should have been written
    assert (output / "manifest.json").exists()
    # Canary and error registries should be in the manifest
    paths = [e["path"] for e in entries]
    assert "canary_registry.json" in paths
    assert "error_registry.json" in paths


def test_scenario_pack_omitted_when_empty() -> None:
    """Entries without a scenario_pack should not include the key in to_dict()."""
    manifest = Manifest(Path("/tmp/fake"))
    manifest.register("file.csv", "csv")

    result = manifest.to_dict()
    assert "scenario_pack" not in result[0]


def test_scenario_pack_present_when_set() -> None:
    """Entries with a scenario_pack include it in to_dict()."""
    manifest = Manifest(Path("/tmp/fake"))
    manifest.register("file.csv", "csv", scenario_pack="cascade_accounting_core")

    result = manifest.to_dict()
    assert result[0]["scenario_pack"] == "cascade_accounting_core"


def test_scenario_pack_from_current_pack_context() -> None:
    """set_current_pack() provides the default for subsequent register() calls."""
    manifest = Manifest(Path("/tmp/fake"))
    manifest.set_current_pack("cascade_accounting_core")
    manifest.register("a.csv", "csv")
    manifest.set_current_pack("ma_legal_hr_diligence")
    manifest.register("b.csv", "csv")
    manifest.set_current_pack("")
    manifest.register("c.csv", "csv")

    result = manifest.to_dict()
    # Sorted by path: a.csv, b.csv, c.csv
    assert result[0]["scenario_pack"] == "cascade_accounting_core"
    assert result[1]["scenario_pack"] == "ma_legal_hr_diligence"
    assert "scenario_pack" not in result[2]


def test_scenario_pack_explicit_overrides_context() -> None:
    """An explicit scenario_pack kwarg takes precedence over set_current_pack()."""
    manifest = Manifest(Path("/tmp/fake"))
    manifest.set_current_pack("cascade_accounting_core")
    manifest.register("file.csv", "csv", scenario_pack="ma_legal_hr_diligence")

    result = manifest.to_dict()
    assert result[0]["scenario_pack"] == "ma_legal_hr_diligence"


def test_scenario_pack_deterministic_in_json(tmp_path: Path) -> None:
    """scenario_pack metadata is stable across reruns."""
    for run_dir in ("run1", "run2"):
        d = tmp_path / run_dir
        d.mkdir()
        (d / "a.csv").write_text("data")

        with Manifest(d) as m:
            m.set_current_pack("cascade_accounting_core")
            m.register("a.csv", "csv", canary="AAAA1111")

    content1 = (tmp_path / "run1" / "manifest.json").read_text()
    content2 = (tmp_path / "run2" / "manifest.json").read_text()
    assert content1 == content2
    # Verify the key is present
    data = json.loads(content1)
    assert data[0]["scenario_pack"] == "cascade_accounting_core"


def test_orchestrator_tags_entries_with_pack(tmp_path: Path) -> None:
    """The orchestrator's generate() should tag manifest entries with scenario_pack."""
    from generate_test_suite import generate
    from generator.config import load_config

    config = load_config("config.yaml")
    output = tmp_path / "suite"

    manifest = generate(config, output)
    entries = manifest.to_dict()
    # Every entry produced by a pack emitter should have scenario_pack
    pack_entries = [e for e in entries if e.get("scenario_pack")]
    assert len(pack_entries) > 0
    # Default generation only runs accounting-core
    packs_seen = {e["scenario_pack"] for e in pack_entries}
    assert "cascade_accounting_core" in packs_seen
    # Top-level registries (canary, error, scoring) are registered after
    # set_current_pack(""), so they should NOT have scenario_pack
    for path in ("canary_registry.json", "error_registry.json", "scoring/scoring_template.xlsx"):
        entry = next((e for e in entries if e["path"] == path), None)
        if entry:
            assert "scenario_pack" not in entry, f"{path} should not have scenario_pack"


def test_manifest_json_deterministic(tmp_path: Path) -> None:
    """Two identical registration sequences must produce byte-identical JSON."""
    for run_dir in ("run1", "run2"):
        d = tmp_path / run_dir
        d.mkdir()
        (d / "b.csv").write_text("data")
        (d / "a.xlsx").write_bytes(b"\x00" * 10)

        with Manifest(d) as m:
            m.register("b.csv", "csv", canary="CCCC3333")
            m.register("a.xlsx", "xlsx", canary="DDDD4444", test_cases=["TC-02"])

    content1 = (tmp_path / "run1" / "manifest.json").read_text()
    content2 = (tmp_path / "run2" / "manifest.json").read_text()
    assert content1 == content2
