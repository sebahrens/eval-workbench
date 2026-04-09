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
