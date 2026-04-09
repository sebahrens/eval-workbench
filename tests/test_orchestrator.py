"""Tests for generate_test_suite.py orchestrator skeleton."""

from __future__ import annotations

from pathlib import Path

from generate_test_suite import _TEST_CASE_COUNT, generate
from generator.config import load_config


def test_generate_creates_directory_tree(tmp_path: Path) -> None:
    """generate() must create the full directory tree per prompt.md §2.3."""
    config = load_config("config.yaml")
    output = tmp_path / "suite"

    generate(config, output)

    # Top-level dirs
    assert (output / "shared_data").is_dir()
    assert (output / "gold_standards").is_dir()
    assert (output / "scoring").is_dir()
    assert (output / "templates").is_dir()
    assert (output / "test_cases").is_dir()

    # Per-test-case dirs
    for i in range(1, _TEST_CASE_COUNT + 1):
        tc = output / "test_cases" / f"TC-{i:02d}"
        assert tc.is_dir(), f"Missing {tc}"
        assert (tc / "input_files").is_dir(), f"Missing {tc}/input_files"


def test_generate_is_idempotent(tmp_path: Path) -> None:
    """Running generate() twice on the same output must not fail."""
    config = load_config("config.yaml")
    output = tmp_path / "suite"

    generate(config, output)
    generate(config, output)  # should not raise
