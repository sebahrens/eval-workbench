"""Tests for generate_test_suite.py orchestrator skeleton."""

from __future__ import annotations

from pathlib import Path

import pytest

from generate_test_suite import generate, main
from generator.config import load_config
from generator.packs import collect_test_case_count, resolve_packs


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
    tc_count = collect_test_case_count(resolve_packs(None))
    for i in range(1, tc_count + 1):
        tc = output / "test_cases" / f"TC-{i:02d}"
        assert tc.is_dir(), f"Missing {tc}"
        assert (tc / "input_files").is_dir(), f"Missing {tc}/input_files"


def test_generate_is_idempotent(tmp_path: Path) -> None:
    """Running generate() twice on the same output must not fail."""
    config = load_config("config.yaml")
    output = tmp_path / "suite"

    generate(config, output)
    generate(config, output)  # should not raise


# ---------------------------------------------------------------------------
# CLI --packs argument tests
# ---------------------------------------------------------------------------


class TestCLIPackSelection:
    """Test --packs CLI argument for selecting scenario packs."""

    def test_default_no_packs_flag(self, tmp_path: Path) -> None:
        """Omitting --packs produces default (accounting-core only) output."""
        output = tmp_path / "suite"
        main(["--output", str(output)])
        default_tc_count = collect_test_case_count(resolve_packs(None))
        for i in range(1, default_tc_count + 1):
            assert (output / "test_cases" / f"TC-{i:02d}").is_dir()
        # Should NOT have legal/HR test case dirs beyond default count
        assert not (output / "test_cases" / f"TC-{default_tc_count + 1:02d}").exists()

    def test_explicit_default_pack(self, tmp_path: Path) -> None:
        """--packs cascade_accounting_core is equivalent to the default."""
        output = tmp_path / "suite"
        main(["--output", str(output), "--packs", "cascade_accounting_core"])
        default_tc_count = collect_test_case_count(resolve_packs(None))
        for i in range(1, default_tc_count + 1):
            assert (output / "test_cases" / f"TC-{i:02d}").is_dir()

    def test_all_packs(self, tmp_path: Path) -> None:
        """--packs all generates every registered pack."""
        output = tmp_path / "suite"
        main(["--output", str(output), "--packs", "all"])
        all_tc_count = collect_test_case_count(resolve_packs(["all"]))
        for i in range(1, all_tc_count + 1):
            assert (output / "test_cases" / f"TC-{i:02d}").is_dir()
        # all packs should produce more TCs than default
        default_tc_count = collect_test_case_count(resolve_packs(None))
        assert all_tc_count > default_tc_count

    def test_unknown_pack_exits_with_error(self, tmp_path: Path, capsys) -> None:
        """--packs nonexistent exits with a clear error message."""
        output = tmp_path / "suite"
        with pytest.raises(SystemExit) as exc_info:
            main(["--output", str(output), "--packs", "nonexistent"])
        assert exc_info.value.code == 1
        captured = capsys.readouterr()
        assert "unknown pack" in captured.err
        assert "nonexistent" in captured.err

    def test_help_text_mentions_packs(self, capsys) -> None:
        """--help output documents the --packs flag."""
        with pytest.raises(SystemExit) as exc_info:
            main(["--help"])
        assert exc_info.value.code == 0
        captured = capsys.readouterr()
        assert "--packs" in captured.out
        assert "cascade_accounting_core" in captured.out
