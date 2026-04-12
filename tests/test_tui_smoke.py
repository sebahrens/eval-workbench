"""Smoke tests for the TUI module skeleton.

These tests verify the import path and help output without requiring
textual to be installed — the full Textual pilot tests live in a
separate test file gated by pytest.importorskip("textual").
"""

from __future__ import annotations

import subprocess
import sys


def test_tui_module_importable():
    """The generator.tui package should be importable even without textual."""
    import generator.tui  # noqa: F401


def test_tui_main_help_flag():
    """``python -m generator.tui --help`` should exit 0 and show usage."""
    result = subprocess.run(
        [sys.executable, "-m", "generator.tui", "--help"],
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0
    assert "scenario configuration" in result.stdout.lower()


def test_configure_subcommand_help():
    """``generate_test_suite.py configure --help`` should exit 0."""
    result = subprocess.run(
        [sys.executable, "generate_test_suite.py", "configure", "--help"],
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0
    assert "configuration" in result.stdout.lower()


def test_generate_still_requires_output():
    """Without a subcommand, --output is still required."""
    result = subprocess.run(
        [sys.executable, "generate_test_suite.py"],
        capture_output=True,
        text=True,
    )
    assert result.returncode != 0
    assert "output" in result.stderr.lower()
