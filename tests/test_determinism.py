"""Quality gate: byte-identical determinism across generator reruns.

Runs generate_test_suite.py twice to separate output directories and asserts
that every generated file is byte-identical.  Also includes a targeted PDF
sub-test that will catch reportlab invariant=True and fpdf2 creation_date
gotchas once PDF formatters are wired in.

Ref: prompt.md §9, bead synth-data-gm2.
"""

from __future__ import annotations

import filecmp
import subprocess
import sys
from pathlib import Path

from generate_test_suite import generate
from generator.config import load_config


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


# -- In-process determinism test (fast, default) ----------------------------


def test_determinism_in_process(tmp_path: Path) -> None:
    """Two in-process generate() calls must produce byte-identical trees."""
    config = load_config("config.yaml")

    run1 = tmp_path / "run1"
    run2 = tmp_path / "run2"

    generate(config, run1)
    generate(config, run2)

    _assert_trees_identical(run1, run2)


# -- Subprocess determinism test (catches module-level side effects) --------


def test_determinism_subprocess(tmp_path: Path) -> None:
    """Two separate process invocations must produce byte-identical trees."""
    run1 = tmp_path / "run1"
    run2 = tmp_path / "run2"

    for out_dir in (run1, run2):
        result = subprocess.run(
            [
                sys.executable,
                "generate_test_suite.py",
                "--config",
                "config.yaml",
                "--output",
                str(out_dir),
            ],
            capture_output=True,
            text=True,
        )
        assert result.returncode == 0, (
            f"Generator failed for {out_dir}:\nstdout: {result.stdout}\nstderr: {result.stderr}"
        )

    _assert_trees_identical(run1, run2)


# -- PDF-specific determinism (activates once PDFs exist) -------------------


def test_pdf_determinism(tmp_path: Path) -> None:
    """Any PDF emitted by the generator must be byte-identical across runs.

    This test is a no-op while the generator emits no PDFs, but will
    automatically activate once PDF formatters land.
    """
    config = load_config("config.yaml")

    run1 = tmp_path / "run1"
    run2 = tmp_path / "run2"

    generate(config, run1)
    generate(config, run2)

    pdfs_1 = sorted(p.relative_to(run1) for p in run1.rglob("*.pdf"))
    pdfs_2 = sorted(p.relative_to(run2) for p in run2.rglob("*.pdf"))

    assert pdfs_1 == pdfs_2, f"PDF file listings differ: {pdfs_1} vs {pdfs_2}"

    for rel in pdfs_1:
        fa = run1 / rel
        fb = run2 / rel
        assert filecmp.cmp(fa, fb, shallow=False), (
            f"PDF not byte-identical: {rel} — check reportlab invariant=True "
            f"and fpdf2 creation_date pinning"
        )
