"""Quality gate: every generated file opens cleanly with its native library.

Walks manifest.json after a full generation run and opens every file by type:
- xlsx → openpyxl
- docx → python-docx
- pdf  → pypdf
- csv  → stdlib csv
- json → stdlib json

No exceptions means the gate passes.
Ref: bead synth-data-fee, prompt.md §9.
"""

from __future__ import annotations

import csv
import json
from pathlib import Path

import openpyxl
import pypdf
from docx import Document

from generate_test_suite import generate
from generator.config import load_config

# Map manifest "type" to an opener that raises on corrupt data.
OPENERS = {
    "xlsx": lambda p: openpyxl.load_workbook(p, data_only=True),
    "docx": lambda p: Document(str(p)),
    "pdf": lambda p: _open_pdf(p),
    "csv": lambda p: _open_csv(p),
    "json": lambda p: json.loads(p.read_text()),
}


def _open_pdf(path: Path) -> None:
    reader = pypdf.PdfReader(str(path))
    # Force page parsing to catch truncated/corrupt PDFs
    for page in reader.pages:
        page.extract_text()


def _open_csv(path: Path) -> None:
    with open(path, newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        for _ in reader:
            pass


# ---------------------------------------------------------------------------
# Generate once, then parametrise over every manifest entry
# ---------------------------------------------------------------------------

_MANIFEST: list[dict] | None = None
_OUTPUT_DIR: Path | None = None


def _ensure_generated() -> tuple[list[dict], Path]:
    """Run the generator once and cache the result for the module."""
    global _MANIFEST, _OUTPUT_DIR  # noqa: PLW0603
    if _MANIFEST is None:
        import tempfile

        outdir = Path(tempfile.mkdtemp(prefix="file_integrity_"))
        cfg = load_config("config.yaml")
        generate(cfg, outdir)
        manifest_path = outdir / "manifest.json"
        _MANIFEST = json.loads(manifest_path.read_text())
        _OUTPUT_DIR = outdir
    return _MANIFEST, _OUTPUT_DIR


def _file_ids() -> list[str]:
    """Return manifest paths for parametrize IDs (runs generation on import)."""
    manifest, _ = _ensure_generated()
    return [entry["path"] for entry in manifest]


def _get_entry(path: str) -> dict:
    manifest, _ = _ensure_generated()
    for entry in manifest:
        if entry["path"] == path:
            return entry
    raise KeyError(path)


class TestFileIntegrity:
    """Every file listed in manifest.json must open without exceptions."""

    def test_manifest_not_empty(self):
        manifest, _ = _ensure_generated()
        assert len(manifest) > 0, "manifest.json is empty"

    def test_all_files_exist(self):
        manifest, outdir = _ensure_generated()
        missing = [e["path"] for e in manifest if not (outdir / e["path"]).is_file()]
        assert missing == [], f"Files in manifest but missing on disk: {missing}"

    def test_all_types_have_openers(self):
        manifest, _ = _ensure_generated()
        types = {e["type"] for e in manifest}
        unsupported = types - set(OPENERS.keys())
        assert unsupported == set(), f"No opener for types: {unsupported}"

    def test_open_xlsx_files(self):
        self._open_files_of_type("xlsx")

    def test_open_docx_files(self):
        self._open_files_of_type("docx")

    def test_open_pdf_files(self):
        self._open_files_of_type("pdf")

    def test_open_csv_files(self):
        self._open_files_of_type("csv")

    def test_open_json_files(self):
        self._open_files_of_type("json")

    def _open_files_of_type(self, filetype: str):
        manifest, outdir = _ensure_generated()
        entries = [e for e in manifest if e["type"] == filetype]
        if not entries:
            return  # No files of this type — nothing to check
        failures: list[str] = []
        for entry in entries:
            fpath = outdir / entry["path"]
            try:
                OPENERS[filetype](fpath)
            except Exception as exc:
                failures.append(f"{entry['path']}: {exc}")
        assert failures == [], (
            f"{len(failures)} {filetype} file(s) failed to open:\n"
            + "\n".join(failures)
        )
