"""Canary generator and registry emitter (section 1.6 of prompt.md).

Produces deterministic 8-character alphanumeric canary codes seeded from
the project seed.  Provides helpers to embed canaries into xlsx metadata,
docx custom properties, PDF metadata, and CSV comment lines.

The registry is a JSON file mapping each canary to the file it belongs to
and the location where it was embedded.
"""

from __future__ import annotations

import json
import random
import string
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from docx import Document
    from fpdf import FPDF
    from openpyxl import Workbook
    from reportlab.pdfgen.canvas import Canvas


# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass
class CanaryEntry:
    """One entry in the canary registry."""

    file_key: str          # Logical file identifier (e.g. "cascade_tb_fy2025")
    canary: str            # 8-char alphanumeric code
    file_path: str = ""    # Relative path once the file is written
    location: str = ""     # Where the canary was embedded (e.g. "Sheet 'TB', Cell A1 comment")


@dataclass
class CanaryRegistry:
    """Complete canary registry for the test suite."""

    entries: dict[str, CanaryEntry] = field(default_factory=dict)

    # -- lookup ---------------------------------------------------------------

    def get(self, file_key: str) -> CanaryEntry:
        """Return the entry for *file_key*, raising KeyError if absent."""
        return self.entries[file_key]

    def canary_for(self, file_key: str) -> str:
        """Return just the canary string for *file_key*."""
        return self.entries[file_key].canary

    # -- mutation -------------------------------------------------------------

    def set_location(self, file_key: str, file_path: str, location: str) -> None:
        """Record where the canary was actually embedded after file generation."""
        entry = self.entries[file_key]
        entry.file_path = file_path
        entry.location = location

    # -- serialisation --------------------------------------------------------

    def to_dict(self) -> list[dict]:
        """Return a sorted list of entry dicts (deterministic order)."""
        return [asdict(self.entries[k]) for k in sorted(self.entries)]

    def write_json(self, path: str | Path) -> None:
        """Write the registry to *path* as formatted JSON."""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w") as f:
            json.dump(self.to_dict(), f, indent=2, sort_keys=True)
            f.write("\n")


# ---------------------------------------------------------------------------
# Generator
# ---------------------------------------------------------------------------

_ALPHABET = string.ascii_uppercase + string.digits


def generate_canary(rng: random.Random) -> str:
    """Return a single 8-character alphanumeric canary code."""
    return "".join(rng.choices(_ALPHABET, k=8))


def build_registry(file_keys: list[str], seed: int = 42) -> CanaryRegistry:
    """Build a CanaryRegistry with unique canaries for each *file_key*.

    Uses a dedicated Random instance seeded from *seed* so that canary
    generation is isolated from other random state.

    Parameters
    ----------
    file_keys:
        Sorted list of logical file identifiers.  The order **must** be
        deterministic (caller should sort before passing).
    seed:
        Integer seed for the canary RNG.
    """
    rng = random.Random(seed)
    registry = CanaryRegistry()
    seen: set[str] = set()

    for key in file_keys:
        # Generate canaries until we get one that is unique.
        canary = generate_canary(rng)
        while canary in seen:
            canary = generate_canary(rng)
        seen.add(canary)
        registry.entries[key] = CanaryEntry(file_key=key, canary=canary)

    return registry


# ---------------------------------------------------------------------------
# Embedding helpers
# ---------------------------------------------------------------------------

def embed_canary_xlsx(wb: Workbook, canary: str) -> str:
    """Embed *canary* as a custom document property on an openpyxl Workbook.

    Returns a human-readable description of where the canary was placed.
    """
    if wb.properties is None:  # pragma: no cover — openpyxl always creates one
        from openpyxl.packaging.core import DocumentProperties
        wb.properties = DocumentProperties()
    wb.properties.description = f"CANARY: {canary}"
    return "Document properties → description"


def embed_canary_docx(doc: Document, canary: str) -> str:
    """Embed *canary* as a core property (comments field) on a python-docx Document.

    Returns a description of the embedding location.
    """
    doc.core_properties.comments = f"CANARY: {canary}"
    return "Core properties → comments"


def embed_canary_pdf_reportlab(canvas: Canvas, canary: str) -> str:
    """Set *canary* as PDF Author metadata on a reportlab Canvas.

    Returns a description of the embedding location.
    """
    canvas.setAuthor(f"CANARY: {canary}")
    return "PDF metadata → Author"


def embed_canary_pdf_fpdf2(pdf: FPDF, canary: str) -> str:
    """Set *canary* as PDF subject metadata on an fpdf2 document.

    Returns a description of the embedding location.
    """
    pdf.set_subject(f"CANARY: {canary}")
    return "PDF metadata → Subject"


def embed_canary_csv_comment(canary: str) -> str:
    """Return a comment line to prepend to a CSV file.

    The caller is responsible for writing this as the first line of the file.
    """
    return f"# CANARY: {canary}\n"
