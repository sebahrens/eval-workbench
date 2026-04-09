"""Tests for generator.canaries — canary generation and registry."""

from __future__ import annotations

import json
import re
from pathlib import Path

from generator.canaries import (
    build_registry,
    embed_canary_csv_comment,
    embed_canary_docx,
    embed_canary_pdf_fpdf2,
    embed_canary_pdf_reportlab,
    embed_canary_xlsx,
    generate_canary,
)

# ---------------------------------------------------------------------------
# Basic canary generation
# ---------------------------------------------------------------------------

CANARY_RE = re.compile(r"^[A-Z0-9]{8}$")


class TestGenerateCanary:
    def test_format(self):
        """Canary is 8 uppercase-alphanumeric characters."""
        import random
        rng = random.Random(42)
        for _ in range(50):
            assert CANARY_RE.match(generate_canary(rng))

    def test_deterministic(self):
        """Same seed produces the same sequence."""
        import random
        a = [generate_canary(random.Random(99)) for _ in range(10)]
        b = [generate_canary(random.Random(99)) for _ in range(10)]
        assert a == b


# ---------------------------------------------------------------------------
# Registry builder
# ---------------------------------------------------------------------------

SAMPLE_KEYS = sorted([
    "cascade_tb_fy2025",
    "cascade_tb_fy2024",
    "employee_roster",
    "master_coa",
    "bank_statements",
])


class TestBuildRegistry:
    def test_all_keys_present(self):
        reg = build_registry(SAMPLE_KEYS)
        assert sorted(reg.entries) == SAMPLE_KEYS

    def test_canaries_unique(self):
        reg = build_registry(SAMPLE_KEYS)
        canaries = [e.canary for e in reg.entries.values()]
        assert len(set(canaries)) == len(canaries)

    def test_deterministic(self):
        r1 = build_registry(SAMPLE_KEYS, seed=42)
        r2 = build_registry(SAMPLE_KEYS, seed=42)
        for key in SAMPLE_KEYS:
            assert r1.canary_for(key) == r2.canary_for(key)

    def test_different_seed_different_canaries(self):
        r1 = build_registry(SAMPLE_KEYS, seed=42)
        r2 = build_registry(SAMPLE_KEYS, seed=99)
        # Extremely unlikely all match by chance
        assert any(
            r1.canary_for(k) != r2.canary_for(k) for k in SAMPLE_KEYS
        )


# ---------------------------------------------------------------------------
# Registry serialisation
# ---------------------------------------------------------------------------

class TestRegistryJson:
    def test_write_and_read(self, tmp_path: Path):
        reg = build_registry(SAMPLE_KEYS)
        reg.set_location("master_coa", "shared_data/master_coa.xlsx", "Properties → description")

        out = tmp_path / "canary_registry.json"
        reg.write_json(out)

        data = json.loads(out.read_text())
        assert isinstance(data, list)
        assert len(data) == len(SAMPLE_KEYS)
        # Must be sorted by file_key
        assert [d["file_key"] for d in data] == SAMPLE_KEYS

    def test_deterministic_json(self, tmp_path: Path):
        """Two writes from the same seed produce identical JSON."""
        r1 = build_registry(SAMPLE_KEYS, seed=42)
        r2 = build_registry(SAMPLE_KEYS, seed=42)

        p1 = tmp_path / "a.json"
        p2 = tmp_path / "b.json"
        r1.write_json(p1)
        r2.write_json(p2)
        assert p1.read_text() == p2.read_text()


# ---------------------------------------------------------------------------
# Embedding helpers
# ---------------------------------------------------------------------------

class TestEmbedXlsx:
    def test_embed(self):
        from openpyxl import Workbook
        wb = Workbook()
        loc = embed_canary_xlsx(wb, "AB12CD34")
        assert "CANARY: AB12CD34" in wb.properties.description
        assert loc  # non-empty description

    def test_findable_in_saved_file(self, tmp_path: Path):
        """Canary is findable when the workbook is saved and re-opened."""
        from openpyxl import Workbook, load_workbook
        wb = Workbook()
        embed_canary_xlsx(wb, "XK7P2M9Q")
        path = tmp_path / "test.xlsx"
        wb.save(path)

        wb2 = load_workbook(path)
        assert "CANARY: XK7P2M9Q" in (wb2.properties.description or "")


class TestEmbedDocx:
    def test_embed(self):
        from docx import Document
        doc = Document()
        loc = embed_canary_docx(doc, "EF56GH78")
        assert "CANARY: EF56GH78" in doc.core_properties.comments
        assert loc

    def test_findable_in_saved_file(self, tmp_path: Path):
        from docx import Document
        doc = Document()
        embed_canary_docx(doc, "LM3N8R2T")
        path = tmp_path / "test.docx"
        doc.save(path)

        doc2 = Document(str(path))
        assert "CANARY: LM3N8R2T" in (doc2.core_properties.comments or "")


class TestEmbedPdfReportlab:
    def test_embed(self, tmp_path: Path):
        from reportlab.pdfgen.canvas import Canvas
        path = tmp_path / "test.pdf"
        c = Canvas(str(path))
        loc = embed_canary_pdf_reportlab(c, "QW5E9Y1A")
        c.save()
        assert loc

        # Verify canary is in the raw PDF bytes
        raw = path.read_bytes()
        assert b"CANARY: QW5E9Y1A" in raw


class TestEmbedPdfFpdf2:
    def test_embed(self, tmp_path: Path):
        from fpdf import FPDF
        pdf = FPDF()
        pdf.add_page()
        loc = embed_canary_pdf_fpdf2(pdf, "ZT4R7K1W")
        path = tmp_path / "test.pdf"
        pdf.output(str(path))
        assert loc

        raw = path.read_bytes()
        assert b"CANARY: ZT4R7K1W" in raw


class TestEmbedCsv:
    def test_comment_format(self):
        line = embed_canary_csv_comment("NP8Q3V6X")
        assert line == "# CANARY: NP8Q3V6X\n"

    def test_findable_in_file(self, tmp_path: Path):
        path = tmp_path / "test.csv"
        line = embed_canary_csv_comment("NP8Q3V6X")
        path.write_text(line + "col1,col2\n1,2\n")
        content = path.read_text()
        assert "CANARY: NP8Q3V6X" in content
