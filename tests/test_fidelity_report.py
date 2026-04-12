"""Tests for generator.model.fidelity_report — fidelity regression report."""

from __future__ import annotations

from pathlib import Path

import pytest

from generator.config import load_config
from generator.model.build import build_model
from generator.model.fidelity_report import build_fidelity_report


@pytest.fixture(scope="module")
def default_config():
    root = Path(__file__).resolve().parent.parent / "config.yaml"
    return load_config(root)


@pytest.fixture(scope="module")
def default_model(default_config):
    return build_model(default_config)


@pytest.fixture(scope="module")
def report(default_model):
    return build_fidelity_report(default_model)


# ── Structure ──────────────────────────────────────────────────────────────


def test_report_is_dict_of_strings(report):
    assert isinstance(report, dict)
    for k, v in report.items():
        assert isinstance(k, str), f"key {k!r} is not a string"
        assert isinstance(v, str), f"value for {k!r} is not a string"


def test_report_has_expected_sections(report):
    sections = {k.split(".")[0] for k in report}
    expected = {
        "employees", "entities", "subsidiaries", "assets", "leases",
        "lease_schedule_rows", "gl_entries", "revenue_records", "opex_records",
        "revenue", "ar_aging", "ar_invoices", "ar_receipts", "ar_collections",
        "ar_allowance", "ic_netting", "tax", "rd", "legal", "hr",
    }
    missing = expected - sections
    assert not missing, f"Missing report sections: {missing}"


# ── Determinism ────────────────────────────────────────────────────────────


def test_report_deterministic(default_config):
    """Two independent model builds produce identical reports."""
    model1 = build_model(default_config)
    report1 = build_fidelity_report(model1)
    model2 = build_model(default_config)
    report2 = build_fidelity_report(model2)
    assert report1 == report2


# ── Data shape sanity ──────────────────────────────────────────────────────


def test_employee_count_reasonable(report):
    total = int(report["employees.total"])
    assert 800 <= total <= 900, f"Expected ~850 employees, got {total}"


def test_revenue_in_tolerance(report):
    rev = float(report["revenue.fy2025.gl_total"])
    assert 198_000_000 <= rev <= 202_000_000, f"FY2025 revenue {rev} out of range"


def test_ic_netting_balanced(report):
    for year in [2023, 2024, 2025]:
        assert report[f"ic_netting.fy{year}.balanced"] == "True", (
            f"IC netting not balanced for FY{year}"
        )


def test_ar_invoice_receipt_tieout(report):
    if "ar_invoices.total" in report and "ar_receipts.total" in report:
        assert report["ar_invoices.total"] == report["ar_receipts.total"], (
            "AR invoices and receipts should tie out"
        )


def test_tax_provisions_present(report):
    for year in [2023, 2024, 2025]:
        assert f"tax.fy{year}.total_provision" in report


def test_rd_qre_present(report):
    assert "rd.qre.total_qres" in report
    total = float(report["rd.qre.total_qres"])
    assert total > 0, "QRE total should be positive"


def test_leases_present(report):
    total = int(report["leases.total"])
    assert total > 0
    assert "leases.initial_liability_total" in report
