"""Tests for generator.model.validation — model-level scenario invariant checks."""

from __future__ import annotations

from pathlib import Path

import pytest

from generator.config import load_config
from generator.model.build import CascadeModel, build_model
from generator.model.validation import validate_model

# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------

@pytest.fixture(scope="module")
def default_config():
    """Load the real config.yaml from the project root."""
    root = Path(__file__).resolve().parent.parent / "config.yaml"
    return load_config(root)


@pytest.fixture(scope="module")
def default_model(default_config):
    """Build the canonical model once for the entire test module."""
    return build_model(default_config)


# ---------------------------------------------------------------------------
# Happy-path: default Cascade scenario passes all invariants
# ---------------------------------------------------------------------------

def test_default_model_passes_all_invariants(default_model, default_config):
    """The default Cascade model must pass every invariant check."""
    errors = validate_model(default_model, config=default_config)
    if errors:
        msg = "\n".join(f"  [{e.category}] {e.message}" for e in errors)
        pytest.fail(f"Invariant violations in default model:\n{msg}")


def test_default_model_passes_without_config(default_model):
    """validate_model works without config (skips headcount total check)."""
    errors = validate_model(default_model, config=None)
    if errors:
        msg = "\n".join(f"  [{e.category}] {e.message}" for e in errors)
        pytest.fail(f"Invariant violations:\n{msg}")


# ---------------------------------------------------------------------------
# Individual invariant checks
# ---------------------------------------------------------------------------

def test_entity_headcount_all_valid(default_model):
    """Every employee belongs to a known entity."""
    valid_codes = set(default_model.entities)
    for emp in default_model.employees:
        assert emp.entity_code in valid_codes, (
            f"Employee {emp.employee_id} has unknown entity '{emp.entity_code}'"
        )


def test_headcount_matches_config(default_model, default_config):
    """Total employee count matches config."""
    expected = default_config.company.employees.total_count
    assert len(default_model.employees) == expected


def test_subsidiaries_have_employees(default_model):
    """Every subsidiary has at least one employee."""
    for code in default_model.subsidiaries:
        entity_emps = [e for e in default_model.employees if e.entity_code == code]
        assert len(entity_emps) > 0, f"Subsidiary {code} has no employees"


def test_ic_netting_all_years(default_model):
    """IC accounts net to zero for all fiscal years."""
    import datetime

    from generator.model.consolidation import verify_ic_elimination

    for year in [2023, 2024, 2025]:
        eoy = datetime.date(year, 12, 31)
        is_zero, total, _ = verify_ic_elimination(default_model.ledger, eoy)
        assert is_zero, f"IC accounts do not net to zero for FY{year}: imbalance={total}"


def test_consolidated_revenue_in_range(default_model):
    """FY2025 consolidated revenue is within $198M–$202M."""
    from decimal import Decimal

    from generator.model.consolidation import build_income_statement

    is_stmt = build_income_statement(default_model.ledger, 2025)
    rev = is_stmt.total_revenue
    assert Decimal("198_000_000") <= rev <= Decimal("202_000_000"), (
        f"FY2025 revenue ${rev:,.0f} outside $198M–$202M range"
    )


# ---------------------------------------------------------------------------
# Error-path: validation catches bad model state
# ---------------------------------------------------------------------------

def test_detects_unknown_entity_employee(default_model, default_config):
    """Validation catches an employee with a bogus entity code."""
    import datetime

    from generator.model.employees import Employee

    # Create a model with a bad employee injected
    bad_emp = Employee(
        employee_id="E-9999",
        name="Test Bad",
        entity_code="ZZ",  # Not a real entity
        entity_name="Bogus Entity",
        department="Engineering",
        title="Engineer",
        hire_date=datetime.date(2023, 1, 1),
        annual_salary=80000,
        state="OR",
        cost_center="CC-ZZ-001",
        is_rd_eligible=False,
        termination_date=None,
    )
    patched = CascadeModel(
        entities=default_model.entities,
        subsidiaries=default_model.subsidiaries,
        ledger=default_model.ledger,
        revenue_records=default_model.revenue_records,
        employees=list(default_model.employees) + [bad_emp],
        opex_records=default_model.opex_records,
        assets=default_model.assets,
        leases=default_model.leases,
        lease_schedules=default_model.lease_schedules,
        tax_provisions=default_model.tax_provisions,
        bank=default_model.bank,
    )
    errors = validate_model(patched, config=default_config)
    headcount_errors = [e for e in errors if e.category == "headcount"]
    assert any("ZZ" in e.message for e in headcount_errors), (
        "Should detect employee with unknown entity 'ZZ'"
    )


def test_detects_headcount_mismatch(default_model, default_config):
    """Validation catches when employee count doesn't match config."""
    # Drop one employee
    patched = CascadeModel(
        entities=default_model.entities,
        subsidiaries=default_model.subsidiaries,
        ledger=default_model.ledger,
        revenue_records=default_model.revenue_records,
        employees=default_model.employees[:-1],
        opex_records=default_model.opex_records,
        assets=default_model.assets,
        leases=default_model.leases,
        lease_schedules=default_model.lease_schedules,
        tax_provisions=default_model.tax_provisions,
        bank=default_model.bank,
    )
    errors = validate_model(patched, config=default_config)
    headcount_errors = [e for e in errors if e.category == "headcount"]
    assert any("mismatch" in e.message.lower() for e in headcount_errors)
