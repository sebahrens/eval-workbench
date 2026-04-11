"""Tests for generator.model.rd — R&D projects, time records, and payroll tie-outs (TC-08)."""

import random
from decimal import Decimal
from pathlib import Path

import pytest

from generator.config import load_config
from generator.model.build import build_model
from generator.model.employees import generate_employees
from generator.model.rd import (
    QRE_2025_TARGET,
    RD_EMPLOYEE_COUNT,
    RD_PROJECTS,
    TARGET_CREDIT,
    compute_qres,
    generate_rd_projects,
    generate_supply_expenses,
    generate_time_records,
)


def _make_roster():
    rng = random.Random(42)
    return generate_employees(rng)


class TestRDProjects:
    """R&D project definitions."""

    def test_project_count(self) -> None:
        projects = generate_rd_projects()
        assert len(projects) == 12

    def test_qualification_breakdown(self) -> None:
        """8 qualify, 2 borderline, 2 don't."""
        projects = generate_rd_projects()
        quals = [p.qualifies for p in projects]
        assert quals.count("yes") == 8
        assert quals.count("borderline") == 2
        assert quals.count("no") == 2

    def test_unique_codes(self) -> None:
        projects = generate_rd_projects()
        codes = [p.code for p in projects]
        assert len(codes) == len(set(codes))

    def test_non_qualifying_have_zero_weight(self) -> None:
        for p in RD_PROJECTS:
            if p.qualifies == "no":
                assert p.qre_weight == 0.0, f"{p.code} should have zero QRE weight"


class TestTimeRecords:
    """Weekly time record generation."""

    def test_row_count_approximate(self) -> None:
        """Should be ~2,340 rows (45 emp × 52 weeks, 1 row per emp per week)."""
        roster = _make_roster()
        rng = random.Random(42)
        records = generate_time_records(roster, rng)
        # Exactly 1 row per active employee per active week.
        # Some employees may be hired late or terminated, so slightly < 2,340.
        assert 2200 <= len(records) <= 2340

    def test_all_employees_present(self) -> None:
        roster = _make_roster()
        rng = random.Random(42)
        records = generate_time_records(roster, rng)
        emp_ids = {r.employee_id for r in records}
        assert len(emp_ids) == RD_EMPLOYEE_COUNT

    def test_hours_reasonable(self) -> None:
        roster = _make_roster()
        rng = random.Random(42)
        records = generate_time_records(roster, rng)
        for rec in records:
            assert 0.5 <= rec.hours <= 50.0

    def test_valid_project_codes(self) -> None:
        """All project codes must be R&D projects or overhead codes."""
        roster = _make_roster()
        rng = random.Random(42)
        records = generate_time_records(roster, rng)
        rd_codes = {p.code for p in RD_PROJECTS}
        gen_codes = {f"GEN-{i:03d}" for i in range(1, 7)}
        valid_codes = rd_codes | gen_codes
        for rec in records:
            assert rec.project_code in valid_codes, f"Invalid code: {rec.project_code}"

    def test_has_overhead_and_rd_time(self) -> None:
        """Records should include both R&D project and overhead time."""
        roster = _make_roster()
        rng = random.Random(42)
        records = generate_time_records(roster, rng)
        codes = {r.project_code for r in records}
        assert any(c.startswith("RD-") for c in codes), "No R&D project time"
        assert any(c.startswith("GEN-") for c in codes), "No overhead time"

    def test_deterministic(self) -> None:
        roster = _make_roster()
        r1 = generate_time_records(roster, random.Random(42))
        r2 = generate_time_records(roster, random.Random(42))
        assert len(r1) == len(r2)
        for a, b in zip(r1, r2):
            assert a == b


class TestSupplyExpenses:
    """R&D supply expense generation."""

    def test_all_qualifying_projects(self) -> None:
        rng = random.Random(42)
        expenses = generate_supply_expenses(rng)
        qualifying_codes = {p.code for p in RD_PROJECTS if p.qualifies == "yes"}
        for exp in expenses:
            assert exp.project_code in qualifying_codes

    def test_cost_center(self) -> None:
        rng = random.Random(42)
        expenses = generate_supply_expenses(rng)
        for exp in expenses:
            assert exp.cost_center == "3500"

    def test_reasonable_count(self) -> None:
        rng = random.Random(42)
        expenses = generate_supply_expenses(rng)
        assert 96 <= len(expenses) <= 144  # 8-12 per month × 12 months


class TestQREComputation:
    """QRE computation and ASC credit."""

    def test_credit_within_tolerance(self) -> None:
        """ASC credit must be within $500 of $185,000."""
        roster = _make_roster()
        rng = random.Random(42)
        records = generate_time_records(roster, rng)
        expenses = generate_supply_expenses(random.Random(42))
        result = compute_qres(records, expenses, roster)

        assert abs(result.credit - TARGET_CREDIT) <= 500, (
            f"Credit {result.credit} not within $500 of {TARGET_CREDIT}"
        )

    def test_qre_total_near_target(self) -> None:
        roster = _make_roster()
        rng = random.Random(42)
        records = generate_time_records(roster, rng)
        expenses = generate_supply_expenses(random.Random(42))
        result = compute_qres(records, expenses, roster)

        # Total QREs should be close to target (supply scaling makes it exact-ish).
        assert abs(result.total_qres - QRE_2025_TARGET) < 1000

    def test_project_breakdown_exists(self) -> None:
        roster = _make_roster()
        rng = random.Random(42)
        records = generate_time_records(roster, rng)
        expenses = generate_supply_expenses(random.Random(42))
        result = compute_qres(records, expenses, roster)

        assert len(result.qres_by_project) > 0
        assert all(v > 0 for v in result.qres_by_project.values())

    def test_wage_plus_supply_equals_total(self) -> None:
        """Arithmetic tie-out: wage_qres + supply_qres == total_qres."""
        roster = _make_roster()
        records = generate_time_records(roster, random.Random(42))
        expenses = generate_supply_expenses(random.Random(42))
        result = compute_qres(records, expenses, roster)

        assert result.wage_qres + result.supply_qres == result.total_qres

    def test_project_breakdown_sums_to_total(self) -> None:
        """Sum of per-project QREs == total_qres."""
        roster = _make_roster()
        records = generate_time_records(roster, random.Random(42))
        expenses = generate_supply_expenses(random.Random(42))
        result = compute_qres(records, expenses, roster)

        project_sum = sum(result.qres_by_project.values())
        assert abs(project_sum - result.total_qres) < Decimal("1.00")

    def test_qre_deterministic(self) -> None:
        """Full QRE computation is deterministic across reruns."""
        roster = _make_roster()

        r1_records = generate_time_records(roster, random.Random(42))
        r1_expenses = generate_supply_expenses(random.Random(42))
        r1 = compute_qres(r1_records, r1_expenses, roster)

        r2_records = generate_time_records(roster, random.Random(42))
        r2_expenses = generate_supply_expenses(random.Random(42))
        r2 = compute_qres(r2_records, r2_expenses, roster)

        assert r1.wage_qres == r2.wage_qres
        assert r1.supply_qres == r2.supply_qres
        assert r1.total_qres == r2.total_qres
        assert r1.credit == r2.credit
        assert r1.qres_by_project == r2.qres_by_project


# ── Payroll & Model Tie-Out Tests ──────────────────────────────────────────


@pytest.fixture(scope="module")
def _model():
    """Build the full canonical model once for tie-out tests."""
    root = Path(__file__).resolve().parent.parent / "config.yaml"
    config = load_config(root)
    return build_model(config)


class TestRDPayrollTieOuts:
    """Cross-referential integrity: R&D time records ↔ employee roster ↔ QRE."""

    def test_time_record_employees_in_roster(self, _model) -> None:
        """Every employee in time records exists in the canonical roster."""
        roster_ids = {e.employee_id for e in _model.employees}
        for rec in _model.rd_time_records:
            assert rec.employee_id in roster_ids, (
                f"Time record references {rec.employee_id} not in employee roster"
            )

    def test_time_record_employees_are_rd_eligible(self, _model) -> None:
        """Every employee in time records is flagged R&D-eligible."""
        eligible_ids = {
            e.employee_id for e in _model.employees if e.is_rd_eligible
        }
        tr_ids = {r.employee_id for r in _model.rd_time_records}
        for eid in tr_ids:
            assert eid in eligible_ids, (
                f"{eid} has time records but is_rd_eligible=False"
            )

    def test_time_record_employees_are_am_entity(self, _model) -> None:
        """R&D employees all belong to Advanced Materials (AM)."""
        emp_by_id = {e.employee_id: e for e in _model.employees}
        tr_ids = {r.employee_id for r in _model.rd_time_records}
        for eid in tr_ids:
            assert emp_by_id[eid].entity_code == "AM", (
                f"{eid} is entity {emp_by_id[eid].entity_code}, expected AM"
            )

    def test_time_record_names_match_roster(self, _model) -> None:
        """Employee names in time records match the canonical roster."""
        emp_by_id = {e.employee_id: e for e in _model.employees}
        for rec in _model.rd_time_records:
            assert rec.employee_name == emp_by_id[rec.employee_id].name, (
                f"{rec.employee_id}: time record name '{rec.employee_name}' "
                f"!= roster name '{emp_by_id[rec.employee_id].name}'"
            )

    def test_qre_wages_use_w2_salaries(self, _model) -> None:
        """Wage QREs are bounded by total W-2 salaries of R&D employees.

        The sum of wage QREs cannot exceed the sum of W-2 salaries for the
        45 employees who log time — it's a fraction of their salaries.
        """
        tr_ids = {r.employee_id for r in _model.rd_time_records}
        total_w2 = sum(
            e.annual_salary for e in _model.employees
            if e.employee_id in tr_ids
        )
        assert _model.rd_qre_result.wage_qres < Decimal(total_w2), (
            f"Wage QREs ({_model.rd_qre_result.wage_qres}) exceed "
            f"total W-2 salaries ({total_w2})"
        )
        # Wage QREs should also be meaningfully positive
        assert _model.rd_qre_result.wage_qres > Decimal("500_000"), (
            f"Wage QREs suspiciously low: {_model.rd_qre_result.wage_qres}"
        )

    def test_model_rd_fields_populated(self, _model) -> None:
        """CascadeModel has all R&D fields populated after build."""
        assert len(_model.rd_time_records) > 0
        assert len(_model.rd_supply_expenses) > 0
        assert _model.rd_qre_result is not None
        assert _model.rd_qre_result.year == 2025

    def test_model_rd_credit_within_tolerance(self, _model) -> None:
        """Model-level credit matches target (end-to-end through build_model)."""
        assert abs(_model.rd_qre_result.credit - TARGET_CREDIT) <= 500

    def test_model_rd_reproducible(self) -> None:
        """Building the model twice produces identical R&D outputs."""
        root = Path(__file__).resolve().parent.parent / "config.yaml"
        config = load_config(root)
        m1 = build_model(config)
        m2 = build_model(config)

        assert len(m1.rd_time_records) == len(m2.rd_time_records)
        for a, b in zip(m1.rd_time_records, m2.rd_time_records):
            assert a == b

        assert len(m1.rd_supply_expenses) == len(m2.rd_supply_expenses)
        for a, b in zip(m1.rd_supply_expenses, m2.rd_supply_expenses):
            assert a == b

        assert m1.rd_qre_result.credit == m2.rd_qre_result.credit
        assert m1.rd_qre_result.total_qres == m2.rd_qre_result.total_qres
        assert m1.rd_qre_result.qres_by_project == m2.rd_qre_result.qres_by_project

    def test_no_qualifying_projects_excluded_from_qre(self, _model) -> None:
        """Non-qualifying projects (RD-011, RD-012) must not appear in QRE breakdown."""
        non_qualifying = {p.code for p in RD_PROJECTS if p.qualifies == "no"}
        for code in non_qualifying:
            assert code not in _model.rd_qre_result.qres_by_project, (
                f"Non-qualifying project {code} found in QRE breakdown"
            )
