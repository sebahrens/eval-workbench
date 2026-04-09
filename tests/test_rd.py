"""Tests for generator.model.rd — R&D projects and time records (TC-08)."""

import random

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
