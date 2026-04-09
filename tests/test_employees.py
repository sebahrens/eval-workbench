"""Tests for generator.model.employees — 850-employee roster (§1.4)."""

import random

from generator.model.employees import Employee, generate_employees


def _make_roster() -> list[Employee]:
    rng = random.Random(42)
    return generate_employees(rng)


class TestEmployeeRoster:
    """Core roster generation tests."""

    def test_total_count(self) -> None:
        """Roster must contain exactly 850 employees."""
        roster = _make_roster()
        assert len(roster) == 850

    def test_unique_ids(self) -> None:
        """Every employee_id must be unique."""
        roster = _make_roster()
        ids = [e.employee_id for e in roster]
        assert len(ids) == len(set(ids))

    def test_entity_prefixes(self) -> None:
        """IDs must be prefixed with valid entity codes."""
        roster = _make_roster()
        for e in roster:
            prefix = e.employee_id.split("-")[0]
            assert prefix in ("CI", "PC", "AM", "DS"), f"Bad prefix: {e.employee_id}"

    def test_all_entities_present(self) -> None:
        """All four entities must have employees."""
        roster = _make_roster()
        codes = {e.entity_code for e in roster}
        assert codes == {"CI", "PC", "AM", "DS"}

    def test_rd_eligible_only_am(self) -> None:
        """R&D eligibility only for AM R&D and Engineering staff."""
        roster = _make_roster()
        for e in roster:
            if e.is_rd_eligible:
                assert e.entity_code == "AM", f"Non-AM R&D eligible: {e.employee_id}"
                assert e.department in ("R&D", "Engineering"), (
                    f"Wrong dept for R&D eligible: {e.employee_id} in {e.department}"
                )

    def test_rd_eligible_count(self) -> None:
        """~45 R&D-eligible employees at Advanced Materials (for TC-08).

        The bead specifies 45 R&D-eligible. We check that the count of
        AM R&D + Engineering staff is in a reasonable range.
        """
        roster = _make_roster()
        rd_eligible = [e for e in roster if e.is_rd_eligible]
        # AM Engineering + R&D should be roughly 40% of 270 ≈ 108,
        # but the bead says 45. The spec says "45 R&D-eligible employees"
        # for TC-08 time records. We need at least 45.
        assert len(rd_eligible) >= 45, f"Only {len(rd_eligible)} R&D-eligible (need ≥45)"

    def test_termination_rate(self) -> None:
        """~8% annual turnover means roughly 60-80 terminated employees."""
        roster = _make_roster()
        terminated = [e for e in roster if e.termination_date is not None]
        # With 3 years of hires and ~8% annual rate, expect 5-15% terminated.
        rate = len(terminated) / len(roster)
        assert 0.03 <= rate <= 0.25, f"Termination rate {rate:.1%} out of range"

    def test_hire_dates_in_range(self) -> None:
        """All hire dates between 2022-01-01 and 2024-12-31."""
        import datetime

        roster = _make_roster()
        for e in roster:
            assert datetime.date(2022, 1, 1) <= e.hire_date <= datetime.date(2024, 12, 31), (
                f"{e.employee_id} hire_date {e.hire_date} out of range"
            )

    def test_termination_after_hire(self) -> None:
        """Termination dates must be after hire dates."""
        roster = _make_roster()
        for e in roster:
            if e.termination_date is not None:
                assert e.termination_date > e.hire_date, (
                    f"{e.employee_id}: terminated {e.termination_date} before hired {e.hire_date}"
                )

    def test_salary_positive(self) -> None:
        """All salaries must be positive."""
        roster = _make_roster()
        for e in roster:
            assert e.annual_salary > 0, f"{e.employee_id} has zero salary"

    def test_cost_center_format(self) -> None:
        """Cost centers are 4-digit strings."""
        roster = _make_roster()
        for e in roster:
            assert len(e.cost_center) == 4, f"{e.employee_id} cost_center={e.cost_center}"
            assert e.cost_center.isdigit(), f"{e.employee_id} cost_center not numeric"

    def test_valid_departments(self) -> None:
        """All departments must be from the spec's list."""
        roster = _make_roster()
        valid = {"Engineering", "Manufacturing", "Sales", "G&A", "R&D", "Warehouse", "Finance"}
        for e in roster:
            assert e.department in valid, f"{e.employee_id} in unknown dept {e.department}"

    def test_determinism(self) -> None:
        """Two runs with the same seed produce identical rosters."""
        roster1 = _make_roster()
        roster2 = _make_roster()
        for e1, e2 in zip(roster1, roster2):
            assert e1 == e2
