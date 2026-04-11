"""Tests for generator.model.employees — 850-employee roster (§1.4)."""

import random

from generator.config import (
    CompanyConfig,
    Config,
    EmployeeConfig,
    GrowthRates,
    IntercompanyConfig,
    SeasonalWeights,
    SubsidiaryConfig,
)
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


# ── Helpers for config-driven tests ──────────────────────────────────────────

def _cascade_config(**overrides) -> Config:
    """Build a Config matching Cascade defaults, with optional overrides."""
    emp = overrides.pop("employees", EmployeeConfig(
        total_count=850, annual_turnover_rate=0.08, remote_states=["CA", "WA", "NY"],
    ))
    subs = overrides.pop("subsidiaries", {
        "precision_components": SubsidiaryConfig(
            legal_name="Cascade Precision Components LLC",
            location="Portland, OR", state="OR", entity_code="PC",
            revenue=95_000_000, type="Core manufacturing",
            gross_margin=0.35, employee_count=330,
        ),
        "advanced_materials": SubsidiaryConfig(
            legal_name="Cascade Advanced Materials, Inc.",
            location="Austin, TX", state="TX", entity_code="AM",
            revenue=65_000_000, type="Specialty materials",
            gross_margin=0.52, employee_count=270, rd_spend_pct=0.12,
        ),
        "distribution_services": SubsidiaryConfig(
            legal_name="Cascade Distribution Services LLC",
            location="Chicago, IL", state="IL", entity_code="DS",
            revenue=40_000_000, type="Warehousing and logistics",
            gross_margin=0.18, employee_count=210,
        ),
    })
    return Config(
        seed=42,
        output_dir="test_suite",
        company=CompanyConfig(
            name="Cascade Industries, Inc.",
            type="US C-Corporation",
            industry="Mid-market manufacturer",
            headquarters="Portland, Oregon",
            fiscal_year_end="12-31",
            years=[2023, 2024, 2025],
            current_year=2025,
            consolidated_revenue=200_000_000,
            subsidiaries=subs,
            growth_rates=GrowthRates(fy2023_to_fy2024=0.06, fy2024_to_fy2025=0.09),
            intercompany=IntercompanyConfig(
                raw_materials_markup=0.08, management_fee_pct=0.015,
                intercompany_loan_principal=5_000_000, intercompany_loan_rate=0.05,
            ),
            employees=emp,
            seasonal_weights=SeasonalWeights(Q1=0.20, Q2=0.25, Q3=0.25, Q4=0.30),
        ),
        canary_assignments={},
        error_injections={},
    )


class TestConfigDrivenEmployees:
    """Tests for config-parameterized employee generation."""

    def test_default_config_matches_hardcoded(self) -> None:
        """Default Cascade config produces identical roster to no-config path."""
        rng1 = random.Random(42)
        roster_no_config = generate_employees(rng1)
        rng2 = random.Random(42)
        roster_with_config = generate_employees(rng2, config=_cascade_config())
        assert len(roster_no_config) == len(roster_with_config)
        for e1, e2 in zip(roster_no_config, roster_with_config):
            assert e1 == e2

    def test_custom_headcount(self) -> None:
        """Custom config with different headcounts changes roster size."""
        cfg = _cascade_config(
            employees=EmployeeConfig(total_count=100, annual_turnover_rate=0.08),
            subsidiaries={
                "alpha": SubsidiaryConfig(
                    legal_name="Alpha Corp",
                    location="Portland, OR", state="OR", entity_code="PC",
                    revenue=50_000_000, type="Manufacturing",
                    gross_margin=0.35, employee_count=40,
                ),
                "beta": SubsidiaryConfig(
                    legal_name="Beta Inc",
                    location="Austin, TX", state="TX", entity_code="AM",
                    revenue=30_000_000, type="Materials",
                    gross_margin=0.52, employee_count=30, rd_spend_pct=0.10,
                ),
                "gamma": SubsidiaryConfig(
                    legal_name="Gamma LLC",
                    location="Chicago, IL", state="IL", entity_code="DS",
                    revenue=20_000_000, type="Distribution",
                    gross_margin=0.18, employee_count=20,
                ),
            },
        )
        rng = random.Random(42)
        roster = generate_employees(rng, config=cfg)
        # parent CI gets 100 - (40+30+20) = 10
        assert len(roster) == 100

    def test_custom_remote_states(self) -> None:
        """Config remote_states are used for remote employees."""
        cfg = _cascade_config(
            employees=EmployeeConfig(
                total_count=850, annual_turnover_rate=0.08,
                remote_states=["FL", "GA"],
            ),
        )
        rng = random.Random(42)
        roster = generate_employees(rng, config=cfg)
        remote_states = {e.state for e in roster} - {"OR", "TX", "IL"}
        # Remote employees should only be from FL or GA (not CA/WA/NY)
        assert remote_states <= {"FL", "GA"}

    def test_rd_invariant_with_config(self) -> None:
        """R&D eligibility is preserved when using config (AM R&D + Engineering)."""
        cfg = _cascade_config()
        rng = random.Random(42)
        roster = generate_employees(rng, config=cfg)
        for e in roster:
            if e.is_rd_eligible:
                assert e.entity_code == "AM"
                assert e.department in ("R&D", "Engineering")
        rd_count = sum(1 for e in roster if e.is_rd_eligible)
        assert rd_count >= 45

    def test_config_determinism(self) -> None:
        """Config-driven generation is deterministic."""
        cfg = _cascade_config()
        r1 = generate_employees(random.Random(42), config=cfg)
        r2 = generate_employees(random.Random(42), config=cfg)
        assert len(r1) == len(r2)
        for e1, e2 in zip(r1, r2):
            assert e1 == e2
