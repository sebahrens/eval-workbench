"""Tests for the lease portfolio and ASC 842 computations."""

from __future__ import annotations

import datetime
import random

from generator.model.gl import Ledger
from generator.model.leases import (
    EscalationType,
    LeaseType,
    asc842_temp_difference,
    compute_lease_schedules,
    generate_leases,
    post_leases_to_gl,
)


def _leases():
    return generate_leases(random.Random(42))


class TestLeaseGeneration:
    def test_generates_15_leases(self):
        leases = _leases()
        assert len(leases) == 15

    def test_lease_ids_sequential(self):
        leases = _leases()
        ids = [ls.lease_id for ls in leases]
        assert ids == [f"LS-{i:03d}" for i in range(1, 16)]

    def test_all_entities_covered(self):
        leases = _leases()
        entities = {ls.entity_code for ls in leases}
        assert entities == {"CI", "PC", "AM", "DS"}

    def test_two_short_term_exempt(self):
        leases = _leases()
        short_term = [ls for ls in leases if ls.short_term_exempt]
        assert len(short_term) == 2
        for ls in short_term:
            assert ls.term_months <= 12

    def test_three_amendments(self):
        leases = _leases()
        amended = [ls for ls in leases if len(ls.amendments) > 0]
        assert len(amended) == 3

    def test_two_finance_leases(self):
        leases = _leases()
        finance = [ls for ls in leases if ls.lease_type == LeaseType.FINANCE]
        assert len(finance) == 2
        for ls in finance:
            assert ls.purchase_option is True

    def test_escalation_types_varied(self):
        leases = _leases()
        types = {ls.escalation_type for ls in leases}
        assert EscalationType.FIXED_PCT in types
        assert EscalationType.NONE in types

    def test_rou_and_liability_zero_for_short_term(self):
        leases = _leases()
        for ls in leases:
            if ls.short_term_exempt:
                assert ls.rou_asset_initial == 0
                assert ls.lease_liability_initial == 0

    def test_rou_equals_liability_at_commencement(self):
        leases = _leases()
        for ls in leases:
            if not ls.short_term_exempt:
                assert ls.rou_asset_initial == ls.lease_liability_initial
                assert ls.rou_asset_initial > 0

    def test_effective_rent_reflects_amendments(self):
        leases = _leases()
        for ls in leases:
            if ls.amendments:
                last_amend = ls.amendments[-1]
                if last_amend.new_monthly_rent is not None:
                    assert ls.effective_monthly_rent == last_amend.new_monthly_rent

    def test_end_date_after_commencement(self):
        leases = _leases()
        for ls in leases:
            assert ls.end_date > ls.commencement_date


class TestLeaseSchedules:
    def test_schedules_generated(self):
        leases = _leases()
        schedules = compute_lease_schedules(leases)
        assert len(schedules) > 0

    def test_all_active_leases_have_schedules(self):
        leases = _leases()
        schedules = compute_lease_schedules(leases)
        lease_ids_in_schedule = {r.lease_id for r in schedules}
        for ls in leases:
            for year in [2023, 2024, 2025]:
                year_start = datetime.date(year, 1, 1)
                year_end = datetime.date(year, 12, 31)
                if ls.commencement_date <= year_end and ls.end_date >= year_start:
                    assert ls.lease_id in lease_ids_in_schedule, (
                        f"{ls.lease_id} active in {year} but not in schedules"
                    )
                    break

    def test_cash_paid_positive(self):
        leases = _leases()
        schedules = compute_lease_schedules(leases)
        for row in schedules:
            assert row.cash_paid > 0

    def test_rou_and_liability_non_negative(self):
        leases = _leases()
        schedules = compute_lease_schedules(leases)
        for row in schedules:
            assert row.rou_asset_end >= 0, f"{row.lease_id} {row.year}: ROU negative"
            assert row.lease_liability_end >= 0, f"{row.lease_id} {row.year}: liability negative"


class TestASC842TempDifference:
    def test_temp_difference_exists(self):
        leases = _leases()
        schedules = compute_lease_schedules(leases)
        diffs = [asc842_temp_difference(schedules, y) for y in [2023, 2024, 2025]]
        assert any(d != 0 for d in diffs)

    def test_entity_filter(self):
        leases = _leases()
        schedules = compute_lease_schedules(leases)
        total = asc842_temp_difference(schedules, 2025)
        by_entity = sum(
            asc842_temp_difference(schedules, 2025, entity_code=e)
            for e in ["CI", "PC", "AM", "DS"]
        )
        assert total == by_entity


class TestGLPosting:
    def test_posts_without_error(self):
        leases = _leases()
        schedules = compute_lease_schedules(leases)
        ledger = Ledger()
        post_leases_to_gl(ledger, leases, schedules)
        assert len(ledger.entries) > 0

    def test_all_entries_balanced(self):
        leases = _leases()
        schedules = compute_lease_schedules(leases)
        ledger = Ledger()
        post_leases_to_gl(ledger, leases, schedules)
        for entry in ledger.entries:
            assert entry.is_balanced(), (
                f"Unbalanced: {entry.description} "
                f"DR={entry.total_debits()} CR={entry.total_credits()}"
            )


class TestDeterminism:
    def test_two_runs_identical(self):
        leases1 = generate_leases(random.Random(42))
        leases2 = generate_leases(random.Random(42))
        assert len(leases1) == len(leases2)
        for l1, l2 in zip(leases1, leases2):
            assert l1.lease_id == l2.lease_id
            assert l1.rou_asset_initial == l2.rou_asset_initial
            assert l1.lease_liability_initial == l2.lease_liability_initial
            assert l1.commencement_date == l2.commencement_date
