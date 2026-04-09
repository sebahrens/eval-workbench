"""Tests for K-1 partnership investment data (TC-07)."""

from __future__ import annotations

from decimal import Decimal

from generator.model.k1 import (
    K1LayoutType,
    consolidated_totals,
    generate_k1_investments,
)


def _investments():
    return generate_k1_investments()


class TestK1Generation:
    def test_generates_8_investments(self):
        investments = _investments()
        assert len(investments) == 8

    def test_k1_ids_sequential(self):
        investments = _investments()
        ids = [k.k1_id for k in investments]
        assert ids == [f"K1-{i:03d}" for i in range(1, 9)]

    def test_all_tax_year_2025(self):
        for inv in _investments():
            assert inv.tax_year == 2025

    def test_amount_range(self):
        """Amounts range from $5,000 to $2.3M per prompt.md."""
        investments = _investments()
        # Find the min and max ordinary income across all K-1s
        ordinary = [inv.box_1_ordinary_income for inv in investments
                    if inv.box_1_ordinary_income is not None]
        assert min(ordinary) == Decimal("5000")
        assert max(ordinary) == Decimal("2300000")

    def test_three_system_clean_five_varying(self):
        investments = _investments()
        clean = [k for k in investments if k.layout_type == K1LayoutType.SYSTEM_CLEAN]
        varying = [k for k in investments if k.layout_type == K1LayoutType.VARYING]
        assert len(clean) == 3
        assert len(varying) == 5

    def test_one_amended_k1(self):
        investments = _investments()
        amended = [k for k in investments if k.is_amended]
        assert len(amended) == 1

    def test_amended_k1_details(self):
        """The amended K-1 changed ordinary income from $340K to $285K
        and added a $55K guaranteed payment (prompt.md §4 TC-07)."""
        investments = _investments()
        amended = [k for k in investments if k.is_amended][0]

        assert amended.box_1_ordinary_income == Decimal("285000")
        assert amended.box_4c_total_guaranteed_payments == Decimal("55000")

        # Check amendment records
        assert len(amended.amendments) == 2
        income_amend = [a for a in amended.amendments
                        if a.field_changed == "box_1_ordinary_income"][0]
        assert income_amend.original_value == Decimal("340000")
        assert income_amend.amended_value == Decimal("285000")

        gp_amend = [a for a in amended.amendments
                     if a.field_changed == "box_4c_total_guaranteed_payments"][0]
        assert gp_amend.original_value == Decimal("0")
        assert gp_amend.amended_value == Decimal("55000")

    def test_section_199a_present_on_most(self):
        """Section 199A amounts present but N/A to C-corp."""
        investments = _investments()
        with_199a = [k for k in investments if k.section_199a_qbi is not None]
        # Most K-1s have 199A; at least the venture fund (K1-003) doesn't
        assert len(with_199a) >= 6

    def test_boxes_1_through_13_plus_box_20(self):
        """Each K-1 should have at least some box data."""
        for inv in _investments():
            assert inv.total_income > 0 or inv.total_deductions > 0

    def test_multiple_entities_represented(self):
        investments = _investments()
        entities = {k.entity_code for k in investments}
        assert len(entities) >= 3  # CI, PC, AM, DS


class TestK1Consolidation:
    def test_consolidated_totals_not_empty(self):
        totals = consolidated_totals(_investments())
        assert len(totals) > 0

    def test_consolidated_ordinary_income(self):
        totals = consolidated_totals(_investments())
        # Sum: 2300000 + 78000 + 285000 + 520000 + 145000 + 5000 + 890000 = 4223000
        assert totals["Box 1 - Ordinary business income"] == Decimal("4223000")

    def test_consolidated_guaranteed_payments(self):
        totals = consolidated_totals(_investments())
        # Sum: 55000 (K1-004) + 25000 (K1-006) + 35000 (K1-008) = 115000
        assert totals["Box 4c - Guaranteed payments"] == Decimal("115000")

    def test_consolidated_lt_capital_gain(self):
        totals = consolidated_totals(_investments())
        # Sum: 156000 + 92000 + 445000 + 38000 + 175000 = 906000
        assert totals["Box 9a - Net LT capital gain"] == Decimal("906000")

    def test_section_199a_totals_present(self):
        """199A amounts should be consolidated even though N/A for C-corp."""
        totals = consolidated_totals(_investments())
        assert "Section 199A - QBI (N/A to C-corp)" in totals

    def test_deterministic(self):
        """Two calls produce identical results."""
        a = generate_k1_investments()
        b = generate_k1_investments()
        assert len(a) == len(b)
        for ia, ib in zip(a, b):
            assert ia == ib
