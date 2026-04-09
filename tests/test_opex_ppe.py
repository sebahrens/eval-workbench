"""Tests for opex and ppe model modules."""

from __future__ import annotations

import random
from decimal import Decimal

from generator.model.employees import generate_employees
from generator.model.gl import Ledger
from generator.model.opex import (
    generate_opex,
    post_opex_to_gl,
    validate_rd_spend,
)
from generator.model.ppe import (
    FixedAsset,
    compute_depreciation_summary,
    cumulative_book_tax_difference,
    generate_fixed_assets,
    post_asset_acquisitions_to_gl,
    post_depreciation_to_gl,
)
from generator.model.revenue import generate_monthly_revenue


def _setup():
    """Create shared test fixtures."""
    rng = random.Random(42)
    employees = generate_employees(random.Random(42))
    rev_rng = random.Random(42)
    revenue_records = generate_monthly_revenue(rev_rng)

    # Build revenue_by_entity_year lookup
    rev_by_ey: dict[tuple[str, int], Decimal] = {}
    for r in revenue_records:
        key = (r.entity_code, r.year)
        rev_by_ey[key] = rev_by_ey.get(key, Decimal(0)) + r.revenue

    return rng, employees, rev_by_ey


class TestOpex:
    def test_generates_records(self):
        rng, employees, rev_by_ey = _setup()
        records = generate_opex(rng, employees, rev_by_ey)
        assert len(records) > 0

    def test_all_entities_covered(self):
        rng, employees, rev_by_ey = _setup()
        records = generate_opex(rng, employees, rev_by_ey)
        entities = {r.entity_code for r in records}
        assert entities == {"PC", "AM", "DS"}

    def test_all_years_covered(self):
        rng, employees, rev_by_ey = _setup()
        records = generate_opex(rng, employees, rev_by_ey)
        years = {r.year for r in records}
        assert years == {2023, 2024, 2025}

    def test_rd_spend_approximately_12_pct(self):
        """Advanced Materials R&D expense should be ≈12% of its revenue."""
        rng, employees, rev_by_ey = _setup()
        records = generate_opex(rng, employees, rev_by_ey)
        am_rev_2025 = rev_by_ey[("AM", 2025)]
        rd_pct = validate_rd_spend(records, am_rev_2025, year=2025)
        # Should be within 1% of 12% target
        assert Decimal("0.11") <= rd_pct <= Decimal("0.13"), f"R&D spend was {rd_pct:.4f}"

    def test_posts_to_gl_balanced(self):
        rng, employees, rev_by_ey = _setup()
        records = generate_opex(rng, employees, rev_by_ey)
        ledger = Ledger()
        post_opex_to_gl(ledger, records)
        # Every posted entry must be balanced (Ledger.post enforces this)
        assert len(ledger.entries) > 0

    def test_deterministic(self):
        """Two runs with the same seed produce identical records."""
        rng1, emp1, rev1 = _setup()
        records1 = generate_opex(rng1, emp1, rev1)

        rng2, emp2, rev2 = _setup()
        records2 = generate_opex(rng2, emp2, rev2)

        assert len(records1) == len(records2)
        for a, b in zip(records1, records2):
            assert a.year == b.year
            assert a.month == b.month
            assert a.entity_code == b.entity_code
            assert a.category == b.category
            assert a.amount == b.amount


class TestPPE:
    def test_generates_assets(self):
        rng = random.Random(42)
        assets = generate_fixed_assets(rng)
        assert len(assets) > 0

    def test_all_entities_have_assets(self):
        rng = random.Random(42)
        assets = generate_fixed_assets(rng)
        entities = {a.entity_code for a in assets}
        assert entities == {"PC", "AM", "DS"}

    def test_asset_ids_unique(self):
        rng = random.Random(42)
        assets = generate_fixed_assets(rng)
        ids = [a.asset_id for a in assets]
        assert len(ids) == len(set(ids))

    def test_book_depr_annual(self):
        """Straight-line depreciation calculation."""
        asset = FixedAsset(
            asset_id="TEST-001",
            entity_code="PC",
            description="Test Machine",
            asset_class="Machinery",
            acquisition_date=__import__("datetime").date(2023, 1, 1),
            cost=Decimal("100000"),
            salvage_value=Decimal("5000"),
            book_life_years=10,
            macrs_life_years=7,
            asset_account="1420",
            accum_depr_account="1510",
            depr_expense_account="6210",
        )
        annual = asset.book_depr_annual()
        assert annual == Decimal("9500")  # (100000 - 5000) / 10

    def test_macrs_sums_to_cost(self):
        """MACRS depreciation over full life should equal full cost."""
        asset = FixedAsset(
            asset_id="TEST-002",
            entity_code="AM",
            description="Test Equipment",
            asset_class="R&D Equipment",
            acquisition_date=__import__("datetime").date(2020, 1, 1),
            cost=Decimal("100000"),
            salvage_value=Decimal("5000"),
            book_life_years=5,
            macrs_life_years=5,
            asset_account="1420",
            accum_depr_account="1510",
            depr_expense_account="6330",
        )
        # 5-year MACRS has 6 years of percentages (half-year convention)
        total = sum(asset.tax_depr_for_year(2020 + i) for i in range(6))
        assert total == Decimal("100000")

    def test_book_tax_difference_exists(self):
        """Book-tax depreciation difference should be non-zero."""
        rng = random.Random(42)
        assets = generate_fixed_assets(rng)
        summaries = compute_depreciation_summary(assets)
        # At least one year should have a non-zero difference
        diffs = [s.book_tax_difference for s in summaries]
        assert any(d != 0 for d in diffs)

    def test_cumulative_difference_queryable(self):
        """cumulative_book_tax_difference should return a value for TC-06."""
        rng = random.Random(42)
        assets = generate_fixed_assets(rng)
        diff = cumulative_book_tax_difference(assets, through_year=2025)
        # Should be non-zero for a realistic asset base
        assert diff != 0

    def test_posts_depreciation_balanced(self):
        rng = random.Random(42)
        assets = generate_fixed_assets(rng)
        ledger = Ledger()
        post_asset_acquisitions_to_gl(ledger, assets)
        post_depreciation_to_gl(ledger, assets)
        assert len(ledger.entries) > 0

    def test_deterministic(self):
        assets1 = generate_fixed_assets(random.Random(42))
        assets2 = generate_fixed_assets(random.Random(42))
        assert len(assets1) == len(assets2)
        for a, b in zip(assets1, assets2):
            assert a.asset_id == b.asset_id
            assert a.cost == b.cost
            assert a.acquisition_date == b.acquisition_date
