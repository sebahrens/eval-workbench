"""Tests for generator.model.revenue — revenue & COGS generation."""

import random
from decimal import Decimal

from generator.model.gl import Ledger
from generator.model.revenue import (
    PRODUCT_LINES,
    MonthlyRevenue,
    generate_monthly_revenue,
    post_revenue_to_gl,
    validate_consolidated_growth,
    validate_product_line_growth,
)


def _make_records() -> list[MonthlyRevenue]:
    rng = random.Random(42)
    return generate_monthly_revenue(rng)


class TestRevenueGeneration:
    """Core revenue generation tests."""

    def test_record_count(self) -> None:
        """6 product lines × 3 years × 12 months = 216 records."""
        records = _make_records()
        assert len(records) == 216

    def test_all_product_lines_present(self) -> None:
        records = _make_records()
        names = {r.product_line for r in records}
        expected = {pl.name for pl in PRODUCT_LINES}
        assert names == expected

    def test_all_entities_present(self) -> None:
        records = _make_records()
        codes = {r.entity_code for r in records}
        assert codes == {"PC", "AM", "DS"}

    def test_all_months_covered(self) -> None:
        records = _make_records()
        for year in (2023, 2024, 2025):
            months = {r.month for r in records if r.year == year}
            assert months == set(range(1, 13)), f"Missing months in {year}"

    def test_revenue_positive(self) -> None:
        records = _make_records()
        for r in records:
            assert r.revenue > 0, f"Non-positive revenue: {r}"

    def test_cogs_positive(self) -> None:
        records = _make_records()
        for r in records:
            assert r.cogs > 0, f"Non-positive COGS: {r}"

    def test_cogs_less_than_revenue(self) -> None:
        records = _make_records()
        for r in records:
            assert r.cogs < r.revenue, f"COGS >= revenue: {r}"

    def test_sorted_deterministically(self) -> None:
        records = _make_records()
        keys = [(r.year, r.month, r.entity_code, r.product_line) for r in records]
        assert keys == sorted(keys)


class TestEntityRevenueTotals:
    """Entity-level FY2025 revenue targets."""

    def test_precision_components_fy25(self) -> None:
        records = _make_records()
        total = sum(r.revenue for r in records if r.entity_code == "PC" and r.year == 2025)
        assert abs(total - Decimal(95_000_000)) < 1, f"PC FY25: {total}"

    def test_advanced_materials_fy25(self) -> None:
        records = _make_records()
        total = sum(r.revenue for r in records if r.entity_code == "AM" and r.year == 2025)
        assert abs(total - Decimal(65_000_000)) < 1, f"AM FY25: {total}"

    def test_distribution_services_fy25(self) -> None:
        records = _make_records()
        total = sum(r.revenue for r in records if r.entity_code == "DS" and r.year == 2025)
        assert abs(total - Decimal(40_000_000)) < 1, f"DS FY25: {total}"

    def test_consolidated_fy25(self) -> None:
        records = _make_records()
        total = sum(r.revenue for r in records if r.year == 2025)
        assert abs(total - Decimal(200_000_000)) < 1, f"Consolidated FY25: {total}"


class TestGrowthRates:
    """Consolidated and product-line growth rate validation."""

    def test_consolidated_fy24_growth_approx_6pct(self) -> None:
        records = _make_records()
        growth = validate_consolidated_growth(records)
        fy24 = growth["FY2024_growth"]
        assert abs(fy24 - Decimal("0.06")) < Decimal("0.001"), f"FY24 growth: {fy24}"

    def test_consolidated_fy25_growth_9_2pct(self) -> None:
        """TC-03 gold value: consolidated FY25 growth = 9.2%."""
        records = _make_records()
        growth = validate_consolidated_growth(records)
        fy25 = growth["FY2025_growth"]
        # Must be within 0.1% of 9.2% to match gold standard
        assert abs(fy25 - Decimal("0.092")) < Decimal("0.001"), f"FY25 growth: {fy25}"

    def test_specialty_coatings_declines_4pct(self) -> None:
        """TC-03: one product line declined 4%."""
        records = _make_records()
        pl_growth = validate_product_line_growth(records)
        sc_growth = pl_growth["Specialty Coatings"]
        assert abs(sc_growth - Decimal("-0.04")) < Decimal("0.001"), f"Specialty Coatings growth: {sc_growth}"

    def test_two_growth_drivers(self) -> None:
        """TC-03: two product lines drove the growth (well above average)."""
        records = _make_records()
        pl_growth = validate_product_line_growth(records)
        above_avg = [name for name, g in pl_growth.items() if g > Decimal("0.15")]
        assert len(above_avg) >= 1, f"Expected growth drivers, got: {pl_growth}"
        # Advanced Composites should be the clear driver
        assert "Advanced Composites" in above_avg


class TestGrossMargins:
    """COGS correlates at entity gross margins per §2.4 rule 2."""

    def _entity_margin(self, records: list[MonthlyRevenue], entity_code: str) -> Decimal:
        rev = sum(r.revenue for r in records if r.entity_code == entity_code and r.year == 2025)
        cogs = sum(r.cogs for r in records if r.entity_code == entity_code and r.year == 2025)
        return (rev - cogs) / rev

    def test_pc_margin_35pct(self) -> None:
        records = _make_records()
        margin = self._entity_margin(records, "PC")
        assert abs(margin - Decimal("0.35")) < Decimal("0.01"), f"PC margin: {margin}"

    def test_am_margin_52pct(self) -> None:
        records = _make_records()
        margin = self._entity_margin(records, "AM")
        assert abs(margin - Decimal("0.52")) < Decimal("0.01"), f"AM margin: {margin}"

    def test_ds_margin_18pct(self) -> None:
        records = _make_records()
        margin = self._entity_margin(records, "DS")
        assert abs(margin - Decimal("0.18")) < Decimal("0.01"), f"DS margin: {margin}"


class TestSeasonalPattern:
    """Q4 heavy, Q1 dip — §2.4 rule 1."""

    def test_q4_greater_than_q1(self) -> None:
        records = _make_records()
        fy25 = [r for r in records if r.year == 2025]
        q1_rev = sum(r.revenue for r in fy25 if r.month <= 3)
        q4_rev = sum(r.revenue for r in fy25 if r.month >= 10)
        assert q4_rev > q1_rev, f"Q4 ({q4_rev}) should exceed Q1 ({q1_rev})"

    def test_q4_is_largest_quarter(self) -> None:
        records = _make_records()
        fy25 = [r for r in records if r.year == 2025]
        quarters = {}
        for r in fy25:
            q = (r.month - 1) // 3 + 1
            quarters[q] = quarters.get(q, Decimal(0)) + r.revenue
        assert quarters[4] == max(quarters.values())


class TestGLPosting:
    """Revenue and COGS post correctly to the GL."""

    def test_posts_without_error(self) -> None:
        records = _make_records()
        ledger = Ledger()
        post_revenue_to_gl(ledger, records)
        # Should have 2 JEs per record (revenue + COGS)
        assert len(ledger.entries) == len(records) * 2

    def test_all_entries_balanced(self) -> None:
        records = _make_records()
        ledger = Ledger()
        post_revenue_to_gl(ledger, records)
        for entry in ledger.entries:
            assert entry.is_balanced(), f"Unbalanced: {entry.description}"

    def test_ar_balance_positive(self) -> None:
        """A/R should have a net debit balance (asset)."""
        records = _make_records()
        ledger = Ledger()
        post_revenue_to_gl(ledger, records)
        for entity_code in ("PC", "AM", "DS"):
            balances = ledger.balance_by_account(entity_code)
            ar_bal = balances.get("1100", Decimal(0))
            assert ar_bal > 0, f"{entity_code} A/R should be positive: {ar_bal}"


class TestDeterminism:
    """Same seed → same output."""

    def test_two_runs_identical(self) -> None:
        r1 = generate_monthly_revenue(random.Random(42))
        r2 = generate_monthly_revenue(random.Random(42))
        assert len(r1) == len(r2)
        for a, b in zip(r1, r2):
            assert a.revenue == b.revenue
            assert a.cogs == b.cogs
            assert a.year == b.year
            assert a.month == b.month
            assert a.entity_code == b.entity_code
            assert a.product_line == b.product_line
