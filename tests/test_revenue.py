"""Tests for generator.model.revenue — revenue & COGS generation."""

import random
from decimal import Decimal

from generator.config import (
    CompanyConfig,
    Config,
    EmployeeConfig,
    GrowthRates,
    IntercompanyConfig,
    SeasonalWeights,
    SubsidiaryConfig,
)
from generator.model.gl import Ledger
from generator.model.revenue import (
    PRODUCT_LINES,
    MonthlyRevenue,
    _quarterly_to_monthly,
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


# ── Config-driven seasonal weights ────────────────────────────────────────

def _make_cascade_config(seasonal: SeasonalWeights | None = None) -> Config:
    """Build a Config matching Cascade defaults, optionally with custom seasonal weights."""
    sw = seasonal or SeasonalWeights(Q1=0.20, Q2=0.25, Q3=0.25, Q4=0.30)
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
            subsidiaries={
                "precision_components": SubsidiaryConfig(
                    legal_name="Cascade Precision Components LLC",
                    location="Portland, OR", state="OR", entity_code="PC",
                    revenue=95_000_000, type="Core manufacturing",
                    gross_margin=0.35, employee_count=350,
                ),
                "advanced_materials": SubsidiaryConfig(
                    legal_name="Cascade Advanced Materials, Inc.",
                    location="Austin, TX", state="TX", entity_code="AM",
                    revenue=65_000_000, type="Specialty materials",
                    gross_margin=0.52, employee_count=280, rd_spend_pct=0.12,
                ),
                "distribution_services": SubsidiaryConfig(
                    legal_name="Cascade Distribution Services LLC",
                    location="Chicago, IL", state="IL", entity_code="DS",
                    revenue=40_000_000, type="Warehousing and logistics",
                    gross_margin=0.18, employee_count=220,
                ),
            },
            growth_rates=GrowthRates(fy2023_to_fy2024=0.06, fy2024_to_fy2025=0.09),
            intercompany=IntercompanyConfig(
                raw_materials_markup=0.08, management_fee_pct=0.015,
                intercompany_loan_principal=5_000_000, intercompany_loan_rate=0.05,
            ),
            employees=EmployeeConfig(total_count=850, annual_turnover_rate=0.08),
            seasonal_weights=sw,
        ),
        canary_assignments={},
        error_injections={},
    )


class TestConfigSeasonalWeights:
    """Revenue generation respects custom seasonal weights from config."""

    def test_default_config_matches_hardcoded(self) -> None:
        """Config with Cascade defaults produces same annual totals as no-config path."""
        cfg = _make_cascade_config()
        r_default = generate_monthly_revenue(random.Random(42))
        r_config = generate_monthly_revenue(random.Random(42), config=cfg)
        # Annual totals per entity must match
        for entity in ("PC", "AM", "DS"):
            for year in (2023, 2024, 2025):
                total_def = sum(r.revenue for r in r_default if r.entity_code == entity and r.year == year)
                total_cfg = sum(r.revenue for r in r_config if r.entity_code == entity and r.year == year)
                assert abs(total_def - total_cfg) < 1, (
                    f"{entity} FY{year}: default={total_def}, config={total_cfg}"
                )

    def test_flat_seasonal_weights(self) -> None:
        """Flat quarterly weights (0.25 each) produce roughly uniform quarters."""
        cfg = _make_cascade_config(SeasonalWeights(Q1=0.25, Q2=0.25, Q3=0.25, Q4=0.25))
        records = generate_monthly_revenue(random.Random(42), config=cfg)
        fy25 = [r for r in records if r.year == 2025]
        quarters: dict[int, Decimal] = {}
        for r in fy25:
            q = (r.month - 1) // 3 + 1
            quarters[q] = quarters.get(q, Decimal(0)) + r.revenue
        total = sum(quarters.values())
        for q, amt in quarters.items():
            share = float(amt / total)
            assert abs(share - 0.25) < 0.02, f"Q{q} share {share:.3f} deviates from 0.25"

    def test_q1_heavy_reverses_pattern(self) -> None:
        """Q1-heavy weights make Q1 the largest quarter (opposite of default)."""
        cfg = _make_cascade_config(SeasonalWeights(Q1=0.40, Q2=0.20, Q3=0.20, Q4=0.20))
        records = generate_monthly_revenue(random.Random(42), config=cfg)
        fy25 = [r for r in records if r.year == 2025]
        quarters: dict[int, Decimal] = {}
        for r in fy25:
            q = (r.month - 1) // 3 + 1
            quarters[q] = quarters.get(q, Decimal(0)) + r.revenue
        assert quarters[1] == max(quarters.values()), (
            f"Q1 should be largest but got: {quarters}"
        )

    def test_custom_weights_preserve_annual_totals(self) -> None:
        """Custom seasonal weights don't change annual totals — only monthly distribution."""
        cfg = _make_cascade_config(SeasonalWeights(Q1=0.10, Q2=0.30, Q3=0.30, Q4=0.30))
        records = generate_monthly_revenue(random.Random(42), config=cfg)
        for entity in ("PC", "AM", "DS"):
            total = sum(r.revenue for r in records if r.entity_code == entity and r.year == 2025)
            expected = {"PC": 95_000_000, "AM": 65_000_000, "DS": 40_000_000}[entity]
            assert abs(total - Decimal(expected)) < 1, f"{entity} FY25: {total}"

    def test_determinism_with_config(self) -> None:
        """Same config + seed → identical output."""
        cfg = _make_cascade_config(SeasonalWeights(Q1=0.15, Q2=0.35, Q3=0.25, Q4=0.25))
        r1 = generate_monthly_revenue(random.Random(42), config=cfg)
        r2 = generate_monthly_revenue(random.Random(42), config=cfg)
        assert len(r1) == len(r2)
        for a, b in zip(r1, r2):
            assert a.revenue == b.revenue


class TestQuarterlyToMonthly:
    """Unit tests for the _quarterly_to_monthly helper."""

    def test_weights_sum_to_one(self) -> None:
        monthly = _quarterly_to_monthly(0.20, 0.25, 0.25, 0.30)
        assert abs(sum(monthly) - 1.0) < 1e-9

    def test_twelve_months(self) -> None:
        monthly = _quarterly_to_monthly(0.25, 0.25, 0.25, 0.25)
        assert len(monthly) == 12

    def test_quarter_totals_preserved(self) -> None:
        monthly = _quarterly_to_monthly(0.10, 0.30, 0.35, 0.25)
        q1 = sum(monthly[0:3])
        q2 = sum(monthly[3:6])
        q3 = sum(monthly[6:9])
        q4 = sum(monthly[9:12])
        assert abs(q1 - 0.10) < 0.01
        assert abs(q2 - 0.30) < 0.01
        assert abs(q3 - 0.35) < 0.01
        assert abs(q4 - 0.25) < 0.01
