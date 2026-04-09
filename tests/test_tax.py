"""Tests for the tax provision model (TC-06)."""

from __future__ import annotations

import random
from decimal import Decimal

from generator.model.employees import generate_employees
from generator.model.leases import compute_lease_schedules, generate_leases
from generator.model.opex import generate_opex
from generator.model.ppe import generate_fixed_assets
from generator.model.revenue import generate_monthly_revenue
from generator.model.tax import (
    FEDERAL_RATE,
    STATE_RATE,
    compute_provisions_multi_year,
    compute_tax_provision,
    validate_provision,
)


def _setup():
    """Build all dependencies needed by the tax model."""
    employees = generate_employees(random.Random(42))
    rev_records = generate_monthly_revenue(random.Random(42))

    rev_by_ey: dict[tuple[str, int], Decimal] = {}
    for r in rev_records:
        key = (r.entity_code, r.year)
        rev_by_ey[key] = rev_by_ey.get(key, Decimal(0)) + r.revenue

    opex = generate_opex(random.Random(42), employees, rev_by_ey)
    assets = generate_fixed_assets(random.Random(42))
    leases = generate_leases(random.Random(42))
    schedules = compute_lease_schedules(leases)

    # Pre-tax book income per year (consolidated).
    # Approximate: revenue - COGS - opex - depreciation + other income.
    # For now use a simplified computation.
    pre_tax_by_year: dict[int, Decimal] = {}
    for year in [2023, 2024, 2025]:
        total_rev = sum(
            rev_by_ey.get((e, year), Decimal(0)) for e in ["PC", "AM", "DS"]
        )
        total_opex = sum(r.amount for r in opex if r.year == year)
        book_depr = sum(a.book_depr_for_year(year) for a in assets)
        # Gross margin ~35% of revenue → COGS ≈ 65% of revenue
        cogs = (total_rev * Decimal("0.65")).quantize(Decimal("1"))
        pre_tax = total_rev - cogs - total_opex - book_depr
        pre_tax_by_year[year] = pre_tax

    return opex, assets, schedules, pre_tax_by_year


class TestTaxProvision:
    def test_computes_all_years(self):
        opex, assets, schedules, pre_tax = _setup()
        provs = compute_provisions_multi_year(pre_tax, opex, assets, schedules)
        assert set(provs.keys()) == {2023, 2024, 2025}

    def test_current_plus_deferred_equals_total(self):
        opex, assets, schedules, pre_tax = _setup()
        provs = compute_provisions_multi_year(pre_tax, opex, assets, schedules)
        for year, prov in provs.items():
            assert prov.total_current + prov.total_deferred == prov.total_provision, (
                f"FY{year}: current + deferred != total"
            )

    def test_taxable_income_computation(self):
        opex, assets, schedules, pre_tax = _setup()
        provs = compute_provisions_multi_year(pre_tax, opex, assets, schedules)
        for year, prov in provs.items():
            expected = prov.pre_tax_book_income + prov.total_permanent + prov.total_temporary_change
            assert prov.taxable_income == expected, f"FY{year}: taxable income mismatch"

    def test_etr_within_reasonable_range(self):
        """ETR should be between 18% and 32% for a normal C-corp."""
        opex, assets, schedules, pre_tax = _setup()
        provs = compute_provisions_multi_year(pre_tax, opex, assets, schedules)
        for year, prov in provs.items():
            if prov.pre_tax_book_income > 0:
                assert Decimal("0.18") <= prov.effective_tax_rate <= Decimal("0.32"), (
                    f"FY{year}: ETR {prov.effective_tax_rate} out of range"
                )

    def test_permanent_differences_present(self):
        """Should have M&E, tax-exempt interest, stock comp, at minimum."""
        opex, assets, schedules, pre_tax = _setup()
        prov = compute_tax_provision(
            pre_tax[2025], opex, assets, schedules, 2025
        )
        descriptions = {p.description for p in prov.permanent_differences}
        assert any("Meals" in d for d in descriptions)
        assert any("Tax-exempt" in d or "tax-exempt" in d for d in descriptions)
        assert any("Stock" in d or "stock" in d for d in descriptions)

    def test_temporary_differences_present(self):
        """Should have depreciation, lease, warranty, inventory, bonuses, bad debt."""
        opex, assets, schedules, pre_tax = _setup()
        prov = compute_tax_provision(
            pre_tax[2025], opex, assets, schedules, 2025
        )
        descriptions = {t.description for t in prov.temporary_differences}
        assert any("epreciation" in d for d in descriptions)
        assert any("842" in d or "lease" in d.lower() for d in descriptions)
        assert any("arrant" in d for d in descriptions)
        assert any("nventory" in d for d in descriptions)
        assert any("onus" in d for d in descriptions)
        assert any("ad debt" in d.lower() or "Bad debt" in d for d in descriptions)

    def test_rd_credit_positive(self):
        """R&D credit should be positive (AM has significant R&D spend)."""
        opex, assets, schedules, pre_tax = _setup()
        prov = compute_tax_provision(
            pre_tax[2025], opex, assets, schedules, 2025
        )
        assert prov.rd_credit > 0

    def test_validate_provision_passes(self):
        """Validation should return no errors."""
        opex, assets, schedules, pre_tax = _setup()
        provs = compute_provisions_multi_year(pre_tax, opex, assets, schedules)
        for year, prov in provs.items():
            errors = validate_provision(prov)
            assert errors == [], f"FY{year}: {errors}"

    def test_dta_dtl_rollforward(self):
        """DTA/DTL should roll forward consistently year over year."""
        opex, assets, schedules, pre_tax = _setup()
        provs = compute_provisions_multi_year(pre_tax, opex, assets, schedules)

        # FY2024's prior should equal FY2023's current
        assert provs[2024].prior_dta_total == provs[2023].current_dta_total
        assert provs[2024].prior_dtl_total == provs[2023].current_dtl_total
        # FY2025's prior should equal FY2024's current
        assert provs[2025].prior_dta_total == provs[2024].current_dta_total
        assert provs[2025].prior_dtl_total == provs[2024].current_dtl_total

    def test_state_rate_and_federal_rate(self):
        assert FEDERAL_RATE == Decimal("0.21")
        assert STATE_RATE == Decimal("0.062")

    def test_deterministic(self):
        """Two runs should produce identical provisions."""
        opex1, assets1, sch1, pre1 = _setup()
        provs1 = compute_provisions_multi_year(pre1, opex1, assets1, sch1)

        opex2, assets2, sch2, pre2 = _setup()
        provs2 = compute_provisions_multi_year(pre2, opex2, assets2, sch2)

        for year in [2023, 2024, 2025]:
            assert provs1[year].total_provision == provs2[year].total_provision
            assert provs1[year].effective_tax_rate == provs2[year].effective_tax_rate
