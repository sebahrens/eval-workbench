"""Tests for generator.model.ar and generator.model.ap — AR/AP subledgers."""

from __future__ import annotations

import datetime
import random
from decimal import Decimal

from generator.model.ap import (
    VENDORS,
    generate_ap_aging,
    generate_ap_flows,
    post_ap_to_gl,
    validate_ap_equals_gl,
)
from generator.model.ar import (
    CUSTOMERS,
    generate_allowance,
    generate_ar_aging,
    generate_collections,
    post_collections_to_gl,
    validate_ar_equals_gl,
)
from generator.model.gl import Ledger
from generator.model.revenue import (
    generate_monthly_revenue,
    post_revenue_to_gl,
)


def _make_revenue():
    rng = random.Random(42)
    return generate_monthly_revenue(rng)


# ── Customer/vendor definition tests ────────────────────────────────────────


class TestCustomerDefinitions:
    def test_customer_count(self):
        assert len(CUSTOMERS) == 26

    def test_shares_sum_to_one(self):
        for ec in ("PC", "AM", "DS"):
            total = sum(c.revenue_share for c in CUSTOMERS if c.entity_code == ec)
            assert total == Decimal("1.000"), f"{ec} shares sum to {total}"

    def test_top_customer_is_18_pct(self):
        """Acme Manufacturing = 37.9% of PC ($95M) ≈ $36M ≈ 18% of $200M."""
        top = max(CUSTOMERS, key=lambda c: c.revenue_share)
        assert top.name == "Acme Manufacturing Corp"
        # 37.9% of $95M = $36.005M.  $36.005M / $200M = 18.0%
        consol_share = top.revenue_share * Decimal("95000000") / Decimal("200000000")
        assert Decimal("0.17") < consol_share < Decimal("0.19")

    def test_all_dso_positive(self):
        for c in CUSTOMERS:
            assert c.dso > 0


class TestVendorDefinitions:
    def test_vendor_count(self):
        assert len(VENDORS) == 26

    def test_shares_sum_to_one(self):
        for ec in ("PC", "AM", "DS"):
            total = sum(v.purchase_share for v in VENDORS if v.entity_code == ec)
            assert total == Decimal("1.000"), f"{ec} shares sum to {total}"

    def test_all_payment_terms_positive(self):
        for v in VENDORS:
            assert v.payment_terms > 0


# ── AR aging tests ──────────────────────────────────────────────────────────


class TestARaging:
    def test_entry_count(self):
        records = _make_revenue()
        aging = generate_ar_aging(records)
        assert len(aging) == len(CUSTOMERS)

    def test_all_entities_present(self):
        records = _make_revenue()
        aging = generate_ar_aging(records)
        entities = {e.entity_code for e in aging}
        assert entities == {"PC", "AM", "DS"}

    def test_total_positive(self):
        records = _make_revenue()
        aging = generate_ar_aging(records)
        total = sum(e.total for e in aging)
        assert total > 0

    def test_consolidated_ar_reasonable(self):
        """AR should be roughly 40–55 DSO worth of $200M annual revenue."""
        records = _make_revenue()
        aging = generate_ar_aging(records)
        total = sum(e.total for e in aging)
        daily_rev = Decimal("200000000") / Decimal("365")
        implied_dso = total / daily_rev
        assert Decimal("25") < implied_dso < Decimal("65"), (
            f"Implied DSO = {implied_dso:.1f} — outside realistic range"
        )

    def test_high_dso_customer_has_aged_buckets(self):
        records = _make_revenue()
        aging = generate_ar_aging(records)
        westlake = next(e for e in aging if e.customer_id == "CUST-010")
        assert westlake.dso == 120
        # With DSO=120, should have amounts in 60+ or 90+ buckets
        assert westlake.days_60 > 0 or westlake.days_90 > 0

    def test_low_dso_customer_only_current(self):
        records = _make_revenue()
        aging = generate_ar_aging(records)
        cascade = next(e for e in aging if e.customer_id == "CUST-005")
        assert cascade.dso == 30
        # DSO=30 should have nothing beyond current
        assert cascade.days_60 == 0
        assert cascade.days_90 == 0
        assert cascade.days_120_plus == 0

    def test_deterministic(self):
        aging1 = generate_ar_aging(_make_revenue())
        aging2 = generate_ar_aging(_make_revenue())
        for a, b in zip(aging1, aging2):
            assert a.total == b.total


# ── Collections tests ───────────────────────────────────────────────────────


class TestCollections:
    def test_collections_generated(self):
        records = _make_revenue()
        colls = generate_collections(records)
        assert len(colls) > 0

    def test_all_amounts_positive(self):
        records = _make_revenue()
        colls = generate_collections(records)
        for c in colls:
            assert c.amount > 0

    def test_total_collections_less_than_total_revenue(self):
        records = _make_revenue()
        colls = generate_collections(records)
        total_colls = sum(c.amount for c in colls)
        total_rev = sum(
            r.revenue.quantize(Decimal("1")) for r in records
        )
        assert total_colls < total_rev  # some AR still outstanding


# ── AR ↔ GL reconciliation (acceptance criterion) ───────────────────────────


class TestARGLReconciliation:
    def test_ar_aging_equals_gl_balance(self):
        """Core acceptance criterion: AR aging total = GL 1100 balance."""
        records = _make_revenue()
        ledger = Ledger()

        # Post revenue (creates 1100 debits)
        post_revenue_to_gl(ledger, records)

        # Post collections (creates 1100 credits)
        colls = generate_collections(records)
        post_collections_to_gl(ledger, colls)

        # Compute aging
        aging = generate_ar_aging(records, year=2025)

        for ec in ("PC", "AM", "DS"):
            aging_total, gl_balance, match = validate_ar_equals_gl(
                aging, ledger, ec, as_of_date=datetime.date(2025, 12, 31)
            )
            assert match, (
                f"{ec}: AR aging total={aging_total} ≠ GL 1100={gl_balance}"
            )


# ── Allowance tests ─────────────────────────────────────────────────────────


class TestAllowance:
    def test_allowance_generated(self):
        records = _make_revenue()
        allowance = generate_allowance(records)
        assert len(allowance) == 9  # 3 entities × 3 years

    def test_ending_balance_positive(self):
        records = _make_revenue()
        allowance = generate_allowance(records)
        for a in allowance:
            assert a.ending_balance >= 0

    def test_rollforward_consistent(self):
        """ending = beginning + provision for each year."""
        records = _make_revenue()
        allowance = generate_allowance(records)
        for a in allowance:
            assert a.ending_balance == a.beginning_balance + a.provision


# ── AP aging tests ──────────────────────────────────────────────────────────


class TestAPaging:
    def test_entry_count(self):
        records = _make_revenue()
        aging = generate_ap_aging(records)
        assert len(aging) == len(VENDORS)

    def test_total_positive(self):
        records = _make_revenue()
        aging = generate_ap_aging(records)
        total = sum(e.total for e in aging)
        assert total > 0

    def test_consolidated_ap_reasonable(self):
        """AP should be roughly 20-50 DPO worth of annual purchases."""
        records = _make_revenue()
        aging = generate_ap_aging(records)
        total = sum(e.total for e in aging)
        # Annual purchases ≈ $130M (weighted avg of purchase rates × revenue)
        daily_purchases = Decimal("130000000") / Decimal("365")
        implied_dpo = total / daily_purchases
        assert Decimal("10") < implied_dpo < Decimal("60"), (
            f"Implied DPO = {implied_dpo:.1f} — outside realistic range"
        )

    def test_short_terms_vendor_mostly_current(self):
        records = _make_revenue()
        aging = generate_ap_aging(records)
        heritage = next(e for e in aging if e.vendor_id == "VEND-009")
        assert heritage.payment_terms == 15
        # Net 15 should have nothing beyond current
        assert heritage.days_30 == 0
        assert heritage.days_60 == 0
        assert heritage.days_90_plus == 0

    def test_deterministic(self):
        aging1 = generate_ap_aging(_make_revenue())
        aging2 = generate_ap_aging(_make_revenue())
        for a, b in zip(aging1, aging2):
            assert a.total == b.total


# ── AP flows tests ──────────────────────────────────────────────────────────


class TestAPFlows:
    def test_flows_generated(self):
        records = _make_revenue()
        flows = generate_ap_flows(records)
        assert len(flows) > 0

    def test_purchases_positive(self):
        records = _make_revenue()
        flows = generate_ap_flows(records)
        for f in flows:
            assert f.purchases >= 0


# ── AP ↔ GL reconciliation (acceptance criterion) ──────────────────────────


class TestAPGLReconciliation:
    def test_ap_aging_equals_gl_balance(self):
        """Core acceptance criterion: AP aging total = GL 2010 balance."""
        records = _make_revenue()
        ledger = Ledger()

        # Post AP flows (purchases and payments)
        flows = generate_ap_flows(records)
        post_ap_to_gl(ledger, flows)

        # Compute aging
        aging = generate_ap_aging(records, year=2025)

        for ec in ("PC", "AM", "DS"):
            aging_total, gl_balance, match = validate_ap_equals_gl(
                aging, ledger, ec, as_of_date=datetime.date(2025, 12, 31)
            )
            assert match, (
                f"{ec}: AP aging total={aging_total} ≠ GL 2010={gl_balance}"
            )
