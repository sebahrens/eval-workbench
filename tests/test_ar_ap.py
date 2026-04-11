"""Tests for generator.model.ar and generator.model.ap — AR/AP subledgers."""

from __future__ import annotations

import datetime
import random
from decimal import Decimal

from generator.model.ap import (
    VENDORS,
    compute_vendor_annual_purchases,
    generate_ap_aging,
    generate_ap_flows,
    post_ap_to_gl,
    validate_ap_equals_gl,
)
from generator.model.ap_ledger import generate_ap_ledger
from generator.model.ar import (
    CUSTOMERS,
    generate_allowance,
    generate_ar_aging,
    generate_collections,
    generate_invoices,
    generate_receipts,
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


# ── Invoice lifecycle tests ───────────────────────────────────────────────


class TestInvoices:
    def test_invoices_generated(self):
        records = _make_revenue()
        invoices = generate_invoices(records)
        assert len(invoices) > 0

    def test_one_invoice_per_customer_per_month(self):
        """Each entity×month should produce one invoice per customer."""
        records = _make_revenue()
        invoices = generate_invoices(records, years=[2025])
        # 26 customers × 12 months = 312
        assert len(invoices) == len(CUSTOMERS) * 12

    def test_invoice_amounts_match_entity_revenue(self):
        """Sum of customer invoices per entity×month = entity monthly revenue."""
        records = _make_revenue()
        invoices = generate_invoices(records, years=[2025])
        # Sum invoices by (entity, month)
        inv_totals: dict[tuple[str, int], Decimal] = {}
        for inv in invoices:
            key = (inv.entity_code, inv.month)
            inv_totals[key] = inv_totals.get(key, Decimal(0)) + inv.amount

        # Compare to revenue (whole dollars, same rounding as invoices)
        from decimal import ROUND_HALF_UP as _RHU

        rev_totals: dict[tuple[str, int], Decimal] = {}
        for rec in records:
            if rec.year != 2025:
                continue
            key = (rec.entity_code, rec.month)
            rev = rec.revenue.quantize(Decimal("1"), rounding=_RHU)
            rev_totals[key] = rev_totals.get(key, Decimal(0)) + rev

        for key in sorted(inv_totals):
            # Allow rounding tolerance of $1 per customer in that entity×month
            entity_custs = sum(1 for c in CUSTOMERS if c.entity_code == key[0])
            assert abs(inv_totals[key] - rev_totals[key]) <= entity_custs

    def test_due_dates_reflect_dso(self):
        records = _make_revenue()
        invoices = generate_invoices(records, years=[2025])
        for inv in invoices:
            cust = next(c for c in CUSTOMERS if c.id == inv.customer_id)
            expected_due = inv.issue_date + datetime.timedelta(days=cust.dso)
            assert inv.due_date == expected_due

    def test_deterministic(self):
        inv1 = generate_invoices(_make_revenue())
        inv2 = generate_invoices(_make_revenue())
        assert len(inv1) == len(inv2)
        for a, b in zip(inv1, inv2):
            assert a.invoice_id == b.invoice_id
            assert a.amount == b.amount


# ── Receipt lifecycle tests ───────────────────────────────────────────────


class TestReceipts:
    def test_receipts_generated(self):
        records = _make_revenue()
        invoices = generate_invoices(records)
        colls = generate_collections(records)
        receipts = generate_receipts(invoices, colls)
        assert len(receipts) > 0

    def test_receipt_total_matches_invoice_total(self):
        """Every invoice produces exactly one receipt for the full amount."""
        records = _make_revenue()
        invoices = generate_invoices(records)
        receipts = generate_receipts(invoices)

        total_receipts = sum(r.amount for r in receipts)
        total_invoices = sum(i.amount for i in invoices)
        assert total_receipts == total_invoices

    def test_receipts_reference_valid_invoices(self):
        records = _make_revenue()
        invoices = generate_invoices(records)
        receipts = generate_receipts(invoices)

        invoice_ids = {inv.invoice_id for inv in invoices}
        for rct in receipts:
            assert rct.invoice_id in invoice_ids, (
                f"Receipt {rct.receipt_id} references unknown invoice {rct.invoice_id}"
            )

    def test_no_overpayment(self):
        """No invoice should receive more in receipts than its amount."""
        records = _make_revenue()
        invoices = generate_invoices(records)
        receipts = generate_receipts(invoices)

        inv_amounts = {inv.invoice_id: inv.amount for inv in invoices}
        paid: dict[str, Decimal] = {}
        for rct in receipts:
            paid[rct.invoice_id] = paid.get(rct.invoice_id, Decimal(0)) + rct.amount

        for inv_id, total_paid in paid.items():
            assert total_paid <= inv_amounts[inv_id], (
                f"Invoice {inv_id}: paid {total_paid} > amount {inv_amounts[inv_id]}"
            )

    def test_receipt_date_reflects_dso(self):
        """Receipt date should be issue_date + DSO (same month)."""
        records = _make_revenue()
        invoices = generate_invoices(records, years=[2025])
        receipts = generate_receipts(invoices)

        inv_lookup = {inv.invoice_id: inv for inv in invoices}
        for rct in receipts:
            inv = inv_lookup[rct.invoice_id]
            cust = next(c for c in CUSTOMERS if c.id == inv.customer_id)
            expected = inv.issue_date + datetime.timedelta(days=cust.dso)
            assert rct.receipt_date.year == expected.year
            assert rct.receipt_date.month == expected.month

    def test_deterministic(self):
        records = _make_revenue()
        invoices = generate_invoices(records)
        rct1 = generate_receipts(invoices)
        rct2 = generate_receipts(invoices)
        assert len(rct1) == len(rct2)
        for a, b in zip(rct1, rct2):
            assert a.receipt_id == b.receipt_id
            assert a.amount == b.amount


# ── Procure-to-pay lifecycle tie-out ─────────────────────────────────────────


class TestProcureToPayTieOut:
    """AP ledger normal transactions tie to AP subledger purchase volumes."""

    def test_vendor_normal_totals_match_targets(self):
        """Each vendor's normal txn total = target purchases − anomaly amounts."""
        records = _make_revenue()
        rng = random.Random(42)
        employees_rng = random.Random(42)
        from generator.model.employees import generate_employees

        employees = generate_employees(employees_rng)
        result = generate_ap_ledger(rng, employees, revenue_records=records)

        for vs in result.vendor_summaries:
            if vs.target_purchases == Decimal(0):
                continue
            expected_normal = vs.target_purchases - vs.anomaly_total
            assert vs.normal_total == expected_normal, (
                f"{vs.vendor_id}: normal_total={vs.normal_total} ≠ "
                f"target({vs.target_purchases}) − anomaly({vs.anomaly_total}) = "
                f"{expected_normal}"
            )

    def test_entity_ledger_totals_approximate_annual_purchases(self):
        """Per-entity ledger totals (including anomalies) should approximately
        match annual purchase volumes from the AP subledger model."""
        records = _make_revenue()
        vendor_targets = compute_vendor_annual_purchases(records, year=2025)

        rng = random.Random(42)
        employees_rng = random.Random(42)
        from generator.model.employees import generate_employees

        employees = generate_employees(employees_rng)
        result = generate_ap_ledger(rng, employees, revenue_records=records)

        # Sum targets and ledger totals by entity.
        for entity_code in ("PC", "AM", "DS"):
            target_total = sum(
                amt for vid, amt in vendor_targets.items()
                if next(v for v in VENDORS if v.id == vid).entity_code == entity_code
            )
            entity_vendors = {v.id for v in VENDORS if v.entity_code == entity_code}
            ledger_total = sum(
                vs.ledger_total for vs in result.vendor_summaries
                if vs.vendor_id in entity_vendors
            )
            # Allow small tolerance for anomaly amounts from non-trade vendors
            # (VEND-050+) that don't have targets.
            diff = abs(ledger_total - target_total)
            assert diff / target_total < Decimal("0.01"), (
                f"{entity_code}: ledger_total={ledger_total} vs "
                f"target={target_total} (diff={diff})"
            )

    def test_vendor_summaries_cover_all_vendors(self):
        """Vendor summaries should include all 26 trade vendors."""
        records = _make_revenue()
        rng = random.Random(42)
        employees_rng = random.Random(42)
        from generator.model.employees import generate_employees

        employees = generate_employees(employees_rng)
        result = generate_ap_ledger(rng, employees, revenue_records=records)

        summary_ids = {vs.vendor_id for vs in result.vendor_summaries}
        vendor_ids = {v.id for v in VENDORS}
        assert summary_ids == vendor_ids

    def test_ledger_total_equals_normal_plus_anomaly(self):
        """For each vendor, ledger_total = normal_total + anomaly_total."""
        records = _make_revenue()
        rng = random.Random(42)
        employees_rng = random.Random(42)
        from generator.model.employees import generate_employees

        employees = generate_employees(employees_rng)
        result = generate_ap_ledger(rng, employees, revenue_records=records)

        for vs in result.vendor_summaries:
            assert vs.ledger_total == vs.normal_total + vs.anomaly_total, (
                f"{vs.vendor_id}: {vs.ledger_total} ≠ "
                f"{vs.normal_total} + {vs.anomaly_total}"
            )
