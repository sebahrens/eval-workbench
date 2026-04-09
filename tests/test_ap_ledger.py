"""Tests for generator.model.ap_ledger — 52K AP transaction ledger with anomalies."""

from __future__ import annotations

import random
from decimal import Decimal

from generator.model.ap_ledger import (
    _EMPLOYEE_VENDOR_ID,
    _PACIFIC_CONSULTING_ID,
    _US_HOLIDAYS_2025,
    APLedgerResult,
    generate_ap_ledger,
)
from generator.model.employees import generate_employees


def _make_ledger() -> APLedgerResult:
    rng = random.Random(42)
    employees = generate_employees(rng)
    rng2 = random.Random(42)
    return generate_ap_ledger(rng2, employees)


class TestLedgerBasics:
    def test_row_count(self):
        result = _make_ledger()
        assert len(result.transactions) == 52_000

    def test_all_ids_unique(self):
        result = _make_ledger()
        ids = [t.transaction_id for t in result.transactions]
        assert len(ids) == len(set(ids))

    def test_sorted_by_id(self):
        result = _make_ledger()
        ids = [t.transaction_id for t in result.transactions]
        assert ids == sorted(ids)

    def test_all_amounts_positive(self):
        result = _make_ledger()
        for t in result.transactions:
            assert t.amount > Decimal("0"), f"{t.transaction_id} has non-positive amount"

    def test_deterministic(self):
        r1 = _make_ledger()
        r2 = _make_ledger()
        for a, b in zip(r1.transactions[:100], r2.transactions[:100]):
            assert a.transaction_id == b.transaction_id
            assert a.amount == b.amount
            assert a.vendor_id == b.vendor_id


class TestDuplicateAnomaly:
    def test_exact_duplicate_count(self):
        result = _make_ledger()
        exact = result.anomaly_index.get("exact_duplicate", [])
        assert len(exact) == 8  # 4 pairs

    def test_near_duplicate_count(self):
        result = _make_ledger()
        near = result.anomaly_index.get("near_duplicate", [])
        assert len(near) == 8  # 4 pairs

    def test_total_duplicate_exposure(self):
        """Exposure = sum of duplicated amounts (second copy in each pair)."""
        result = _make_ledger()
        by_id = {t.transaction_id: t for t in result.transactions}

        exposure = Decimal("0")
        # Exact duplicates: every other ID is the duplicate copy
        exact_ids = result.anomaly_index["exact_duplicate"]
        for i in range(1, len(exact_ids), 2):
            exposure += by_id[exact_ids[i]].amount

        near_ids = result.anomaly_index["near_duplicate"]
        for i in range(1, len(near_ids), 2):
            exposure += by_id[near_ids[i]].amount

        assert exposure == Decimal("127340.00")


class TestBenfordAnomaly:
    def test_count(self):
        result = _make_ledger()
        benford = result.anomaly_index.get("benford_violation", [])
        assert len(benford) == 35

    def test_amounts_in_range(self):
        result = _make_ledger()
        by_id = {t.transaction_id: t for t in result.transactions}
        for tid in result.anomaly_index["benford_violation"]:
            amt = by_id[tid].amount
            assert Decimal("9900") <= amt <= Decimal("9999.99"), (
                f"{tid} amount {amt} outside $9,900-$9,999 range"
            )


class TestRoundNumberAnomaly:
    def test_count(self):
        result = _make_ledger()
        rnd = result.anomaly_index.get("round_number", [])
        assert len(rnd) == 12

    def test_all_to_pacific_consulting(self):
        result = _make_ledger()
        by_id = {t.transaction_id: t for t in result.transactions}
        for tid in result.anomaly_index["round_number"]:
            assert by_id[tid].vendor_id == _PACIFIC_CONSULTING_ID

    def test_amounts_are_round(self):
        result = _make_ledger()
        by_id = {t.transaction_id: t for t in result.transactions}
        valid = {Decimal("5000.00"), Decimal("10000.00"), Decimal("25000.00")}
        for tid in result.anomaly_index["round_number"]:
            assert by_id[tid].amount in valid


class TestTemporalAnomaly:
    def test_weekend_holiday_count(self):
        result = _make_ledger()
        wh = result.anomaly_index.get("weekend_holiday", [])
        assert len(wh) == 15

    def test_all_on_non_business_days(self):
        result = _make_ledger()
        by_id = {t.transaction_id: t for t in result.transactions}
        for tid in result.anomaly_index["weekend_holiday"]:
            d = by_id[tid].date
            assert d.weekday() >= 5 or d in _US_HOLIDAYS_2025

    def test_invoice_after_payment_count(self):
        result = _make_ledger()
        iap = result.anomaly_index.get("invoice_after_payment", [])
        assert len(iap) == 8


class TestVendorAnomaly:
    def test_similar_name_count(self):
        result = _make_ledger()
        sn = result.anomaly_index.get("similar_name_vendor", [])
        assert len(sn) == 14  # 8 LLC + 6 Inc

    def test_employee_address_vendor_count(self):
        result = _make_ledger()
        ea = result.anomaly_index.get("employee_address_vendor", [])
        assert len(ea) == 5

    def test_employee_address_vendor_id(self):
        result = _make_ledger()
        by_id = {t.transaction_id: t for t in result.transactions}
        for tid in result.anomaly_index["employee_address_vendor"]:
            assert by_id[tid].vendor_id == _EMPLOYEE_VENDOR_ID


class TestSplitAnomaly:
    def test_split_count(self):
        result = _make_ledger()
        sp = result.anomaly_index.get("split_transaction", [])
        assert len(sp) == 6  # 3 pairs


class TestApproverAnomaly:
    def test_self_approved_count(self):
        result = _make_ledger()
        sa = result.anomaly_index.get("self_approved", [])
        assert len(sa) == 1

    def test_self_approved_vendor_is_approver(self):
        result = _make_ledger()
        by_id = {t.transaction_id: t for t in result.transactions}
        tid = result.anomaly_index["self_approved"][0]
        txn = by_id[tid]
        assert txn.vendor_name == txn.approver

    def test_single_approver_cc_count(self):
        result = _make_ledger()
        sc = result.anomaly_index.get("single_approver_cc", [])
        assert len(sc) == 40

    def test_single_approver_dominance(self):
        result = _make_ledger()
        by_id = {t.transaction_id: t for t in result.transactions}
        approvers = [
            by_id[tid].approver
            for tid in result.anomaly_index["single_approver_cc"]
        ]
        from collections import Counter
        counts = Counter(approvers)
        dominant = counts.most_common(1)[0][1]
        assert dominant == 38  # 95% of 40
