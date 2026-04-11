"""Tests for TC-13 — Forensic AP Transaction Analysis formatter.

Verifies:
- 52,000-row AP transaction CSV with canary
- All 7 anomaly categories with correct counts:
  - 8 duplicate payments (4 exact + 4 near-dup), exposure $127,340
  - 35 Benford's Law violations ($9,900-$9,999)
  - 12 round-number payments to Pacific Consulting Group
  - 15 weekend/holiday + 8 invoice-after-payment temporal anomalies
  - Vendor anomalies: JKL Services LLC/Inc + employee-address vendor PC-0342
  - 3 sets of split transactions below $10K threshold
  - Approver anomalies: 1 self-approved + 1 single-approver cost center
- Gold standard structure and scoring hints
- Prompt and expected behavior markdown files
"""

from __future__ import annotations

import json
import random
import tempfile
from decimal import Decimal
from pathlib import Path

from generator.canaries import CanaryRegistry
from generator.canaries import build_registry as build_canary_registry
from generator.errors import ErrorRegistry
from generator.formatters.tc13 import _CSV_COLUMNS, emit_tc13
from generator.golds.framework import emit_gold
from generator.manifest import Manifest
from generator.model.ap_ledger import (
    _EMPLOYEE_VENDOR_NAME,
    _JKL_INC_NAME,
    _JKL_LLC_NAME,
    _PACIFIC_CONSULTING_NAME,
    _SINGLE_APPROVER_CC,
    _TARGET_EMPLOYEE_ID,
    APLedgerResult,
    generate_ap_ledger,
)
from generator.model.build import CascadeModel, build_model

# ---------------------------------------------------------------------------
# Module-level fixture: build the model and run emit_tc13 once
# ---------------------------------------------------------------------------

_MODEL: CascadeModel | None = None
_OUTPUT: Path | None = None
_CANARIES: CanaryRegistry | None = None
_ERRORS: ErrorRegistry | None = None
_MANIFEST: Manifest | None = None

_CANARY_KEYS = sorted(["tc13_ap_transactions"])


def _ensure_emitted() -> tuple[CascadeModel, Path, CanaryRegistry, ErrorRegistry, Manifest]:
    global _MODEL, _OUTPUT, _CANARIES, _ERRORS, _MANIFEST  # noqa: PLW0603
    if _MODEL is None:
        _MODEL = build_model(seed=42)
        _OUTPUT = Path(tempfile.mkdtemp(prefix="tc13_test_"))
        _CANARIES = build_canary_registry(_CANARY_KEYS, seed=42)
        _ERRORS = ErrorRegistry()
        _MANIFEST = Manifest(_OUTPUT)
        _MANIFEST.__enter__()

        emit_tc13(_MODEL, _OUTPUT, _CANARIES, _ERRORS, _MANIFEST)

        # Emit gold standard (registered via @register_gold)
        emit_gold("TC-13", _CANARIES, _ERRORS, _OUTPUT / "gold_standards", model=_MODEL)

        _MANIFEST.__exit__(None, None, None)
    return _MODEL, _OUTPUT, _CANARIES, _ERRORS, _MANIFEST


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_LEDGER: APLedgerResult | None = None


def _ensure_ledger() -> APLedgerResult:
    """Generate the AP ledger once for anomaly tests."""
    global _LEDGER  # noqa: PLW0603
    if _LEDGER is None:
        model = build_model(seed=42)
        rng = random.Random(42)
        _LEDGER = generate_ap_ledger(rng, model.employees)
    return _LEDGER


_INPUT_DIR = "test_cases/TC-13/input_files"


# ---------------------------------------------------------------------------
# CSV file structure
# ---------------------------------------------------------------------------


class TestCSVStructure:
    """Verify the AP transactions CSV file is well-formed."""

    def test_csv_exists(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        csv_path = output / _INPUT_DIR / "ap_transactions_fy2025.csv"
        assert csv_path.exists()

    def test_csv_row_count_52000(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        csv_path = output / _INPUT_DIR / "ap_transactions_fy2025.csv"
        lines = csv_path.read_text().splitlines()
        # First line is canary comment, second is header, rest are data
        canary_line = lines[0]
        assert canary_line.startswith("# CANARY:")
        header = lines[1]
        assert header == ",".join(_CSV_COLUMNS)
        data_rows = lines[2:]
        assert len(data_rows) == 52_000

    def test_csv_columns_match_spec(self) -> None:
        expected = [
            "transaction_id", "date", "vendor_id", "vendor_name",
            "amount", "description", "approver", "cost_center",
            "payment_method", "invoice_number",
        ]
        assert _CSV_COLUMNS == expected

    def test_canary_embedded(self) -> None:
        _, output, canaries, _, _ = _ensure_emitted()
        csv_path = output / _INPUT_DIR / "ap_transactions_fy2025.csv"
        first_line = csv_path.read_text().splitlines()[0]
        canary = canaries.canary_for("tc13_ap_transactions")
        assert canary in first_line


# ---------------------------------------------------------------------------
# Anomaly counts from the ledger model
# ---------------------------------------------------------------------------


class TestDuplicatePayments:
    """Anomaly 1: 4 exact + 4 near duplicates, exposure $127,340."""

    def test_exact_duplicate_count(self) -> None:
        result = _ensure_ledger()
        exact = result.anomaly_index.get("exact_duplicate", [])
        assert len(exact) == 8  # 4 pairs = 8 txn IDs

    def test_near_duplicate_count(self) -> None:
        result = _ensure_ledger()
        near = result.anomaly_index.get("near_duplicate", [])
        assert len(near) == 8  # 4 pairs = 8 txn IDs

    def test_total_duplicate_exposure_127340(self) -> None:
        result = _ensure_ledger()
        # Exposure = sum of the second copy in each pair
        exact_ids = result.anomaly_index.get("exact_duplicate", [])
        near_ids = result.anomaly_index.get("near_duplicate", [])

        exact_txns = [t for t in result.transactions if t.transaction_id in exact_ids]
        near_txns = [t for t in result.transactions if t.transaction_id in near_ids]

        # Every other transaction is the duplicate (pairs)
        exact_exposure = sum(
            t.amount for i, t in enumerate(
                [t for t in exact_txns if t.anomaly_type == "exact_duplicate"]
            ) if i % 2 == 1
        )
        near_exposure = sum(
            t.amount for i, t in enumerate(
                [t for t in near_txns if t.anomaly_type == "near_duplicate"]
            ) if i % 2 == 1
        )
        total = exact_exposure + near_exposure
        assert total == Decimal("127340.00")

    def test_exact_duplicates_share_invoice(self) -> None:
        """Exact duplicates have identical vendor + amount + invoice."""
        result = _ensure_ledger()
        exact_ids = result.anomaly_index["exact_duplicate"]
        exact_txns = [t for t in result.transactions if t.transaction_id in exact_ids]
        # Group into pairs (consecutive IDs)
        for i in range(0, len(exact_txns), 2):
            a, b = exact_txns[i], exact_txns[i + 1]
            assert a.vendor_id == b.vendor_id
            assert a.amount == b.amount
            assert a.invoice_number == b.invoice_number

    def test_near_duplicates_differ_by_one_digit(self) -> None:
        """Near-duplicate invoice numbers differ by exactly 1 digit in last position."""
        result = _ensure_ledger()
        near_ids = result.anomaly_index["near_duplicate"]
        near_txns = [t for t in result.transactions if t.transaction_id in near_ids]
        for i in range(0, len(near_txns), 2):
            a, b = near_txns[i], near_txns[i + 1]
            assert a.vendor_id == b.vendor_id
            assert a.amount == b.amount
            # Invoice numbers differ in last character only
            assert a.invoice_number[:-1] == b.invoice_number[:-1]
            assert a.invoice_number[-1] != b.invoice_number[-1]


class TestBenfordViolation:
    """Anomaly 2: 35 transactions between $9,900-$9,999."""

    def test_benford_count(self) -> None:
        result = _ensure_ledger()
        benford = result.anomaly_index.get("benford_violation", [])
        assert len(benford) == 35

    def test_benford_amounts_in_range(self) -> None:
        result = _ensure_ledger()
        benford_ids = result.anomaly_index["benford_violation"]
        benford_txns = [t for t in result.transactions if t.transaction_id in benford_ids]
        for t in benford_txns:
            assert Decimal("9900.00") <= t.amount <= Decimal("9999.99")


class TestRoundNumberPayments:
    """Anomaly 3: 12 round-number payments to Pacific Consulting Group."""

    def test_round_number_count(self) -> None:
        result = _ensure_ledger()
        rn = result.anomaly_index.get("round_number", [])
        assert len(rn) == 12

    def test_round_number_vendor(self) -> None:
        result = _ensure_ledger()
        rn_ids = result.anomaly_index["round_number"]
        rn_txns = [t for t in result.transactions if t.transaction_id in rn_ids]
        for t in rn_txns:
            assert t.vendor_name == _PACIFIC_CONSULTING_NAME

    def test_round_number_amounts(self) -> None:
        result = _ensure_ledger()
        rn_ids = result.anomaly_index["round_number"]
        rn_txns = [t for t in result.transactions if t.transaction_id in rn_ids]
        allowed = {Decimal("5000.00"), Decimal("10000.00"), Decimal("25000.00")}
        for t in rn_txns:
            assert t.amount in allowed


class TestTemporalAnomalies:
    """Anomaly 4: 15 weekend/holiday + 8 invoice-after-payment."""

    def test_weekend_holiday_count(self) -> None:
        result = _ensure_ledger()
        wh = result.anomaly_index.get("weekend_holiday", [])
        assert len(wh) == 15

    def test_invoice_after_payment_count(self) -> None:
        result = _ensure_ledger()
        iap = result.anomaly_index.get("invoice_after_payment", [])
        assert len(iap) == 8

    def test_weekend_holiday_dates_are_non_business(self) -> None:
        from generator.model.ap_ledger import _is_non_business_day

        result = _ensure_ledger()
        wh_ids = result.anomaly_index["weekend_holiday"]
        wh_txns = [t for t in result.transactions if t.transaction_id in wh_ids]
        for t in wh_txns:
            assert _is_non_business_day(t.date), f"{t.transaction_id} date {t.date} is a business day"

    def test_invoice_after_payment_description_contains_future_date(self) -> None:
        """Invoice date embedded in description must be after the payment date."""
        import re

        result = _ensure_ledger()
        iap_ids = result.anomaly_index["invoice_after_payment"]
        iap_txns = [t for t in result.transactions if t.transaction_id in iap_ids]
        for t in iap_txns:
            # Description pattern: "Invoice INV-xxx-xxxxx dated YYYY-MM-DD — ..."
            match = re.search(r"dated (\d{4}-\d{2}-\d{2})", t.description)
            assert match, f"No date in description for {t.transaction_id}"
            from datetime import date as Date
            invoice_date = Date.fromisoformat(match.group(1))
            assert invoice_date > t.date, (
                f"{t.transaction_id}: invoice date {invoice_date} not after payment date {t.date}"
            )


class TestVendorAnomalies:
    """Anomaly 5: similar-name vendors + employee-address vendor."""

    def test_similar_name_vendor_count(self) -> None:
        result = _ensure_ledger()
        snv = result.anomaly_index.get("similar_name_vendor", [])
        # 8 JKL LLC + 6 JKL Inc = 14 transactions
        assert len(snv) == 14

    def test_similar_name_vendors_are_jkl(self) -> None:
        result = _ensure_ledger()
        snv_ids = result.anomaly_index["similar_name_vendor"]
        snv_txns = [t for t in result.transactions if t.transaction_id in snv_ids]
        vendor_names = {t.vendor_name for t in snv_txns}
        assert vendor_names == {_JKL_LLC_NAME, _JKL_INC_NAME}

    def test_employee_address_vendor_count(self) -> None:
        result = _ensure_ledger()
        eav = result.anomaly_index.get("employee_address_vendor", [])
        assert len(eav) == 5

    def test_employee_address_vendor_name(self) -> None:
        result = _ensure_ledger()
        eav_ids = result.anomaly_index["employee_address_vendor"]
        eav_txns = [t for t in result.transactions if t.transaction_id in eav_ids]
        for t in eav_txns:
            assert t.vendor_name == _EMPLOYEE_VENDOR_NAME

    def test_employee_address_vendor_links_to_target(self) -> None:
        """Anomaly detail references the target employee ID."""
        result = _ensure_ledger()
        eav_ids = result.anomaly_index["employee_address_vendor"]
        eav_txns = [t for t in result.transactions if t.transaction_id in eav_ids]
        for t in eav_txns:
            assert _TARGET_EMPLOYEE_ID in (t.anomaly_detail or "")


class TestSplitTransactions:
    """Anomaly 6: 3 sets of split transactions below $10K threshold."""

    def test_split_transaction_count(self) -> None:
        result = _ensure_ledger()
        st = result.anomaly_index.get("split_transaction", [])
        assert len(st) == 6  # 3 pairs = 6 txn IDs

    def test_split_pairs_same_vendor_and_approver(self) -> None:
        result = _ensure_ledger()
        st_ids = result.anomaly_index["split_transaction"]
        st_txns = sorted(
            [t for t in result.transactions if t.transaction_id in st_ids],
            key=lambda t: t.transaction_id,
        )
        for i in range(0, len(st_txns), 2):
            a, b = st_txns[i], st_txns[i + 1]
            assert a.vendor_id == b.vendor_id
            assert a.approver == b.approver
            assert a.cost_center == b.cost_center

    def test_split_individual_amounts_below_threshold(self) -> None:
        result = _ensure_ledger()
        st_ids = result.anomaly_index["split_transaction"]
        st_txns = [t for t in result.transactions if t.transaction_id in st_ids]
        for t in st_txns:
            assert t.amount < Decimal("10000.00")

    def test_split_combined_amounts(self) -> None:
        """Each split pair sums above the $10K threshold (except the $9,750 pair)."""
        result = _ensure_ledger()
        st_ids = result.anomaly_index["split_transaction"]
        st_txns = sorted(
            [t for t in result.transactions if t.transaction_id in st_ids],
            key=lambda t: t.transaction_id,
        )
        expected_sums = [
            Decimal("14500.00"),
            Decimal("9750.00"),
            Decimal("12000.00"),
        ]
        for i in range(0, len(st_txns), 2):
            pair_sum = st_txns[i].amount + st_txns[i + 1].amount
            assert pair_sum == expected_sums[i // 2]


class TestApproverAnomalies:
    """Anomaly 7: self-approved + single-approver cost center."""

    def test_self_approved_count(self) -> None:
        result = _ensure_ledger()
        sa = result.anomaly_index.get("self_approved", [])
        assert len(sa) == 1

    def test_self_approved_vendor_equals_approver(self) -> None:
        """The vendor name (payee) matches the approver name."""
        result = _ensure_ledger()
        sa_ids = result.anomaly_index["self_approved"]
        sa_txns = [t for t in result.transactions if t.transaction_id in sa_ids]
        for t in sa_txns:
            assert t.vendor_name == t.approver

    def test_single_approver_cc_count(self) -> None:
        result = _ensure_ledger()
        sacc = result.anomaly_index.get("single_approver_cc", [])
        assert len(sacc) == 40

    def test_single_approver_cc_is_2600(self) -> None:
        result = _ensure_ledger()
        sacc_ids = result.anomaly_index["single_approver_cc"]
        sacc_txns = [t for t in result.transactions if t.transaction_id in sacc_ids]
        for t in sacc_txns:
            assert t.cost_center == _SINGLE_APPROVER_CC

    def test_single_approver_dominance_95_percent(self) -> None:
        """95% (38/40) of CC 2600 transactions have the same approver."""
        result = _ensure_ledger()
        sacc_ids = result.anomaly_index["single_approver_cc"]
        sacc_txns = [t for t in result.transactions if t.transaction_id in sacc_ids]
        approver_counts: dict[str, int] = {}
        for t in sacc_txns:
            approver_counts[t.approver] = approver_counts.get(t.approver, 0) + 1
        dominant = max(approver_counts.values())
        assert dominant == 38


# ---------------------------------------------------------------------------
# Total row count
# ---------------------------------------------------------------------------


class TestTotalTransactions:
    """Verify the ledger has exactly 52,000 transactions."""

    def test_total_transaction_count(self) -> None:
        result = _ensure_ledger()
        assert len(result.transactions) == 52_000

    def test_all_anomaly_categories_present(self) -> None:
        result = _ensure_ledger()
        expected_types = {
            "exact_duplicate",
            "near_duplicate",
            "benford_violation",
            "round_number",
            "weekend_holiday",
            "invoice_after_payment",
            "similar_name_vendor",
            "employee_address_vendor",
            "split_transaction",
            "self_approved",
            "single_approver_cc",
        }
        assert set(result.anomaly_index.keys()) == expected_types


# ---------------------------------------------------------------------------
# Gold standard
# ---------------------------------------------------------------------------


class TestGoldStandard:
    """Verify TC-13 gold standard JSON structure."""

    def test_gold_json_exists(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        gold_path = output / "gold_standards" / "TC-13_gold.json"
        assert gold_path.exists()

    def test_gold_has_expected_keys(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        gold_path = output / "gold_standards" / "TC-13_gold.json"
        gold = json.loads(gold_path.read_text())
        assert "expected_outputs" in gold
        assert "canary_verification" in gold
        assert "scoring_hints" in gold

    def test_gold_total_transactions(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        gold_path = output / "gold_standards" / "TC-13_gold.json"
        gold = json.loads(gold_path.read_text())
        assert gold["expected_outputs"]["total_transactions"] == 52_000

    def test_gold_anomaly_categories(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        gold_path = output / "gold_standards" / "TC-13_gold.json"
        gold = json.loads(gold_path.read_text())
        categories = gold["expected_outputs"]["anomaly_categories"]
        assert "duplicate_payments" in categories
        assert "benford_violation" in categories
        assert "round_number" in categories
        assert "temporal_anomalies" in categories
        assert "vendor_anomalies" in categories
        assert "split_transactions" in categories
        assert "approver_anomalies" in categories

    def test_gold_duplicate_exposure(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        gold_path = output / "gold_standards" / "TC-13_gold.json"
        gold = json.loads(gold_path.read_text())
        exposure = gold["expected_outputs"]["anomaly_categories"]["duplicate_payments"]["total_exposure"]
        assert exposure == "127340.00"

    def test_gold_canary_verification(self) -> None:
        _, output, canaries, _, _ = _ensure_emitted()
        gold_path = output / "gold_standards" / "TC-13_gold.json"
        gold = json.loads(gold_path.read_text())
        assert gold["canary_verification"]["read_ap_transactions"] == canaries.canary_for("tc13_ap_transactions")

    def test_gold_scoring_hints(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        gold_path = output / "gold_standards" / "TC-13_gold.json"
        gold = json.loads(gold_path.read_text())
        hints = gold["scoring_hints"]
        assert "$127,340" in hints["accuracy"]
        assert _TARGET_EMPLOYEE_ID in hints["accuracy"]


# ---------------------------------------------------------------------------
# Prompt and expected behavior files
# ---------------------------------------------------------------------------


class TestPromptAndExpectedBehavior:
    """Verify TC-13 markdown files exist with expected content."""

    def test_prompt_exists(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "test_cases" / "TC-13" / "prompt.md"
        assert path.exists()

    def test_prompt_mentions_forensic(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "test_cases" / "TC-13" / "prompt.md"
        text = path.read_text()
        assert "forensic" in text.lower()
        assert "accounts payable" in text.lower()

    def test_expected_behavior_exists(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "test_cases" / "TC-13" / "expected_behavior.md"
        assert path.exists()

    def test_expected_behavior_mentions_all_categories(self) -> None:
        _, output, _, _, _ = _ensure_emitted()
        path = output / "test_cases" / "TC-13" / "expected_behavior.md"
        text = path.read_text()
        assert "Duplicate" in text
        assert "Benford" in text
        assert "Round" in text or "round" in text
        assert "Temporal" in text
        assert "Vendor" in text
        assert "Split" in text
        assert "Approver" in text
        assert "$127,340" in text
        assert _TARGET_EMPLOYEE_ID in text
