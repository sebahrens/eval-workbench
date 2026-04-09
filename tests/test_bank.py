"""Tests for generator.model.bank — bank transactions and reconciliation."""

import random
from decimal import Decimal

from generator.model.bank import (
    ADJUSTED_BALANCE,
    BANK_INTEREST,
    BANK_SERVICE_CHARGES,
    CONFIRMATION_BALANCE,
    DEPOSITS_IN_TRANSIT,
    GL_ENDING_BALANCE,
    OUTSTANDING_CHECKS,
    TOTAL_DEPOSITS_IN_TRANSIT,
    TOTAL_OUTSTANDING_CHECKS,
    BankModel,
    generate_bank_model,
    validate_reconciliation,
)


def _make_model() -> BankModel:
    rng = random.Random(42)
    return generate_bank_model(rng)


class TestGoldStandardConstants:
    """Verify the gold standard reconciliation constants from prompt.md."""

    def test_adjusted_balance(self) -> None:
        assert ADJUSTED_BALANCE == Decimal("4287331")

    def test_confirmation_balance(self) -> None:
        assert CONFIRMATION_BALANCE == Decimal("4312117")

    def test_outstanding_checks_count(self) -> None:
        assert len(OUTSTANDING_CHECKS) == 4

    def test_deposits_in_transit_count(self) -> None:
        assert len(DEPOSITS_IN_TRANSIT) == 2

    def test_bank_side_reconciliation(self) -> None:
        """Bank ending + DIT - OC = adjusted balance."""
        result = CONFIRMATION_BALANCE + TOTAL_DEPOSITS_IN_TRANSIT - TOTAL_OUTSTANDING_CHECKS
        assert result == ADJUSTED_BALANCE

    def test_book_side_reconciliation(self) -> None:
        """GL ending + interest - charges = adjusted balance."""
        result = GL_ENDING_BALANCE + BANK_INTEREST - BANK_SERVICE_CHARGES
        assert result == ADJUSTED_BALANCE


class TestBankModelGeneration:
    """Test the generated bank model."""

    def test_bank_transaction_count(self) -> None:
        """Bank statement must have exactly 340 transactions."""
        model = _make_model()
        assert len(model.bank_transactions) == 340

    def test_bank_ending_balance(self) -> None:
        model = _make_model()
        last_txn = model.bank_transactions[-1]
        assert last_txn.running_balance == CONFIRMATION_BALANCE

    def test_gl_ending_balance(self) -> None:
        model = _make_model()
        last_entry = model.gl_entries[-1]
        assert last_entry.running_balance == GL_ENDING_BALANCE

    def test_outstanding_checks_in_gl(self) -> None:
        model = _make_model()
        oc_entries = [e for e in model.gl_entries if e.category == "outstanding_check"]
        assert len(oc_entries) == 4
        oc_total = sum(e.credit for e in oc_entries)
        assert oc_total == TOTAL_OUTSTANDING_CHECKS

    def test_deposits_in_transit_in_gl(self) -> None:
        model = _make_model()
        dit_entries = [e for e in model.gl_entries if e.category == "deposit_in_transit"]
        assert len(dit_entries) == 2
        dit_total = sum(e.debit for e in dit_entries)
        assert dit_total == TOTAL_DEPOSITS_IN_TRANSIT

    def test_bank_only_items(self) -> None:
        """Interest and charges appear on bank statement but not GL."""
        model = _make_model()
        bank_only = [t for t in model.bank_transactions if not t.matched]
        assert len(bank_only) == 2
        categories = {t.category for t in bank_only}
        assert categories == {"interest", "service_charge"}

    def test_gl_only_items(self) -> None:
        """Outstanding checks and deposits in transit in GL but not bank."""
        model = _make_model()
        gl_only = [e for e in model.gl_entries if not e.matched]
        assert len(gl_only) == 6  # 4 OC + 2 DIT

    def test_running_balance_monotonic_start(self) -> None:
        """Bank statement running balance starts from starting balance."""
        model = _make_model()
        first = model.bank_transactions[0]
        expected = model.bank_starting_balance + first.amount
        assert first.running_balance == expected

    def test_gl_dates_in_december(self) -> None:
        """All GL entries should be in December 2025."""
        model = _make_model()
        for entry in model.gl_entries:
            assert entry.date.year == 2025
            assert entry.date.month == 12

    def test_bank_dates_in_december(self) -> None:
        """All bank transactions should be in December 2025."""
        model = _make_model()
        for txn in model.bank_transactions:
            assert txn.date.year == 2025
            assert txn.date.month == 12

    def test_bank_dates_sorted(self) -> None:
        """Bank transactions should be in chronological order."""
        model = _make_model()
        dates = [t.date for t in model.bank_transactions]
        assert dates == sorted(dates)


class TestReconciliationValidation:
    """Test the reconciliation validator."""

    def test_valid_model_passes(self) -> None:
        model = _make_model()
        errors = validate_reconciliation(model)
        assert errors == [], f"Validation errors: {errors}"


class TestDeterminism:
    """Verify that the bank model is deterministic."""

    def test_same_seed_same_output(self) -> None:
        model1 = generate_bank_model(random.Random(42))
        model2 = generate_bank_model(random.Random(42))

        assert len(model1.bank_transactions) == len(model2.bank_transactions)
        for t1, t2 in zip(model1.bank_transactions, model2.bank_transactions):
            assert t1.date == t2.date
            assert t1.description == t2.description
            assert t1.amount == t2.amount
            assert t1.running_balance == t2.running_balance

        assert len(model1.gl_entries) == len(model2.gl_entries)
        for e1, e2 in zip(model1.gl_entries, model2.gl_entries):
            assert e1.date == e2.date
            assert e1.description == e2.description
            assert e1.debit == e2.debit
            assert e1.credit == e2.credit
            assert e1.running_balance == e2.running_balance
