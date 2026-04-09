"""Tests for generator.model.intercompany — intercompany ledger."""

import random
from decimal import Decimal

from generator.model.gl import Ledger
from generator.model.intercompany import (
    IC_LOAN_PRINCIPAL,
    IC_LOAN_RATE,
    MANAGEMENT_FEE_PCT,
    SERVICES_OPERATING_MARGIN,
    ICTransaction,
    generate_ic_transactions,
    post_ic_loan_principal,
    post_ic_transactions_to_gl,
)
from generator.model.revenue import generate_monthly_revenue


def _make_monthly_revenue_dict() -> dict[tuple[str, int, int], Decimal]:
    """Build entity-level monthly revenue dict from revenue module records."""
    rng = random.Random(42)
    records = generate_monthly_revenue(rng)
    totals: dict[tuple[str, int, int], Decimal] = {}
    for r in records:
        key = (r.entity_code, r.year, r.month)
        totals[key] = totals.get(key, Decimal(0)) + r.revenue
    return totals


def _make_transactions() -> list[ICTransaction]:
    rev = _make_monthly_revenue_dict()
    return generate_ic_transactions(rev)


class TestICTransactionGeneration:
    """Core IC transaction generation tests."""

    def test_all_types_present(self) -> None:
        txns = _make_transactions()
        types = {t.tx_type for t in txns}
        assert types == {"goods", "services", "management_fee", "interest"}

    def test_goods_seller_buyer(self) -> None:
        """Goods: PC sells to AM."""
        txns = [t for t in _make_transactions() if t.tx_type == "goods"]
        assert all(t.seller_entity == "PC" for t in txns)
        assert all(t.buyer_entity == "AM" for t in txns)

    def test_services_seller_is_ds(self) -> None:
        """Services: DS sells to PC and AM."""
        txns = [t for t in _make_transactions() if t.tx_type == "services"]
        assert all(t.seller_entity == "DS" for t in txns)
        buyers = {t.buyer_entity for t in txns}
        assert buyers == {"PC", "AM"}

    def test_management_fee_seller_is_ci(self) -> None:
        """Mgmt fees: CI charges PC, AM, DS."""
        txns = [t for t in _make_transactions() if t.tx_type == "management_fee"]
        assert all(t.seller_entity == "CI" for t in txns)
        buyers = {t.buyer_entity for t in txns}
        assert buyers == {"PC", "AM", "DS"}

    def test_interest_parties(self) -> None:
        """Interest: CI earns from AM."""
        txns = [t for t in _make_transactions() if t.tx_type == "interest"]
        assert all(t.seller_entity == "CI" for t in txns)
        assert all(t.buyer_entity == "AM" for t in txns)

    def test_interest_monthly_amount(self) -> None:
        """Each month's interest = $5M × 5% / 12 = $20,833."""
        txns = [t for t in _make_transactions() if t.tx_type == "interest"]
        expected = (IC_LOAN_PRINCIPAL * IC_LOAN_RATE / Decimal("12")).quantize(Decimal("1"))
        for t in txns:
            assert t.amount == expected

    def test_interest_count(self) -> None:
        """36 months of interest (3 years × 12)."""
        txns = [t for t in _make_transactions() if t.tx_type == "interest"]
        assert len(txns) == 36

    def test_all_amounts_positive(self) -> None:
        txns = _make_transactions()
        for t in txns:
            assert t.amount > 0, f"Non-positive IC amount: {t}"

    def test_transactions_sorted_by_date(self) -> None:
        txns = _make_transactions()
        dates = [t.date for t in txns]
        assert dates == sorted(dates)


class TestICConsolidationElimination:
    """The core acceptance criterion: all 9xxx accounts net to zero on consolidation."""

    def _build_consolidated_ledger(self) -> Ledger:
        ledger = Ledger()
        rev_dict = _make_monthly_revenue_dict()
        txns = generate_ic_transactions(rev_dict)
        post_ic_loan_principal(ledger)
        post_ic_transactions_to_gl(ledger, txns)
        return ledger

    def test_all_9xxx_accounts_net_to_zero(self) -> None:
        """On consolidation, the sum of all 9xxx IC accounts = zero across entities."""
        ledger = self._build_consolidated_ledger()

        # Aggregate balances across ALL entities
        consolidated: dict[str, Decimal] = {}
        for entity in ("CI", "PC", "AM", "DS"):
            for acct, bal in ledger.balance_by_account(entity).items():
                if acct.startswith("9"):
                    consolidated[acct] = consolidated.get(acct, Decimal(0)) + bal

        # Total of all 9xxx accounts must be zero
        total = sum(consolidated.values())
        assert total == Decimal(0), (
            f"IC accounts do not eliminate on consolidation: total = {total}\n"
            + "\n".join(f"  {a}: {b}" for a, b in sorted(consolidated.items()))
        )

    def test_revenue_expense_pairs_net_to_zero(self) -> None:
        """Each IC revenue/expense pair eliminates (same amounts, opposite signs)."""
        ledger = self._build_consolidated_ledger()

        consolidated: dict[str, Decimal] = {}
        for entity in ("CI", "PC", "AM", "DS"):
            for acct, bal in ledger.balance_by_account(entity).items():
                if acct.startswith("9"):
                    consolidated[acct] = consolidated.get(acct, Decimal(0)) + bal

        # Revenue (credit-normal) + Expense (debit-normal) pairs
        pairs = [
            ("9100", "9110"),  # Management fees
            ("9120", "9130"),  # Goods
            ("9140", "9150"),  # Interest
            ("9160", "9170"),  # Services
        ]
        for rev, exp in pairs:
            rev_bal = consolidated.get(rev, Decimal(0))
            exp_bal = consolidated.get(exp, Decimal(0))
            net = rev_bal + exp_bal
            assert net == Decimal(0), (
                f"Revenue/expense pair {rev}/{exp} doesn't eliminate: "
                f"{rev}={rev_bal}, {exp}={exp_bal}, net={net}"
            )

    def test_total_receivables_equal_total_payables(self) -> None:
        """Total IC receivables across all entities = total IC payables."""
        ledger = self._build_consolidated_ledger()

        total_recv = Decimal(0)
        total_pay = Decimal(0)
        recv_accts = {"9010", "9030", "9050", "9070"}
        pay_accts = {"9020", "9040", "9060", "9080"}

        for entity in ("CI", "PC", "AM", "DS"):
            for acct, bal in ledger.balance_by_account(entity).items():
                if acct in recv_accts:
                    total_recv += bal
                elif acct in pay_accts:
                    total_pay += bal

        # Receivables are debit-normal (positive), payables are credit-normal (negative)
        assert total_recv + total_pay == Decimal(0), (
            f"IC receivables ({total_recv}) + payables ({total_pay}) != 0"
        )

    def test_loan_principal_eliminates(self) -> None:
        """IC Loan Receivable (9200) + IC Loan Payable (9210) = 0."""
        ledger = self._build_consolidated_ledger()
        consolidated: dict[str, Decimal] = {}
        for entity in ("CI", "PC", "AM", "DS"):
            for acct, bal in ledger.balance_by_account(entity).items():
                consolidated[acct] = consolidated.get(acct, Decimal(0)) + bal

        loan_recv = consolidated.get("9200", Decimal(0))
        loan_pay = consolidated.get("9210", Decimal(0))
        assert loan_recv + loan_pay == Decimal(0), (
            f"Loan principal doesn't eliminate: 9200={loan_recv}, 9210={loan_pay}"
        )


class TestICServicesMargin:
    """Verify the services transfer runs at 11.2% operating margin (TC-09 flag)."""

    def test_services_operating_margin_is_112_pct(self) -> None:
        """The services IC margin should be exactly 11.2% for TC-09 detection."""
        # The SERVICES_OPERATING_MARGIN constant is what TC-09 gold standard expects
        assert SERVICES_OPERATING_MARGIN == Decimal("0.112")


class TestICManagementFees:
    """Verify management fee calculations."""

    def test_management_fees_are_15_pct_of_sub_revenue(self) -> None:
        """Each month's mgmt fee = 1.5% of that entity's monthly revenue."""
        rev_dict = _make_monthly_revenue_dict()
        txns = [t for t in _make_transactions() if t.tx_type == "management_fee"]

        for t in txns:
            entity_rev = rev_dict[(t.buyer_entity, t.date.year, t.date.month)]
            expected = (entity_rev * MANAGEMENT_FEE_PCT).quantize(Decimal("1"))
            assert t.amount == expected, (
                f"Mgmt fee mismatch for {t.buyer_entity} {t.date}: "
                f"expected {expected}, got {t.amount}"
            )
