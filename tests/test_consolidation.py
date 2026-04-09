"""Tests for generator.model.consolidation — consolidation engine + IC eliminations."""

import datetime
import random
from decimal import Decimal

from generator.model.consolidation import (
    build_balance_sheet,
    build_cash_flow,
    build_income_statement,
    consolidated_trial_balance_eliminated,
    validate_consolidation,
    verify_ic_elimination,
)
from generator.model.gl import Ledger
from generator.model.intercompany import (
    generate_ic_transactions,
    post_ic_loan_principal,
    post_ic_transactions_to_gl,
)
from generator.model.revenue import generate_monthly_revenue, post_revenue_to_gl


def _build_ledger() -> Ledger:
    """Build a ledger with revenue + IC transactions for testing."""
    rng = random.Random(42)
    records = generate_monthly_revenue(rng)
    ledger = Ledger()
    post_revenue_to_gl(ledger, records)

    totals: dict[tuple[str, int, int], Decimal] = {}
    for r in records:
        key = (r.entity_code, r.year, r.month)
        totals[key] = totals.get(key, Decimal(0)) + r.revenue

    ic_txns = generate_ic_transactions(totals)
    post_ic_loan_principal(ledger)
    post_ic_transactions_to_gl(ledger, ic_txns)
    return ledger


class TestICElimination:
    """IC accounts must net to zero on consolidation."""

    def test_ic_zero_all_years(self) -> None:
        ledger = _build_ledger()
        for year in (2023, 2024, 2025):
            eoy = datetime.date(year, 12, 31)
            is_zero, total, _ = verify_ic_elimination(ledger, eoy)
            assert is_zero, f"FY{year}: IC imbalance = {total}"

    def test_eliminated_tb_has_no_9xxx(self) -> None:
        ledger = _build_ledger()
        tb = consolidated_trial_balance_eliminated(
            ledger, datetime.date(2025, 12, 31)
        )
        ic_accts = [a for a in tb if a.startswith("9")]
        assert ic_accts == [], f"9xxx accounts remain after elimination: {ic_accts}"


class TestBalanceSheet:
    """Consolidated BS must balance (A = L + E) for all years."""

    def test_bs_balanced_all_years(self) -> None:
        ledger = _build_ledger()
        for year in (2023, 2024, 2025):
            bs = build_balance_sheet(ledger, datetime.date(year, 12, 31))
            assert bs.is_balanced, (
                f"FY{year}: A={bs.total_assets} != L+E="
                f"{bs.total_liabilities + bs.total_equity}"
            )

    def test_assets_positive(self) -> None:
        ledger = _build_ledger()
        bs = build_balance_sheet(ledger, datetime.date(2025, 12, 31))
        assert bs.total_assets > 0


class TestIncomeStatement:
    """Consolidated IS must show correct revenue."""

    def test_fy2025_revenue_in_range(self) -> None:
        ledger = _build_ledger()
        stmt = build_income_statement(ledger, 2025)
        assert Decimal("198_000_000") <= stmt.total_revenue <= Decimal("202_000_000"), (
            f"FY2025 revenue {stmt.total_revenue} not in $198–202M range"
        )

    def test_gross_profit_positive(self) -> None:
        ledger = _build_ledger()
        for year in (2023, 2024, 2025):
            stmt = build_income_statement(ledger, year)
            assert stmt.gross_profit > 0, f"FY{year}: negative gross profit"

    def test_net_income_positive(self) -> None:
        ledger = _build_ledger()
        for year in (2023, 2024, 2025):
            stmt = build_income_statement(ledger, year)
            assert stmt.net_income > 0, f"FY{year}: negative net income"

    def test_no_ic_in_revenue(self) -> None:
        """IC revenue/expense must not inflate consolidated revenue."""
        ledger = _build_ledger()
        stmt = build_income_statement(ledger, 2025)
        # Revenue should be ~$200M, not inflated by IC
        assert stmt.total_revenue < Decimal("210_000_000")


class TestCashFlow:
    """Cash flow statement basic checks."""

    def test_ending_cash_matches_bs(self) -> None:
        ledger = _build_ledger()
        cf = build_cash_flow(ledger, 2025)
        bs = build_balance_sheet(ledger, datetime.date(2025, 12, 31))
        # Cash on BS (1010 + 1020 + 1030 + 1050)
        bs_cash = sum(
            bs.assets.get(acct, Decimal(0))
            for acct in ("1010", "1020", "1030", "1050")
        )
        assert cf.ending_cash == bs_cash


class TestValidateConsolidation:
    """Integration test for the validate_consolidation helper."""

    def test_all_checks_pass(self) -> None:
        ledger = _build_ledger()
        for year in (2023, 2024, 2025):
            results = validate_consolidation(ledger, year)
            assert results["ic_accounts_zero"] is True
            assert results["bs_balanced"] is True

    def test_fy2025_revenue_flag(self) -> None:
        ledger = _build_ledger()
        results = validate_consolidation(ledger, 2025)
        assert results["revenue_in_range"] is True
