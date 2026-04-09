"""Tests for generator.model.views — derived financial views for formatters."""

import datetime
import random
from decimal import Decimal

from generator.model.consolidation import (
    ALL_ENTITIES,
    build_income_statement,
    consolidated_trial_balance,
)
from generator.model.gl import Ledger
from generator.model.intercompany import (
    generate_ic_transactions,
    post_ic_loan_principal,
    post_ic_transactions_to_gl,
)
from generator.model.revenue import generate_monthly_revenue, post_revenue_to_gl
from generator.model.views import (
    all_entity_trial_balances,
    entity_trial_balance,
    monthly_pnl,
    monthly_pnl_summary,
)


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


class TestEntityTrialBalance:
    """Per-entity trial balance views."""

    def test_each_entity_has_entries(self) -> None:
        ledger = _build_ledger()
        eoy = datetime.date(2025, 12, 31)
        for entity in ALL_ENTITIES:
            tb = entity_trial_balance(ledger, entity, eoy)
            assert len(tb) > 0, f"Entity {entity} has no GL balances"

    def test_entity_tb_debits_equal_credits(self) -> None:
        """For each entity, sum of all account balances (debit-credit) should be zero."""
        ledger = _build_ledger()
        eoy = datetime.date(2025, 12, 31)
        for entity in ALL_ENTITIES:
            tb = entity_trial_balance(ledger, entity, eoy)
            total = sum(tb.values())
            assert total == 0, (
                f"Entity {entity}: trial balance off by {total} "
                f"(debits != credits)"
            )

    def test_consolidated_equals_sum_of_entities(self) -> None:
        """Consolidated TB should equal the sum of all entity TBs."""
        ledger = _build_ledger()
        eoy = datetime.date(2025, 12, 31)
        entity_tbs = all_entity_trial_balances(ledger, eoy)
        consol_tb = consolidated_trial_balance(ledger, eoy)

        # Sum entity TBs
        summed: dict[str, Decimal] = {}
        for etb in entity_tbs.values():
            for acct, bal in etb.items():
                summed[acct] = summed.get(acct, Decimal(0)) + bal
        # Drop zeros
        summed = {a: b for a, b in summed.items() if b != 0}

        assert summed == consol_tb

    def test_fy25_consolidated_sums_to_zero(self) -> None:
        """FY25 consolidated trial balance debits must equal credits."""
        ledger = _build_ledger()
        eoy = datetime.date(2025, 12, 31)
        tb = consolidated_trial_balance(ledger, eoy)
        total = sum(tb.values())
        assert total == 0, f"Consolidated TB off by {total}"


class TestMonthlyPnL:
    """Monthly P&L breakdown."""

    def test_twelve_months(self) -> None:
        ledger = _build_ledger()
        months = monthly_pnl(ledger, 2025)
        assert len(months) == 12

    def test_monthly_aggregates_match_annual_is(self) -> None:
        """Sum of monthly P&L across 12 months must match annual IS detail."""
        ledger = _build_ledger()
        months = monthly_pnl(ledger, 2025)
        annual_is = build_income_statement(ledger, 2025)

        # Sum monthly activity per account
        annual_from_monthly: dict[str, Decimal] = {}
        for month_activity in months:
            for acct, bal in month_activity.items():
                annual_from_monthly[acct] = (
                    annual_from_monthly.get(acct, Decimal(0)) + bal
                )

        # Compare against IS detail — filter to IS accounts only (4xxx–8xxx)
        # since the IS detail dict may include non-IS accounts from _year_activity
        is_prefixes = {"4", "5", "6", "7", "8"}
        annual_is_accounts = {
            acct: bal
            for acct, bal in annual_is.detail.items()
            if acct[0] in is_prefixes and not acct.startswith("9")
        }

        for acct, annual_bal in annual_is_accounts.items():
            monthly_sum = annual_from_monthly.get(acct, Decimal(0))
            assert monthly_sum == annual_bal, (
                f"Account {acct}: monthly sum {monthly_sum} != "
                f"annual IS {annual_bal}"
            )

        # And the reverse: no extra accounts in monthly that aren't in annual
        for acct in annual_from_monthly:
            assert acct in annual_is_accounts, (
                f"Account {acct} in monthly P&L but not in annual IS"
            )

    def test_only_is_accounts(self) -> None:
        """Monthly P&L should only contain IS accounts (4xxx–8xxx), no BS or IC."""
        ledger = _build_ledger()
        months = monthly_pnl(ledger, 2025)
        for i, month_activity in enumerate(months):
            for acct in month_activity:
                prefix = acct[0]
                assert prefix in {"4", "5", "6", "7", "8"}, (
                    f"Month {i + 1}: non-IS account {acct} in P&L"
                )
                assert not acct.startswith("9"), (
                    f"Month {i + 1}: IC account {acct} in P&L"
                )


class TestMonthlyPnLSummary:
    """Monthly P&L summary by category."""

    def test_twelve_months(self) -> None:
        ledger = _build_ledger()
        summaries = monthly_pnl_summary(ledger, 2025)
        assert len(summaries) == 12

    def test_revenue_positive_every_month(self) -> None:
        ledger = _build_ledger()
        summaries = monthly_pnl_summary(ledger, 2025)
        for i, s in enumerate(summaries):
            assert s["revenue"] > 0, f"Month {i + 1}: revenue <= 0"

    def test_annual_revenue_matches_is(self) -> None:
        """Sum of monthly revenue must match annual IS total_revenue."""
        ledger = _build_ledger()
        summaries = monthly_pnl_summary(ledger, 2025)
        annual_is = build_income_statement(ledger, 2025)
        monthly_revenue = sum(s["revenue"] for s in summaries)
        assert monthly_revenue == annual_is.total_revenue, (
            f"Monthly revenue {monthly_revenue} != annual {annual_is.total_revenue}"
        )

    def test_annual_net_income_matches_is(self) -> None:
        """Sum of monthly net income must match annual IS net_income."""
        ledger = _build_ledger()
        summaries = monthly_pnl_summary(ledger, 2025)
        annual_is = build_income_statement(ledger, 2025)
        monthly_ni = sum(s["net_income"] for s in summaries)
        assert monthly_ni == annual_is.net_income, (
            f"Monthly NI {monthly_ni} != annual {annual_is.net_income}"
        )

    def test_math_consistency(self) -> None:
        """Each month's derived fields must be internally consistent."""
        ledger = _build_ledger()
        summaries = monthly_pnl_summary(ledger, 2025)
        for i, s in enumerate(summaries):
            assert s["gross_profit"] == s["revenue"] - s["cogs"], f"Month {i + 1}"
            assert s["operating_income"] == s["gross_profit"] - s["opex"], f"Month {i + 1}"
            assert s["pre_tax_income"] == s["operating_income"] - s["other"], f"Month {i + 1}"
            assert s["net_income"] == s["pre_tax_income"] - s["tax"], f"Month {i + 1}"
