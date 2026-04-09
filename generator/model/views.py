"""Derived financial views for Phase 2 formatters (§1.1–1.3 of prompt.md).

Pure functions producing:
- Trial balance (entity × account × period)
- Balance sheet (consolidated)
- Income statement (consolidated)
- Cash flow statement (indirect method, consolidated)
- Monthly P&L by account (consolidated, IC-eliminated)

All derived from the GL — never hand-entered.  These are the inputs that
Phase 2 TC formatters consume.

Re-exports consolidated views from consolidation.py where they already exist,
and adds entity-level and monthly breakdowns that consolidation doesn't provide.
"""

from __future__ import annotations

import datetime
from decimal import Decimal

from generator.model.coa import ACCOUNTS_BY_NUMBER, AccountType
from generator.model.consolidation import (
    ALL_ENTITIES,
    BalanceSheet,
    CashFlowStatement,
    IncomeStatement,
    build_balance_sheet,
    build_cash_flow,
    build_income_statement,
    consolidated_trial_balance,
    consolidated_trial_balance_eliminated,
)
from generator.model.gl import Ledger

# Re-export consolidated views so formatters import from one place
__all__ = [
    "BalanceSheet",
    "CashFlowStatement",
    "IncomeStatement",
    "build_balance_sheet",
    "build_cash_flow",
    "build_income_statement",
    "consolidated_trial_balance",
    "consolidated_trial_balance_eliminated",
    "entity_trial_balance",
    "monthly_pnl",
]


# ── Entity-level trial balance ───────────────────────────────────────────────


def entity_trial_balance(
    ledger: Ledger,
    entity_code: str,
    as_of_date: datetime.date,
) -> dict[str, Decimal]:
    """Trial balance for a single entity as of a date.

    Returns {account_number: net_balance} (debit-minus-credit convention).
    Only accounts with non-zero balances are included.
    """
    return ledger.balance_by_account(entity_code, as_of_date=as_of_date)


def all_entity_trial_balances(
    ledger: Ledger,
    as_of_date: datetime.date,
) -> dict[str, dict[str, Decimal]]:
    """Trial balances for every entity, keyed by entity code.

    Returns {entity_code: {account_number: net_balance}}.
    """
    return {
        entity: entity_trial_balance(ledger, entity, as_of_date)
        for entity in ALL_ENTITIES
    }


# ── Monthly P&L ─────────────────────────────────────────────────────────────


def monthly_pnl(
    ledger: Ledger,
    year: int,
) -> list[dict[str, Decimal]]:
    """Monthly P&L by account for a fiscal year (consolidated, IC-eliminated).

    Returns a list of 12 dicts (index 0 = January, 11 = December).
    Each dict maps account_number → net activity (debit-minus-credit) for
    IS-type accounts (4xxx–8xxx) excluding IC (9xxx).

    Sign convention follows the GL: revenue accounts are negative
    (credit-normal), expense accounts are positive (debit-normal).
    """
    months: list[dict[str, Decimal]] = []

    for month in range(1, 13):
        start = datetime.date(year, month, 1)
        # End of month: day 28–31
        if month == 12:
            end = datetime.date(year, 12, 31)
        else:
            end = datetime.date(year, month + 1, 1) - datetime.timedelta(days=1)

        entries = ledger.filter_by_date_range(start, end)

        activity: dict[str, Decimal] = {}
        for entry in entries:
            if entry.entity_code not in ALL_ENTITIES:
                continue
            for line in entry.lines:
                acct = line.account
                # Skip non-IS accounts and IC accounts
                if acct[0] not in {"4", "5", "6", "7", "8"} or acct[0] == "9":
                    continue
                activity[acct] = activity.get(acct, Decimal(0)) + (
                    line.debit - line.credit
                )

        months.append(activity)

    return months


def monthly_pnl_summary(
    ledger: Ledger,
    year: int,
) -> list[dict[str, Decimal]]:
    """Monthly P&L summarised by category (revenue, cogs, opex, other, tax).

    Returns a list of 12 dicts with keys:
      revenue, cogs, gross_profit, opex, operating_income,
      other, pre_tax_income, tax, net_income

    Revenue is presented as a positive number (sign-flipped from GL convention).
    """
    raw = monthly_pnl(ledger, year)
    summaries: list[dict[str, Decimal]] = []

    for month_activity in raw:
        s: dict[str, Decimal] = {
            "revenue": Decimal(0),
            "cogs": Decimal(0),
            "opex": Decimal(0),
            "other": Decimal(0),
            "tax": Decimal(0),
        }

        for acct, bal in month_activity.items():
            acct_info = ACCOUNTS_BY_NUMBER.get(acct)
            if acct_info is None:
                continue

            if acct_info.account_type == AccountType.REVENUE:
                s["revenue"] += -bal  # credit-normal → flip to positive
            elif acct_info.account_type == AccountType.COGS:
                s["cogs"] += bal
            elif acct_info.account_type == AccountType.OPEX:
                s["opex"] += bal
            elif acct_info.account_type == AccountType.OTHER:
                s["other"] += bal
            elif acct_info.account_type == AccountType.TAX:
                s["tax"] += bal

        s["gross_profit"] = s["revenue"] - s["cogs"]
        s["operating_income"] = s["gross_profit"] - s["opex"]
        s["pre_tax_income"] = s["operating_income"] - s["other"]
        s["net_income"] = s["pre_tax_income"] - s["tax"]

        summaries.append(s)

    return summaries
