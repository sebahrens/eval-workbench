"""Cascade Industries consolidation engine + IC eliminations (§1.1–1.3).

Consolidates the four entities (CI, PC, AM, DS) into a single consolidated
view by summing all entity balances and eliminating intercompany (9xxx)
accounts.  Provides:

1. Consolidated trial balance (all accounts, all entities summed, IC zeroed).
2. Balance sheet view (assets, liabilities, equity) with A = L + E check.
3. Income statement view (revenue, COGS, opex, other, tax).
4. Cash flow statement (indirect method, derived from BS changes).

Acceptance criteria (from the bead):
- Consolidated revenue in $198–202M range (FY2025).
- Balance sheet balances for all 3 years (assets = liabilities + equity).
- All IC (9xxx) accounts zero out on consolidation.
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass, field
from decimal import Decimal

from generator.model.coa import ACCOUNTS_BY_NUMBER, AccountType
from generator.model.gl import Ledger

# ── Entity codes ──────────────────────────────────────────────────────────────

ALL_ENTITIES = ("CI", "PC", "AM", "DS")

# Account prefix → classification for financial statements
_BS_PREFIXES = {"1", "2", "3"}  # Assets, Liabilities, Equity
_IS_PREFIXES = {"4", "5", "6", "7", "8"}  # Revenue through Tax
_IC_PREFIX = "9"


# ── Consolidated Trial Balance ────────────────────────────────────────────────


def consolidated_trial_balance(
    ledger: Ledger,
    as_of_date: datetime.date,
) -> dict[str, Decimal]:
    """Sum balances across all entities for each account, as of a date.

    Returns {account_number: net_balance} for all accounts with non-zero
    balances.  IC (9xxx) accounts are included and should sum to zero.
    """
    combined: dict[str, Decimal] = {}
    for entity in ALL_ENTITIES:
        entity_bals = ledger.balance_by_account(entity, as_of_date=as_of_date)
        for acct, bal in entity_bals.items():
            combined[acct] = combined.get(acct, Decimal(0)) + bal

    return {acct: bal for acct, bal in sorted(combined.items()) if bal != 0}


def consolidated_trial_balance_eliminated(
    ledger: Ledger,
    as_of_date: datetime.date,
) -> dict[str, Decimal]:
    """Consolidated trial balance with IC accounts eliminated.

    Drops all 9xxx accounts.  The caller should verify they sum to zero
    before using this.
    """
    tb = consolidated_trial_balance(ledger, as_of_date)
    return {acct: bal for acct, bal in tb.items() if not acct.startswith(_IC_PREFIX)}


# ── IC Elimination Verification ───────────────────────────────────────────────


def verify_ic_elimination(
    ledger: Ledger,
    as_of_date: datetime.date,
) -> tuple[bool, Decimal, dict[str, Decimal]]:
    """Check that all 9xxx accounts net to zero on consolidation.

    Returns (is_zero, total_imbalance, per_account_balances).
    """
    tb = consolidated_trial_balance(ledger, as_of_date)
    ic_balances = {acct: bal for acct, bal in tb.items() if acct.startswith(_IC_PREFIX)}
    total = sum(ic_balances.values(), Decimal(0))
    return total == 0, total, ic_balances


# ── Financial Statement Views ────────────────────────────────────────────────


@dataclass
class BalanceSheet:
    """Consolidated balance sheet at a point in time."""

    as_of_date: datetime.date
    total_assets: Decimal = Decimal(0)
    total_liabilities: Decimal = Decimal(0)
    total_equity: Decimal = Decimal(0)
    retained_earnings_plugin: Decimal = Decimal(0)
    assets: dict[str, Decimal] = field(default_factory=dict)
    liabilities: dict[str, Decimal] = field(default_factory=dict)
    equity: dict[str, Decimal] = field(default_factory=dict)

    @property
    def is_balanced(self) -> bool:
        return self.total_assets == self.total_liabilities + self.total_equity


@dataclass
class IncomeStatement:
    """Consolidated income statement for a period."""

    year: int
    total_revenue: Decimal = Decimal(0)
    total_cogs: Decimal = Decimal(0)
    gross_profit: Decimal = Decimal(0)
    total_opex: Decimal = Decimal(0)
    operating_income: Decimal = Decimal(0)
    total_other: Decimal = Decimal(0)
    pre_tax_income: Decimal = Decimal(0)
    total_tax: Decimal = Decimal(0)
    net_income: Decimal = Decimal(0)
    detail: dict[str, Decimal] = field(default_factory=dict)


def build_income_statement(
    ledger: Ledger,
    year: int,
) -> IncomeStatement:
    """Build consolidated income statement for a fiscal year.

    Revenue accounts (4xxx) have credit-normal balances, so their
    balance_by_account value is negative (debit - credit).  We negate
    to get positive revenue.  COGS/opex (5xxx, 6xxx) are debit-normal.
    """
    # Get year's activity: filter entries to the year, then compute balances
    # from just those entries.
    year_balances = _year_activity(ledger, year)

    stmt = IncomeStatement(year=year)

    for acct, bal in sorted(year_balances.items()):
        if acct.startswith(_IC_PREFIX):
            continue  # Eliminate IC

        acct_info = ACCOUNTS_BY_NUMBER.get(acct)
        if acct_info is None:
            continue

        stmt.detail[acct] = bal

        if acct_info.account_type == AccountType.REVENUE:
            # Revenue is credit-normal → balance_by_account returns negative
            stmt.total_revenue += -bal
        elif acct_info.account_type == AccountType.COGS:
            stmt.total_cogs += bal
        elif acct_info.account_type == AccountType.OPEX:
            stmt.total_opex += bal
        elif acct_info.account_type == AccountType.OTHER:
            stmt.total_other += bal
        elif acct_info.account_type == AccountType.TAX:
            stmt.total_tax += bal

    stmt.gross_profit = stmt.total_revenue - stmt.total_cogs
    stmt.operating_income = stmt.gross_profit - stmt.total_opex
    stmt.pre_tax_income = stmt.operating_income - stmt.total_other
    stmt.net_income = stmt.pre_tax_income - stmt.total_tax

    return stmt


def build_balance_sheet(
    ledger: Ledger,
    as_of_date: datetime.date,
) -> BalanceSheet:
    """Build consolidated balance sheet as of a given date.

    Assets (1xxx) are debit-normal → positive balance = asset.
    Liabilities (2xxx) are credit-normal → negative balance = liability.
    Equity (3xxx) is credit-normal → negative balance = equity.

    We flip the sign on liabilities and equity so they show as positive
    values in the BalanceSheet dataclass, matching A = L + E convention.

    Retained earnings (3100) is computed as the plug:
      RE = cumulative net income from inception through as_of_date.
    This is implicit in the GL — IS accounts accumulate into RE.
    """
    tb = consolidated_trial_balance_eliminated(ledger, as_of_date)

    bs = BalanceSheet(as_of_date=as_of_date)

    for acct, bal in tb.items():
        acct_info = ACCOUNTS_BY_NUMBER.get(acct)
        if acct_info is None:
            continue

        if acct_info.account_type == AccountType.ASSET:
            bs.assets[acct] = bal
            bs.total_assets += bal
        elif acct_info.account_type == AccountType.LIABILITY:
            # Credit-normal: bal is negative, flip to positive
            bs.liabilities[acct] = -bal
            bs.total_liabilities += -bal
        elif acct_info.account_type == AccountType.EQUITY:
            # Credit-normal: bal is negative, flip to positive
            bs.equity[acct] = -bal
            bs.total_equity += -bal

    # Income statement accounts (4xxx–8xxx) represent undistributed
    # earnings that haven't been closed to retained earnings.
    # Plug them into retained earnings for the BS to balance.
    is_accounts_net = Decimal(0)
    for acct, bal in tb.items():
        prefix = acct[0]
        if prefix in _IS_PREFIXES:
            # Revenue (credit-normal) → negative bal → contributes positively to RE
            # Expenses (debit-normal) → positive bal → reduces RE
            is_accounts_net += bal

    # is_accounts_net = sum of (debits - credits) for IS accounts
    # Net income = -(is_accounts_net) because revenue credits > expense debits
    # This needs to be ADDED to equity
    bs.retained_earnings_plugin = -is_accounts_net
    bs.total_equity += -is_accounts_net

    return bs


# ── Cash Flow (Indirect Method) ──────────────────────────────────────────────


@dataclass
class CashFlowStatement:
    """Consolidated cash flow statement (indirect method) for a fiscal year."""

    year: int
    net_income: Decimal = Decimal(0)

    # Operating adjustments
    depreciation_amortization: Decimal = Decimal(0)
    changes_in_working_capital: Decimal = Decimal(0)
    cash_from_operations: Decimal = Decimal(0)

    # Investing
    capex: Decimal = Decimal(0)
    cash_from_investing: Decimal = Decimal(0)

    # Financing
    debt_changes: Decimal = Decimal(0)
    cash_from_financing: Decimal = Decimal(0)

    # Net change
    net_change_in_cash: Decimal = Decimal(0)
    beginning_cash: Decimal = Decimal(0)
    ending_cash: Decimal = Decimal(0)

    # Working capital detail
    working_capital_detail: dict[str, Decimal] = field(default_factory=dict)


def build_cash_flow(
    ledger: Ledger,
    year: int,
) -> CashFlowStatement:
    """Build consolidated cash flow (indirect method) from BS changes.

    Compares consolidated balance sheets at year-end vs prior year-end
    and derives cash flow categories from account changes.
    """
    current_eoy = datetime.date(year, 12, 31)
    prior_eoy = datetime.date(year - 1, 12, 31)

    current_tb = consolidated_trial_balance_eliminated(ledger, current_eoy)
    prior_tb = consolidated_trial_balance_eliminated(ledger, prior_eoy)

    # Get net income for the year
    is_stmt = build_income_statement(ledger, year)

    cf = CashFlowStatement(year=year, net_income=is_stmt.net_income)

    # ── Operating activities (indirect) ─────────────────────────────
    # Start with net income, add back non-cash charges, adjust for WC changes

    # Depreciation & amortization (non-cash charges added back)
    year_activity = _year_activity(ledger, year)
    dep_accounts = {"6210", "6220", "6330"}  # Depreciation, amortization, R&D equip dep
    for acct in dep_accounts:
        cf.depreciation_amortization += year_activity.get(acct, Decimal(0))

    # Working capital changes: ΔAR, ΔInventory, ΔPrepaid, ΔAP, ΔAccrued
    # Change = current - prior; for assets, an increase uses cash (negative)
    # For liabilities, an increase provides cash (positive)
    wc_asset_accounts = {
        "1100": "Accounts Receivable",
        "1150": "Allowance for Doubtful Accounts",
        "1200": "Inventory — Raw Materials",
        "1210": "Inventory — Work in Process",
        "1220": "Inventory — Finished Goods",
        "1230": "Inventory Reserve",
        "1300": "Prepaid Insurance",
        "1310": "Prepaid Rent",
        "1320": "Prepaid Other",
        "1350": "Other Current Assets",
    }
    wc_liability_accounts = {
        "2010": "Accounts Payable",
        "2020": "Accrued Expenses",
        "2030": "Accrued Payroll",
        "2040": "Accrued Bonuses",
        "2050": "Warranty Reserve",
        "2060": "Sales Tax Payable",
        "2070": "Income Tax Payable — Federal",
        "2075": "Income Tax Payable — State",
        "2080": "Payroll Tax Payable",
        "2120": "Customer Deposits",
        "2130": "Deferred Revenue",
        "2150": "Other Current Liabilities",
    }

    total_wc_change = Decimal(0)

    for acct, label in wc_asset_accounts.items():
        current_bal = current_tb.get(acct, Decimal(0))
        prior_bal = prior_tb.get(acct, Decimal(0))
        change = current_bal - prior_bal
        # Asset increase = cash used (negative for cash flow)
        cash_impact = -change
        if cash_impact != 0:
            cf.working_capital_detail[acct] = cash_impact
            total_wc_change += cash_impact

    for acct, label in wc_liability_accounts.items():
        current_bal = current_tb.get(acct, Decimal(0))
        prior_bal = prior_tb.get(acct, Decimal(0))
        change = current_bal - prior_bal
        # Liability balances are negative in TB (credit-normal)
        # An increase (more negative) = source of cash (positive for CF)
        cash_impact = -change  # Flip: more negative liability → positive cash
        if cash_impact != 0:
            cf.working_capital_detail[acct] = cash_impact
            total_wc_change += cash_impact

    cf.changes_in_working_capital = total_wc_change
    cf.cash_from_operations = (
        cf.net_income + cf.depreciation_amortization + cf.changes_in_working_capital
    )

    # ── Investing activities ────────────────────────────────────────
    # CapEx = change in gross PP&E (1400–1460) + change in intangibles (1700–1710)
    capex_accounts = {"1400", "1410", "1420", "1430", "1440", "1450", "1460",
                      "1600", "1610", "1700", "1710"}
    for acct in sorted(capex_accounts):
        current_bal = current_tb.get(acct, Decimal(0))
        prior_bal = prior_tb.get(acct, Decimal(0))
        change = current_bal - prior_bal
        cf.capex -= change  # Increase in assets = cash outflow (negative)

    cf.cash_from_investing = cf.capex

    # ── Financing activities ────────────────────────────────────────
    # Change in debt (2100–2210), equity changes (3010–3030, 3300)
    financing_accounts = {"2100", "2110", "2200", "2210",
                          "2300", "2310", "2320", "2330",
                          "3010", "3020", "3030", "3300"}
    for acct in sorted(financing_accounts):
        current_bal = current_tb.get(acct, Decimal(0))
        prior_bal = prior_tb.get(acct, Decimal(0))
        change = current_bal - prior_bal
        # Liability increase (more negative) = cash inflow
        # Equity increase (more negative for credit-normal) = cash inflow
        # Treasury stock (3030, debit-normal) increase = cash outflow
        acct_info = ACCOUNTS_BY_NUMBER.get(acct)
        if acct_info and acct_info.account_type == AccountType.EQUITY and acct == "3030":
            # Treasury stock is debit-normal
            cf.debt_changes -= change
        else:
            cf.debt_changes -= change  # Credit-normal: more negative = inflow

    cf.cash_from_financing = cf.debt_changes

    # ── Net change & balances ───────────────────────────────────────
    cf.net_change_in_cash = (
        cf.cash_from_operations + cf.cash_from_investing + cf.cash_from_financing
    )

    # Cash accounts: 1010, 1020, 1030, 1050
    cash_accounts = {"1010", "1020", "1030", "1050"}
    cf.beginning_cash = sum(
        prior_tb.get(acct, Decimal(0)) for acct in cash_accounts
    )
    cf.ending_cash = sum(
        current_tb.get(acct, Decimal(0)) for acct in cash_accounts
    )

    return cf


# ── Helpers ───────────────────────────────────────────────────────────────────


def _year_activity(
    ledger: Ledger,
    year: int,
) -> dict[str, Decimal]:
    """Compute consolidated account activity for a single fiscal year.

    Returns {account: net_debit_minus_credit} for all IS-type accounts,
    summed across all entities, for entries dated within the year.
    """
    start = datetime.date(year, 1, 1)
    end = datetime.date(year, 12, 31)
    entries = ledger.filter_by_date_range(start, end)

    activity: dict[str, Decimal] = {}
    for entry in entries:
        if entry.entity_code not in ALL_ENTITIES:
            continue
        for line in entry.lines:
            activity[line.account] = activity.get(line.account, Decimal(0)) + (
                line.debit - line.credit
            )

    return activity


def validate_consolidation(
    ledger: Ledger,
    year: int,
) -> dict[str, str | Decimal | bool]:
    """Run all consolidation acceptance checks for a given year.

    Returns a dict of check results for reporting.
    """
    eoy = datetime.date(year, 12, 31)

    # 1. IC elimination
    ic_zero, ic_total, ic_detail = verify_ic_elimination(ledger, eoy)

    # 2. Balance sheet balance
    bs = build_balance_sheet(ledger, eoy)

    # 3. Consolidated revenue range (FY2025 only)
    is_stmt = build_income_statement(ledger, year)

    results: dict[str, str | Decimal | bool] = {
        "year": year,
        "ic_accounts_zero": ic_zero,
        "ic_total_imbalance": ic_total,
        "bs_balanced": bs.is_balanced,
        "total_assets": bs.total_assets,
        "total_liabilities": bs.total_liabilities,
        "total_equity": bs.total_equity,
        "consolidated_revenue": is_stmt.total_revenue,
        "net_income": is_stmt.net_income,
    }

    if year == 2025:
        results["revenue_in_range"] = (
            Decimal("198_000_000") <= is_stmt.total_revenue <= Decimal("202_000_000")
        )

    return results
