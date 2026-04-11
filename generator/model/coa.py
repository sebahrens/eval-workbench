"""Cascade Industries chart of accounts (~120 accounts, §1.5 of prompt.md).

Provides the canonical Account dataclass and a pre-built list/dict of all
accounts. Account numbers follow the convention:

    1xxx  Assets
    2xxx  Liabilities
    3xxx  Equity
    4xxx  Revenue
    5xxx  Cost of Goods Sold
    6xxx  Operating Expenses (SG&A)
    7xxx  Other Income / Expense
    8xxx  Tax accounts
    9xxx  Intercompany

Granular R&D expense accounts support TC-08 (R&D credit study).
Separate lease liability / ROU asset accounts support TC-04 (ASC 842).
Intercompany accounts come in matched pairs (receivable ↔ payable, etc.).
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum


class AccountType(Enum):
    ASSET = "Asset"
    LIABILITY = "Liability"
    EQUITY = "Equity"
    REVENUE = "Revenue"
    COGS = "Cost of Goods Sold"
    OPEX = "Operating Expense"
    OTHER = "Other Income/Expense"
    TAX = "Tax"
    INTERCOMPANY = "Intercompany"


class NormalBalance(Enum):
    DEBIT = "Debit"
    CREDIT = "Credit"


# Map account type → expected normal balance
_TYPE_NORMAL: dict[AccountType, NormalBalance] = {
    AccountType.ASSET: NormalBalance.DEBIT,
    AccountType.LIABILITY: NormalBalance.CREDIT,
    AccountType.EQUITY: NormalBalance.CREDIT,
    AccountType.REVENUE: NormalBalance.CREDIT,
    AccountType.COGS: NormalBalance.DEBIT,
    AccountType.OPEX: NormalBalance.DEBIT,
    AccountType.OTHER: NormalBalance.DEBIT,  # default; credits override below
    AccountType.TAX: NormalBalance.DEBIT,
    AccountType.INTERCOMPANY: NormalBalance.DEBIT,  # varies; set per account
}

# Prefix → AccountType
_PREFIX_TYPE: dict[str, AccountType] = {
    "1": AccountType.ASSET,
    "2": AccountType.LIABILITY,
    "3": AccountType.EQUITY,
    "4": AccountType.REVENUE,
    "5": AccountType.COGS,
    "6": AccountType.OPEX,
    "7": AccountType.OTHER,
    "8": AccountType.TAX,
    "9": AccountType.INTERCOMPANY,
}


@dataclass(frozen=True)
class Account:
    """A single GL account in the Cascade Industries chart of accounts."""

    number: str  # 4-digit string, e.g. "1010"
    name: str
    account_type: AccountType
    normal_balance: NormalBalance
    description: str

    @property
    def prefix(self) -> str:
        return self.number[0]


def _acct(number: str, name: str, description: str, *, credit: bool = False) -> Account:
    """Convenience builder — derives type from prefix, uses default normal balance."""
    atype = _PREFIX_TYPE[number[0]]
    normal = NormalBalance.CREDIT if credit else _TYPE_NORMAL[atype]
    return Account(
        number=number,
        name=name,
        account_type=atype,
        normal_balance=normal,
        description=description,
    )


# ── 1xxx  Assets ──────────────────────────────────────────────────────────────
_ASSETS = [
    _acct("1010", "Cash — Operating", "Primary operating bank accounts"),
    _acct("1015", "Cash — Collections Clearing", "Customer collections pending sweep to operating account"),
    _acct("1020", "Cash — Payroll", "Dedicated payroll bank account"),
    _acct("1030", "Cash — Restricted", "Restricted cash (bond covenants)"),
    _acct("1050", "Short-Term Investments", "Money market and short-term securities"),
    _acct("1100", "Accounts Receivable — Trade", "Trade receivables from customers"),
    _acct("1110", "Accounts Receivable — Related Parties", "Receivables from non-consolidated related parties"),
    _acct("1150", "Allowance for Doubtful Accounts", "Contra-asset: estimated uncollectible AR", credit=True),
    _acct("1200", "Inventory — Raw Materials", "Raw materials and components"),
    _acct("1210", "Inventory — Work in Process", "Partially completed goods"),
    _acct("1220", "Inventory — Finished Goods", "Completed goods awaiting sale"),
    _acct("1230", "Inventory Reserve", "Contra-asset: obsolescence and LCM reserve", credit=True),
    _acct("1300", "Prepaid Insurance", "Prepaid insurance premiums"),
    _acct("1310", "Prepaid Rent", "Prepaid lease and rent payments"),
    _acct("1320", "Prepaid Other", "Other prepaid expenses"),
    _acct("1350", "Other Current Assets", "Miscellaneous current assets"),
    _acct("1400", "Property, Plant & Equipment — Land", "Land (non-depreciable)"),
    _acct("1410", "Property, Plant & Equipment — Buildings", "Buildings and leasehold improvements"),
    _acct("1420", "Property, Plant & Equipment — Machinery", "Manufacturing machinery and equipment"),
    _acct("1430", "Property, Plant & Equipment — Vehicles", "Fleet and delivery vehicles"),
    _acct("1440", "Property, Plant & Equipment — Furniture & Fixtures", "Office furniture and fixtures"),
    _acct("1450", "Property, Plant & Equipment — Computer Equipment", "Servers, workstations, and IT hardware"),
    _acct("1460", "Property, Plant & Equipment — Construction in Progress", "CIP for assets not yet placed in service"),
    _acct("1500", "Accumulated Depreciation — Buildings", "Contra-asset for buildings", credit=True),
    _acct("1510", "Accumulated Depreciation — Machinery", "Contra-asset for machinery", credit=True),
    _acct("1520", "Accumulated Depreciation — Vehicles", "Contra-asset for vehicles", credit=True),
    _acct("1530", "Accumulated Depreciation — Furniture & Fixtures", "Contra-asset for FF&E", credit=True),
    _acct("1540", "Accumulated Depreciation — Computer Equipment", "Contra-asset for IT hardware", credit=True),
    # ASC 842 — Right-of-Use Assets (TC-04)
    _acct("1600", "ROU Asset — Operating Leases", "Right-of-use assets under ASC 842 operating leases"),
    _acct("1610", "ROU Asset — Finance Leases", "Right-of-use assets under ASC 842 finance leases"),
    _acct(
        "1620", "Accumulated Amortization — ROU Operating",
        "Contra-asset for ROU operating lease assets", credit=True,
    ),
    _acct("1630", "Accumulated Amortization — ROU Finance", "Contra-asset for ROU finance lease assets", credit=True),
    _acct("1700", "Intangible Assets — Patents", "Capitalized patent costs"),
    _acct("1710", "Intangible Assets — Software", "Capitalized internal-use software"),
    _acct("1720", "Accumulated Amortization — Intangibles", "Contra-asset for intangible amortization", credit=True),
    _acct("1800", "Goodwill", "Goodwill from business combinations"),
    _acct("1900", "Deferred Tax Asset — Current", "Current portion of deferred tax assets"),
    _acct("1910", "Deferred Tax Asset — Non-Current", "Long-term deferred tax assets"),
    _acct("1950", "Other Non-Current Assets", "Miscellaneous long-term assets"),
]

# ── 2xxx  Liabilities ─────────────────────────────────────────────────────────
_LIABILITIES = [
    _acct("2010", "Accounts Payable — Trade", "Trade payables to vendors"),
    _acct("2020", "Accounts Payable — Accrued Expenses", "Accrued but unpaid expenses"),
    _acct("2030", "Accrued Payroll & Benefits", "Accrued salaries, wages, and benefits"),
    _acct("2040", "Accrued Bonuses", "Year-end bonus accrual (temporary difference for TC-06)"),
    _acct("2050", "Accrued Warranty Reserve", "Warranty obligation reserve (temporary difference)"),
    _acct("2060", "Sales Tax Payable", "Collected sales tax awaiting remittance"),
    _acct("2070", "Income Tax Payable — Federal", "Federal income tax currently payable"),
    _acct("2075", "Income Tax Payable — State", "State income tax currently payable"),
    _acct("2080", "Payroll Tax Payable", "Employer FICA, FUTA, SUTA payable"),
    _acct("2100", "Short-Term Debt — Line of Credit", "Revolving credit facility"),
    _acct("2110", "Current Portion of Long-Term Debt", "Current maturities of long-term debt"),
    _acct("2120", "Customer Deposits", "Customer advance payments and deposits"),
    _acct("2130", "Deferred Revenue", "Unearned revenue (advance billings)"),
    _acct("2150", "Other Current Liabilities", "Miscellaneous current liabilities"),
    _acct("2200", "Long-Term Debt — Term Loan", "Bank term loan (non-current)"),
    _acct("2210", "Long-Term Debt — Equipment Financing", "Equipment-secured financing"),
    # ASC 842 — Lease Liabilities (TC-04)
    _acct("2300", "Lease Liability — Operating (Current)", "Current portion of ASC 842 operating lease obligations"),
    _acct("2310", "Lease Liability — Operating (Non-Current)", "Non-current operating lease obligations"),
    _acct("2320", "Lease Liability — Finance (Current)", "Current portion of ASC 842 finance lease obligations"),
    _acct("2330", "Lease Liability — Finance (Non-Current)", "Non-current finance lease obligations"),
    _acct("2400", "Deferred Tax Liability — Current", "Current portion of deferred tax liabilities"),
    _acct("2410", "Deferred Tax Liability — Non-Current", "Long-term deferred tax liabilities"),
    _acct("2500", "Bad Debt Reserve", "Allowance for bad debt (liability-side reserve)"),
    _acct("2600", "Other Non-Current Liabilities", "Miscellaneous long-term liabilities"),
]

# ── 3xxx  Equity ──────────────────────────────────────────────────────────────
_EQUITY = [
    _acct("3010", "Common Stock", "Par value of issued common shares"),
    _acct("3020", "Additional Paid-In Capital", "Excess over par from stock issuances"),
    _acct("3030", "Treasury Stock", "Reacquired shares at cost", credit=False),
    _acct("3100", "Retained Earnings", "Accumulated undistributed profits"),
    _acct("3200", "Accumulated Other Comprehensive Income", "OCI items (FX, pensions, hedges)"),
    _acct("3300", "Dividends Declared", "Dividends declared during the period", credit=False),
]

# ── 4xxx  Revenue ─────────────────────────────────────────────────────────────
_REVENUE = [
    _acct("4010", "Revenue — Product Sales", "Revenue from manufactured goods"),
    _acct("4020", "Revenue — Service", "Revenue from engineering and consulting services"),
    _acct("4030", "Revenue — Distribution & Logistics", "Revenue from warehousing and freight"),
    _acct("4040", "Revenue — Contract & Government", "Government and long-term contract revenue"),
    _acct("4100", "Sales Returns & Allowances", "Contra-revenue: returns and credits", credit=False),
    _acct("4110", "Sales Discounts", "Contra-revenue: early-payment discounts", credit=False),
]

# ── 5xxx  Cost of Goods Sold ──────────────────────────────────────────────────
_COGS = [
    _acct("5010", "COGS — Direct Materials", "Raw materials consumed in production"),
    _acct("5020", "COGS — Direct Labor", "Production labor costs"),
    _acct("5030", "COGS — Manufacturing Overhead", "Allocated factory overhead"),
    _acct("5040", "COGS — Freight In", "Inbound freight on materials"),
    _acct("5050", "COGS — Inventory Adjustments", "Inventory shrinkage, write-downs, and scrap"),
]

# ── 6xxx  Operating Expenses ──────────────────────────────────────────────────
_OPEX = [
    _acct("6010", "Salaries & Wages — Administrative", "Admin and corporate staff compensation"),
    _acct("6020", "Salaries & Wages — Sales", "Sales team compensation"),
    _acct("6030", "Employee Benefits", "Health, dental, life, 401(k) match"),
    _acct("6040", "Payroll Taxes", "Employer payroll tax expense"),
    _acct("6050", "Stock-Based Compensation", "Equity compensation expense (permanent difference)"),
    _acct("6100", "Rent Expense — Operating Leases", "ASC 842 operating lease expense (straight-line)"),
    _acct("6110", "Utilities", "Electric, gas, water, and waste services"),
    _acct("6120", "Office Supplies", "General office consumables"),
    _acct("6130", "Telecommunications", "Phone, internet, and data services"),
    _acct("6140", "Insurance — General", "Property, casualty, and general liability premiums"),
    _acct("6150", "Travel & Entertainment", "Business travel and lodging"),
    _acct("6160", "Meals & Entertainment", "Meals expense (50% non-deductible — permanent difference)"),
    _acct("6170", "Professional Fees — Legal", "Legal services"),
    _acct("6180", "Professional Fees — Accounting & Audit", "External audit and advisory fees"),
    _acct("6190", "Professional Fees — Consulting", "Management and IT consulting"),
    _acct("6200", "Advertising & Marketing", "Marketing campaigns and promotional costs"),
    _acct("6210", "Depreciation Expense", "Book depreciation (temporary difference vs. MACRS)"),
    _acct("6220", "Amortization Expense", "Amortization of intangible assets"),
    _acct("6230", "Bad Debt Expense", "Provision for doubtful accounts"),
    _acct("6240", "Repairs & Maintenance", "Facility and equipment repairs"),
    _acct("6250", "Shipping & Delivery", "Outbound freight and delivery costs"),
    _acct("6260", "Software & Subscriptions", "SaaS and software license costs"),
    _acct("6270", "Training & Development", "Employee training and continuing education"),
    _acct("6280", "Fines & Penalties", "Regulatory fines (non-deductible — permanent difference)"),
    # R&D expense accounts — granular for TC-08 (Section 41 credit study)
    _acct("6300", "R&D — Salaries & Wages", "R&D-eligible employee compensation"),
    _acct("6310", "R&D — Contract Research", "Payments to third-party research contractors (65%)"),
    _acct("6320", "R&D — Supplies & Materials", "Lab supplies and prototype materials"),
    _acct("6330", "R&D — Equipment Depreciation", "Depreciation on R&D-dedicated equipment"),
    _acct("6340", "R&D — Software & Tools", "Specialized R&D software and simulation tools"),
    _acct("6350", "R&D — Testing & Certification", "Product testing and regulatory certification"),
    _acct("6360", "R&D — Travel", "Travel for R&D site visits and conferences"),
    _acct("6370", "R&D — Other", "Miscellaneous R&D expenses"),
    _acct("6400", "Warranty Expense", "Warranty claims and service costs"),
]

# ── 7xxx  Other Income / Expense ──────────────────────────────────────────────
_OTHER = [
    _acct("7010", "Interest Income", "Interest on deposits and investments", credit=True),
    _acct("7020", "Interest Expense", "Interest on debt and financing"),
    _acct("7030", "Gain on Sale of Assets", "Gains on disposal of PP&E", credit=True),
    _acct("7040", "Loss on Sale of Assets", "Losses on disposal of PP&E"),
    _acct("7050", "Foreign Exchange Gain/Loss", "Realized and unrealized FX differences"),
    _acct("7060", "Miscellaneous Income", "Other non-operating income", credit=True),
    _acct("7070", "Miscellaneous Expense", "Other non-operating expenses"),
    _acct("7080", "Tax-Exempt Interest Income", "Municipal bond interest (permanent difference)", credit=True),
]

# ── 8xxx  Tax ─────────────────────────────────────────────────────────────────
_TAX = [
    _acct("8010", "Federal Income Tax Expense — Current", "Current year federal income tax"),
    _acct("8020", "State Income Tax Expense — Current", "Current year state income tax"),
    _acct("8030", "Federal Income Tax Expense — Deferred", "Deferred federal income tax"),
    _acct("8040", "State Income Tax Expense — Deferred", "Deferred state income tax"),
    _acct("8050", "R&D Tax Credit", "Section 41 research credit (reduces tax expense)", credit=True),
    _acct("8060", "Other Tax Credits", "Other business tax credits", credit=True),
]

# ── 9xxx  Intercompany ────────────────────────────────────────────────────────
# Matched pairs: each receivable (debit-normal) has a corresponding payable (credit-normal).
_INTERCOMPANY = [
    _acct("9010", "IC Receivable — Precision Components", "Due from Precision Components LLC"),
    _acct("9020", "IC Payable — Precision Components", "Due to Precision Components LLC", credit=True),
    _acct("9030", "IC Receivable — Advanced Materials", "Due from Advanced Materials, Inc."),
    _acct("9040", "IC Payable — Advanced Materials", "Due to Advanced Materials, Inc.", credit=True),
    _acct("9050", "IC Receivable — Distribution Services", "Due from Distribution Services LLC"),
    _acct("9060", "IC Payable — Distribution Services", "Due to Distribution Services LLC", credit=True),
    _acct("9070", "IC Receivable — Parent", "Due from Cascade Industries (parent)"),
    _acct("9080", "IC Payable — Parent", "Due to Cascade Industries (parent)", credit=True),
    _acct("9100", "IC Revenue — Management Fees", "Management fee income from subsidiaries", credit=True),
    _acct("9110", "IC Expense — Management Fees", "Management fee expense to parent"),
    _acct("9120", "IC Revenue — Goods", "Intercompany sales of goods", credit=True),
    _acct("9130", "IC COGS — Goods", "Cost of intercompany goods sold"),
    _acct("9140", "IC Interest Income", "Interest on intercompany loans", credit=True),
    _acct("9150", "IC Interest Expense", "Interest on intercompany borrowings"),
    _acct("9160", "IC Revenue — Services", "Intercompany services revenue (warehouse fees)", credit=True),
    _acct("9170", "IC Expense — Services", "Intercompany services expense (warehouse fees)"),
    _acct("9200", "IC Loan Receivable", "Intercompany loan principal — lender side"),
    _acct("9210", "IC Loan Payable", "Intercompany loan principal — borrower side", credit=True),
]

# ── Master list & lookup ──────────────────────────────────────────────────────

ACCOUNTS: list[Account] = sorted(
    _ASSETS + _LIABILITIES + _EQUITY + _REVENUE + _COGS + _OPEX + _OTHER + _TAX + _INTERCOMPANY,
    key=lambda a: a.number,
)

ACCOUNTS_BY_NUMBER: dict[str, Account] = {a.number: a for a in ACCOUNTS}


def validate_coa() -> list[str]:
    """Return a list of validation errors (empty = valid).

    Checks:
      1. All account numbers are unique 4-digit strings.
      2. Account type matches its numeric prefix.
      3. Every IC receivable (debit) has a matching IC payable (credit) and vice-versa.
    """
    errors: list[str] = []

    # Check uniqueness
    seen: dict[str, str] = {}
    for acct in ACCOUNTS:
        if len(acct.number) != 4 or not acct.number.isdigit():
            errors.append(f"{acct.number}: account number must be a 4-digit string")
        if acct.number in seen:
            errors.append(f"{acct.number}: duplicate ('{seen[acct.number]}' and '{acct.name}')")
        seen[acct.number] = acct.name

        # Prefix → type
        expected_type = _PREFIX_TYPE.get(acct.number[0])
        if expected_type and acct.account_type != expected_type:
            errors.append(
                f"{acct.number} ({acct.name}): type {acct.account_type.value} "
                f"doesn't match prefix (expected {expected_type.value})"
            )

    # IC pair check: for 9xxx, ensure both sides present
    ic_receivable_names = {
        a.name for a in ACCOUNTS if a.number.startswith("9") and a.normal_balance == NormalBalance.DEBIT
    }
    ic_payable_names = {
        a.name for a in ACCOUNTS if a.number.startswith("9") and a.normal_balance == NormalBalance.CREDIT
    }
    # Check entity-specific IC pairs (90xx range)
    for acct in ACCOUNTS:
        if acct.number[:2] == "90" and "Receivable" in acct.name:
            entity_part = acct.name.replace("IC Receivable — ", "")
            payable_name = f"IC Payable — {entity_part}"
            if payable_name not in ic_payable_names:
                errors.append(f"{acct.number} ({acct.name}): missing IC payable counterpart '{payable_name}'")
        elif acct.number[:2] == "90" and "Payable" in acct.name:
            entity_part = acct.name.replace("IC Payable — ", "")
            receivable_name = f"IC Receivable — {entity_part}"
            if receivable_name not in ic_receivable_names:
                errors.append(f"{acct.number} ({acct.name}): missing IC receivable counterpart '{receivable_name}'")

    return errors
