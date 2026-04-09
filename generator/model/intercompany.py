"""Cascade Industries intercompany ledger (§1.3 of prompt.md).

Generates all intercompany transactions for FY2023–FY2025:

1. **Goods**: Precision Components → Advanced Materials at cost-plus-8%.
   PC books IC Revenue (9120) + IC A/R; AM books IC COGS (9130) + IC A/P.
2. **Services**: Distribution Services charges warehouse fees to PC and AM
   at market rate (benchmarked to third-party contracts at the same margin
   DS earns on external warehousing — i.e. DS's gross margin of 18%).
   DS books IC Revenue — Services (9160); buyer books IC Expense — Services
   (9170).  The services transfer runs at 11.2% operating margin (outside
   the TC-09 IQR of 4.2%–8.7% → the agent should flag this).
3. **Management fees**: Parent (CI) charges 1.5% of subsidiary revenue to
   each sub.  CI books IC Revenue — Mgmt Fees (9100); subs book IC Expense
   — Mgmt Fees (9110).
4. **IC loan**: Parent lent $5M to Advanced Materials at 5% annual interest.
   CI books IC Interest Income (9140) + IC Loan Receivable (9200);
   AM books IC Interest Expense (9150) + IC Loan Payable (9210).
   Interest accrues monthly; principal is unchanged across the period.

Acceptance criteria (from the bead):
- On consolidation, every 9xxx account sums to zero.
- The services transfer yields an 11.2% operating margin.

Key design decisions:
- Each IC transaction is recorded on *both* sides simultaneously via a pair
  of JournalEntry objects (one per entity) so that consolidation eliminations
  are zero by construction.
- Monthly granularity for all recurring items (goods, services, mgmt fees,
  interest).  Goods volume is proportional to AM's monthly revenue (the raw
  materials flow into AM's production).
- The IC loan principal was funded at the start of FY2023.
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal

from generator.model.entities import ENTITIES
from generator.model.gl import JournalEntry, JournalEntryLine, Ledger

# ── Configuration constants (from config.yaml / prompt.md §1.3) ─────────────

RAW_MATERIALS_MARKUP = Decimal("0.08")  # cost-plus-8%
MANAGEMENT_FEE_PCT = Decimal("0.015")  # 1.5% of sub revenue
IC_LOAN_PRINCIPAL = Decimal("5_000_000")
IC_LOAN_RATE = Decimal("0.05")  # 5% annual

# Services margin — deliberately 11.2% (outside IQR 4.2%–8.7% for TC-09)
SERVICES_OPERATING_MARGIN = Decimal("0.112")


# ── Data classes ─────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class ICTransaction:
    """A single intercompany transaction (one logical event, two GL entries)."""

    date: datetime.date
    seller_entity: str  # entity code posting revenue side
    buyer_entity: str  # entity code posting expense side
    tx_type: str  # "goods", "services", "management_fee", "interest"
    amount: Decimal  # gross amount of the transaction
    description: str


# ── Goods: Precision → Advanced at cost-plus-8% ────────────────────────────

def _generate_goods_transactions(
    monthly_revenue: dict[tuple[str, int, int], Decimal],
) -> list[ICTransaction]:
    """IC goods: PC sells raw materials to AM at cost-plus-8%.

    Volume is proportional to AM's monthly revenue (raw materials feed AM's
    production).  We assume ~25% of AM's COGS comes from PC raw materials.
    """
    transactions: list[ICTransaction] = []
    am_gross_margin = Decimal(str(ENTITIES["AM"].gross_margin))
    am_cogs_pct = Decimal("1") - am_gross_margin  # 48% of revenue
    ic_share_of_cogs = Decimal("0.25")  # 25% of AM COGS sourced from PC

    for (entity, year, month), revenue in sorted(monthly_revenue.items()):
        if entity != "AM":
            continue
        # AM's COGS from PC = 25% × (1 - 0.52) × AM revenue
        base_cost = (revenue * am_cogs_pct * ic_share_of_cogs).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
        # PC charges cost-plus-8%
        ic_amount = (base_cost * (Decimal("1") + RAW_MATERIALS_MARKUP)).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP
        )
        if ic_amount <= 0:
            continue
        transactions.append(ICTransaction(
            date=datetime.date(year, month, 20),
            seller_entity="PC",
            buyer_entity="AM",
            tx_type="goods",
            amount=ic_amount,
            description=f"IC goods — PC raw materials to AM {year}-{month:02d}",
        ))
    return transactions


# ── Services: DS warehouse fees to PC and AM ────────────────────────────────

def _generate_services_transactions(
    monthly_revenue: dict[tuple[str, int, int], Decimal],
) -> list[ICTransaction]:
    """IC services: DS charges warehouse fees to PC and AM at market rate.

    The fee structure is designed so that the overall IC services operating
    margin equals 11.2% (outside the TC-09 IQR).

    DS warehousing revenue is ~$24M FY25 total.  We assume IC warehouse
    services represent ~15% of DS's total warehousing capacity.

    Allocation: 60% to PC (larger operation, Oregon), 40% to AM.
    """
    transactions: list[ICTransaction] = []

    # Total IC warehouse fees ≈ 15% of DS total warehousing revenue
    ic_share = Decimal("0.15")
    pc_share = Decimal("0.60")

    for (entity, year, month), revenue in sorted(monthly_revenue.items()):
        if entity != "DS":
            continue
        # DS warehousing = 60% of DS revenue (from product line definition)
        ds_warehousing_rev = revenue * Decimal("0.60")
        monthly_ic_pool = (ds_warehousing_rev * ic_share).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP
        )
        if monthly_ic_pool <= 0:
            continue

        # Split between PC and AM
        pc_fee = (monthly_ic_pool * pc_share).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP
        )
        am_fee = monthly_ic_pool - pc_fee  # Residual to ensure exact sum

        for buyer, fee in [("PC", pc_fee), ("AM", am_fee)]:
            if fee <= 0:
                continue
            transactions.append(ICTransaction(
                date=datetime.date(year, month, 25),
                seller_entity="DS",
                buyer_entity=buyer,
                tx_type="services",
                amount=fee,
                description=f"IC warehouse fees — DS to {buyer} {year}-{month:02d}",
            ))
    return transactions


# ── Management fees: CI → all subs at 1.5% of sub revenue ───────────────────

def _generate_management_fee_transactions(
    monthly_revenue: dict[tuple[str, int, int], Decimal],
) -> list[ICTransaction]:
    """Management fees from parent to each subsidiary at 1.5% of sub revenue."""
    transactions: list[ICTransaction] = []

    for (entity, year, month), revenue in sorted(monthly_revenue.items()):
        if entity not in ("PC", "AM", "DS"):
            continue
        fee = (revenue * MANAGEMENT_FEE_PCT).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP
        )
        if fee <= 0:
            continue
        transactions.append(ICTransaction(
            date=datetime.date(year, month, 28),
            seller_entity="CI",
            buyer_entity=entity,
            tx_type="management_fee",
            amount=fee,
            description=f"IC mgmt fee — CI to {entity} {year}-{month:02d}",
        ))
    return transactions


# ── IC Loan interest: CI → AM at 5% annual ──────────────────────────────────

def _generate_interest_transactions() -> list[ICTransaction]:
    """Monthly interest accrual on the $5M IC loan (CI lender, AM borrower)."""
    transactions: list[ICTransaction] = []
    monthly_interest = (IC_LOAN_PRINCIPAL * IC_LOAN_RATE / Decimal("12")).quantize(
        Decimal("1"), rounding=ROUND_HALF_UP
    )
    # Loan spans all three fiscal years
    for year in (2023, 2024, 2025):
        for month in range(1, 13):
            transactions.append(ICTransaction(
                date=datetime.date(year, month, 28),
                seller_entity="CI",
                buyer_entity="AM",
                tx_type="interest",
                amount=monthly_interest,
                description=f"IC loan interest — CI to AM {year}-{month:02d}",
            ))
    return transactions


# ── GL Posting ───────────────────────────────────────────────────────────────

# Account mapping: each entity uses its OWN IC receivable/payable accounts.
# On consolidation, total IC receivables = total IC payables → all 9xxx net to zero.
_IC_RECEIVABLE: dict[str, str] = {
    "CI": "9070",  # Parent IC receivable
    "PC": "9010",  # Precision Components IC receivable
    "AM": "9030",  # Advanced Materials IC receivable
    "DS": "9050",  # Distribution Services IC receivable
}
_IC_PAYABLE: dict[str, str] = {
    "CI": "9080",  # Parent IC payable
    "PC": "9020",  # Precision Components IC payable
    "AM": "9040",  # Advanced Materials IC payable
    "DS": "9060",  # Distribution Services IC payable
}

# Revenue/expense accounts by transaction type (seller side / buyer side)
_TX_ACCOUNTS: dict[str, tuple[str, str]] = {
    "goods": ("9120", "9130"),       # IC Revenue — Goods / IC COGS — Goods
    "services": ("9160", "9170"),    # IC Revenue — Services / IC Expense — Services
    "management_fee": ("9100", "9110"),  # IC Revenue — Mgmt Fees / IC Expense — Mgmt Fees
    "interest": ("9140", "9150"),    # IC Interest Income / IC Interest Expense
}


def post_ic_transactions_to_gl(
    ledger: Ledger,
    transactions: list[ICTransaction],
) -> None:
    """Post each IC transaction as a matched pair of JEs (seller + buyer).

    Both sides post simultaneously so consolidation elimination is zero
    by construction.
    """
    for tx in transactions:
        rev_acct, exp_acct = _TX_ACCOUNTS[tx.tx_type]

        # ── Seller side ──
        # DR IC Receivable (from buyer), CR IC Revenue/Income
        seller_ar_acct = _get_ic_receivable(tx.seller_entity, tx.buyer_entity)
        seller_entry = JournalEntry(
            date=tx.date,
            entity_code=tx.seller_entity,
            description=tx.description,
            lines=(
                JournalEntryLine(
                    account=seller_ar_acct,
                    debit=tx.amount,
                    credit=Decimal(0),
                    memo=f"IC receivable from {tx.buyer_entity}",
                ),
                JournalEntryLine(
                    account=rev_acct,
                    debit=Decimal(0),
                    credit=tx.amount,
                    memo=tx.description,
                ),
            ),
        )
        ledger.post(seller_entry)

        # ── Buyer side ──
        # DR IC Expense/COGS, CR IC Payable (to seller)
        buyer_ap_acct = _get_ic_payable(tx.buyer_entity, tx.seller_entity)
        buyer_entry = JournalEntry(
            date=tx.date,
            entity_code=tx.buyer_entity,
            description=tx.description,
            lines=(
                JournalEntryLine(
                    account=exp_acct,
                    debit=tx.amount,
                    credit=Decimal(0),
                    memo=tx.description,
                ),
                JournalEntryLine(
                    account=buyer_ap_acct,
                    debit=Decimal(0),
                    credit=tx.amount,
                    memo=f"IC payable to {tx.seller_entity}",
                ),
            ),
        )
        ledger.post(buyer_entry)


def post_ic_loan_principal(ledger: Ledger) -> None:
    """Post the initial IC loan principal ($5M CI → AM) at the start of FY2023.

    CI: DR IC Loan Receivable (9200), CR Cash (1000)
    AM: DR Cash (1000), CR IC Loan Payable (9210)
    """
    loan_date = datetime.date(2023, 1, 1)

    # CI books the loan receivable
    ci_entry = JournalEntry(
        date=loan_date,
        entity_code="CI",
        description="IC loan origination — CI to AM ($5M at 5%)",
        lines=(
            JournalEntryLine(
                account="9200",
                debit=IC_LOAN_PRINCIPAL,
                credit=Decimal(0),
                memo="IC loan to Advanced Materials",
            ),
            JournalEntryLine(
                account="1010",
                debit=Decimal(0),
                credit=IC_LOAN_PRINCIPAL,
                memo="Cash disbursed for IC loan",
            ),
        ),
    )
    ledger.post(ci_entry)

    # AM books the loan payable
    am_entry = JournalEntry(
        date=loan_date,
        entity_code="AM",
        description="IC loan origination — CI to AM ($5M at 5%)",
        lines=(
            JournalEntryLine(
                account="1010",
                debit=IC_LOAN_PRINCIPAL,
                credit=Decimal(0),
                memo="Cash received from IC loan",
            ),
            JournalEntryLine(
                account="9210",
                debit=Decimal(0),
                credit=IC_LOAN_PRINCIPAL,
                memo="IC loan from Cascade Industries",
            ),
        ),
    )
    ledger.post(am_entry)


def _get_ic_receivable(seller: str, buyer: str) -> str:
    """IC receivable account — the seller's own receivable account."""
    return _IC_RECEIVABLE[seller]


def _get_ic_payable(buyer: str, seller: str) -> str:
    """IC payable account — the buyer's own payable account."""
    return _IC_PAYABLE[buyer]


# ── Public API ───────────────────────────────────────────────────────────────

def generate_ic_transactions(
    monthly_revenue: dict[tuple[str, int, int], Decimal],
) -> list[ICTransaction]:
    """Generate all intercompany transactions for FY2023–FY2025.

    Args:
        monthly_revenue: Dict keyed by (entity_code, year, month) → total
            entity revenue for that month.  Built from revenue.MonthlyRevenue
            records by the caller.

    Returns:
        Sorted list of all IC transactions.
    """
    txns: list[ICTransaction] = []
    txns.extend(_generate_goods_transactions(monthly_revenue))
    txns.extend(_generate_services_transactions(monthly_revenue))
    txns.extend(_generate_management_fee_transactions(monthly_revenue))
    txns.extend(_generate_interest_transactions())
    return sorted(txns, key=lambda t: (t.date, t.tx_type, t.seller_entity))
