"""Cascade Industries accounts receivable subledger + aging.

Generates customer-level AR aging (current/30/60/90/120+ buckets) derived from
the revenue model.  Posts cash collections to GL so account 1100 net balance
equals the AR aging total at each year-end.  Computes bad debt allowance based
on aging reserve rates and posts reclassification entries to move the reserve
from Accrued Expenses (2020) to Allowance for Doubtful Accounts (1150).

Key constraints from the spec:
- Top customer = 18 % of consolidated revenue (TC-11 QofE requirement).
- DSO varies by customer (30–120 days), producing realistic aging distribution.
- AR aging totals must equal GL 1100 balance (acceptance criterion).
- Bad debt reserve rates: Current 1 %, 30-day 3 %, 60-day 10 %, 90-day 25 %,
  120+ 50 %.

Feeds: TC-05 (AR memo), TC-14 (cash flow), TC-11 (QofE customer concentration).
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal

from generator.model.gl import JournalEntry, JournalEntryLine, Ledger
from generator.model.revenue import MonthlyRevenue

# ── Customer definitions ─────────────────────────────────────────────────────
# Hardcoded for determinism.  Top customer Acme Manufacturing =
# 37.9 % of PC ($95 M) = $36.005 M ≈ 18 % of consolidated $200 M.

@dataclass(frozen=True)
class Customer:
    """A trade customer in the AR subledger."""

    id: str
    name: str
    entity_code: str
    revenue_share: Decimal  # fraction of entity annual revenue
    dso: int  # days sales outstanding


# Revenue shares sum to 1.000 per entity.
CUSTOMERS: tuple[Customer, ...] = (
    # ── Precision Components ($95 M) ──────────────────────────────────────
    Customer("CUST-001", "Acme Manufacturing Corp", "PC", Decimal("0.379"), 45),
    Customer("CUST-002", "Northwest Precision Industries", "PC", Decimal("0.150"), 35),
    Customer("CUST-003", "Columbia River Works", "PC", Decimal("0.120"), 40),
    Customer("CUST-004", "Pacific Rim Components", "PC", Decimal("0.095"), 50),
    Customer("CUST-005", "Cascade Defense Systems", "PC", Decimal("0.075"), 30),
    Customer("CUST-006", "Mountain View Fabrication", "PC", Decimal("0.060"), 55),
    Customer("CUST-007", "Summit Industrial Group", "PC", Decimal("0.050"), 35),
    Customer("CUST-008", "Valley Tool & Die", "PC", Decimal("0.040"), 40),
    Customer("CUST-009", "Heritage Machining Co", "PC", Decimal("0.020"), 90),
    Customer("CUST-010", "Westlake Assembly Inc", "PC", Decimal("0.011"), 120),
    # ── Advanced Materials ($65 M) ────────────────────────────────────────
    Customer("CUST-011", "TechAlloy Systems", "AM", Decimal("0.220"), 40),
    Customer("CUST-012", "Quantum Materials Inc", "AM", Decimal("0.180"), 35),
    Customer("CUST-013", "NextGen Composites LLC", "AM", Decimal("0.150"), 50),
    Customer("CUST-014", "Atlas Chemical Supply", "AM", Decimal("0.120"), 45),
    Customer("CUST-015", "Meridian Coatings Corp", "AM", Decimal("0.100"), 60),
    Customer("CUST-016", "Frontier Materials Lab", "AM", Decimal("0.090"), 30),
    Customer("CUST-017", "Apex Surface Technologies", "AM", Decimal("0.080"), 75),
    Customer("CUST-018", "Pacific Polymer Group", "AM", Decimal("0.060"), 45),
    # ── Distribution Services ($40 M) ─────────────────────────────────────
    Customer("CUST-019", "GlobalTrade Logistics", "DS", Decimal("0.200"), 35),
    Customer("CUST-020", "Continental Supply Chain", "DS", Decimal("0.160"), 40),
    Customer("CUST-021", "Inland Distribution Corp", "DS", Decimal("0.140"), 30),
    Customer("CUST-022", "Pacific Warehouse Group", "DS", Decimal("0.120"), 45),
    Customer("CUST-023", "Metro Fulfillment Services", "DS", Decimal("0.100"), 50),
    Customer("CUST-024", "CrossCountry Freight Inc", "DS", Decimal("0.100"), 35),
    Customer("CUST-025", "Harbor Shipping Co", "DS", Decimal("0.090"), 55),
    Customer("CUST-026", "Regional Express Logistics", "DS", Decimal("0.090"), 90),
)

# Quick validation at import time.
for _ec in ("PC", "AM", "DS"):
    _total = sum(c.revenue_share for c in CUSTOMERS if c.entity_code == _ec)
    assert _total == Decimal("1.000"), f"Customer shares for {_ec} sum to {_total}"


# ── AR aging data structures ─────────────────────────────────────────────────

@dataclass
class ARAgingEntry:
    """Year-end AR aging for one customer."""

    customer_id: str
    customer_name: str
    entity_code: str
    dso: int
    current: Decimal        # 0–30 days
    days_30: Decimal        # 31–60 days
    days_60: Decimal        # 61–90 days
    days_90: Decimal        # 91–120 days
    days_120_plus: Decimal  # 120+ days

    @property
    def total(self) -> Decimal:
        return (self.current + self.days_30 + self.days_60
                + self.days_90 + self.days_120_plus)


@dataclass
class MonthlyCollection:
    """Cash collected from AR in a given month for one entity."""

    year: int
    month: int
    entity_code: str
    amount: Decimal  # positive = cash received


@dataclass
class AllowanceAnalysis:
    """Annual bad debt allowance rollforward for one entity."""

    year: int
    entity_code: str
    beginning_balance: Decimal
    provision: Decimal  # increase to allowance for the year
    ending_balance: Decimal


# ── Reserve rates (fraction of each aging bucket) ────────────────────────────

RESERVE_RATES: dict[str, Decimal] = {
    "current": Decimal("0.01"),
    "days_30": Decimal("0.03"),
    "days_60": Decimal("0.10"),
    "days_90": Decimal("0.25"),
    "days_120_plus": Decimal("0.50"),
}


# ── Internal helpers ─────────────────────────────────────────────────────────

def _fraction_outstanding(age_days: int, dso: int) -> float:
    """Fraction of an invoice still outstanding based on age and customer DSO.

    Linear transition: 100 % when age < DSO − 15, 0 % when age > DSO + 15.
    """
    low = dso - 15
    high = dso + 15
    if age_days <= low:
        return 1.0
    if age_days >= high:
        return 0.0
    return (high - age_days) / 30.0


def _aging_bucket(age_days: int) -> str:
    """Map invoice age (days) to aging-bucket field name."""
    if age_days <= 30:
        return "current"
    if age_days <= 60:
        return "days_30"
    if age_days <= 90:
        return "days_60"
    if age_days <= 120:
        return "days_90"
    return "days_120_plus"


def _entity_monthly_revenue(
    revenue_records: list[MonthlyRevenue],
) -> dict[tuple[str, int, int], Decimal]:
    """Build lookup (entity_code, year, month) → total revenue (whole dollars).

    Uses the same rounding as ``post_revenue_to_gl`` (quantize to 1).
    """
    lookup: dict[tuple[str, int, int], Decimal] = {}
    for rec in revenue_records:
        key = (rec.entity_code, rec.year, rec.month)
        rev = rec.revenue.quantize(Decimal("1"), rounding=ROUND_HALF_UP)
        lookup[key] = lookup.get(key, Decimal(0)) + rev
    return lookup


def _customer_outstanding(
    cust: Customer,
    ref_year: int,
    ref_month: int,
    rev_lookup: dict[tuple[str, int, int], Decimal],
) -> tuple[Decimal, dict[str, Decimal]]:
    """Compute outstanding AR for a customer at month-end.

    Returns (total, {bucket_name: amount}).  Uses approximate ages
    (16 + lookback × 30 days) for consistency between aging and collections.
    """
    buckets: dict[str, Decimal] = {
        "current": Decimal(0),
        "days_30": Decimal(0),
        "days_60": Decimal(0),
        "days_90": Decimal(0),
        "days_120_plus": Decimal(0),
    }

    for lookback in range(6):
        m = ref_month - lookback
        y = ref_year
        if m <= 0:
            m += 12
            y -= 1

        entity_rev = rev_lookup.get((cust.entity_code, y, m), Decimal(0))
        if entity_rev == 0:
            continue

        cust_rev = (entity_rev * cust.revenue_share).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP
        )

        # Approximate age: current-month invoices are ~16 days old at month-end.
        age_days = 16 + lookback * 30

        frac = _fraction_outstanding(age_days, cust.dso)
        if frac <= 0:
            continue

        outstanding = (cust_rev * Decimal(str(frac))).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP
        )
        bucket = _aging_bucket(age_days)
        buckets[bucket] += outstanding

    total = sum(buckets.values())
    return total, buckets


def _entity_ar_at_month_end(
    entity_code: str,
    year: int,
    month: int,
    rev_lookup: dict[tuple[str, int, int], Decimal],
) -> Decimal:
    """Compute total AR balance for an entity at a given month-end."""
    total = Decimal(0)
    for cust in CUSTOMERS:
        if cust.entity_code == entity_code:
            cust_total, _ = _customer_outstanding(cust, year, month, rev_lookup)
            total += cust_total
    return total


# ── Public API ───────────────────────────────────────────────────────────────

def generate_ar_aging(
    revenue_records: list[MonthlyRevenue],
    year: int = 2025,
) -> list[ARAgingEntry]:
    """Compute year-end AR aging by customer for the given fiscal year.

    Returns a sorted list of ARAgingEntry (entity_code, customer_id).
    """
    rev_lookup = _entity_monthly_revenue(revenue_records)
    entries: list[ARAgingEntry] = []

    for cust in CUSTOMERS:
        _, buckets = _customer_outstanding(cust, year, 12, rev_lookup)
        entries.append(ARAgingEntry(
            customer_id=cust.id,
            customer_name=cust.name,
            entity_code=cust.entity_code,
            dso=cust.dso,
            **buckets,
        ))

    entries.sort(key=lambda e: (e.entity_code, e.customer_id))
    return entries


def generate_collections(
    revenue_records: list[MonthlyRevenue],
    years: list[int] | None = None,
) -> list[MonthlyCollection]:
    """Compute monthly cash collections from AR by entity.

    Collections = revenue posted to AR (1100) − ΔAR for each month.
    """
    if years is None:
        years = [2023, 2024, 2025]

    rev_lookup = _entity_monthly_revenue(revenue_records)
    collections: list[MonthlyCollection] = []

    for entity_code in sorted({"PC", "AM", "DS"}):
        prev_ar = Decimal(0)

        for year in sorted(years):
            for month in range(1, 13):
                rev_this_month = rev_lookup.get(
                    (entity_code, year, month), Decimal(0)
                )

                ar_end = _entity_ar_at_month_end(
                    entity_code, year, month, rev_lookup
                )

                collection = rev_this_month - (ar_end - prev_ar)

                if collection > 0:
                    collections.append(MonthlyCollection(
                        year=year,
                        month=month,
                        entity_code=entity_code,
                        amount=collection,
                    ))

                prev_ar = ar_end

    collections.sort(key=lambda c: (c.year, c.month, c.entity_code))
    return collections


def generate_allowance(
    revenue_records: list[MonthlyRevenue],
    years: list[int] | None = None,
) -> list[AllowanceAnalysis]:
    """Compute annual bad debt allowance rollforward by entity.

    Ending balance = Σ(aging_bucket × reserve_rate).  Provision = ΔBalance.
    """
    if years is None:
        years = [2023, 2024, 2025]

    results: list[AllowanceAnalysis] = []

    for entity_code in sorted({"PC", "AM", "DS"}):
        prev_balance = Decimal(0)

        for year in sorted(years):
            aging = generate_ar_aging(revenue_records, year=year)
            entity_aging = [e for e in aging if e.entity_code == entity_code]

            ending_balance = Decimal(0)
            for entry in entity_aging:
                for bucket, rate in RESERVE_RATES.items():
                    ending_balance += getattr(entry, bucket) * rate
            ending_balance = ending_balance.quantize(
                Decimal("1"), rounding=ROUND_HALF_UP
            )

            provision = ending_balance - prev_balance

            results.append(AllowanceAnalysis(
                year=year,
                entity_code=entity_code,
                beginning_balance=prev_balance,
                provision=provision,
                ending_balance=ending_balance,
            ))
            prev_balance = ending_balance

    return results


# ── GL posting ───────────────────────────────────────────────────────────────

def post_collections_to_gl(
    ledger: Ledger,
    collections: list[MonthlyCollection],
) -> None:
    """Post collection entries: DR 1015 (Cash — Collections Clearing), CR 1100 (AR Trade).

    Uses account 1015 rather than 1010 (Cash — Operating) because the
    model's operating cash flows are handled centrally through the parent
    bank model (TC-02).  Account 1015 represents cash collected from
    customers but not yet swept into the operating bank account.
    """
    for coll in collections:
        if coll.amount <= 0:
            continue

        date = datetime.date(coll.year, coll.month, 25)

        entry = JournalEntry(
            date=date,
            entity_code=coll.entity_code,
            description=f"AR collections {coll.year}-{coll.month:02d}",
            lines=(
                JournalEntryLine(
                    account="1015",
                    debit=coll.amount,
                    credit=Decimal(0),
                    memo="Cash collected — clearing",
                ),
                JournalEntryLine(
                    account="1100",
                    debit=Decimal(0),
                    credit=coll.amount,
                    memo="Cash collected from customers",
                ),
            ),
        )
        ledger.post(entry)


def post_allowance_to_gl(
    ledger: Ledger,
    allowance: list[AllowanceAnalysis],
) -> None:
    """Post allowance reclassification: DR 2020, CR 1150.

    The opex module posts bad debt expense as DR 6230, CR 2020.  This
    reclassification moves the credit to the proper contra-asset account
    (Allowance for Doubtful Accounts, 1150).
    """
    for rec in allowance:
        if rec.provision <= 0:
            continue

        date = datetime.date(rec.year, 12, 31)

        je = JournalEntry(
            date=date,
            entity_code=rec.entity_code,
            description=f"Bad debt allowance reclassification FY{rec.year}",
            lines=(
                JournalEntryLine(
                    account="2020",
                    debit=rec.provision,
                    credit=Decimal(0),
                    memo="Reclassify bad debt from accrued expenses",
                ),
                JournalEntryLine(
                    account="1150",
                    debit=Decimal(0),
                    credit=rec.provision,
                    memo="Allowance for doubtful accounts",
                ),
            ),
        )
        ledger.post(je)


# ── Invoice / receipt lifecycle records ──────────────────────────────────────
# These provide a canonical chain: Invoice → Receipt → Bank deposit
# so that every AR dollar traces from revenue recognition to cash collection.


@dataclass
class Invoice:
    """A customer invoice derived from monthly revenue."""

    invoice_id: str  # e.g. "INV-PC-2025-01-001"
    customer_id: str
    customer_name: str
    entity_code: str
    year: int
    month: int
    issue_date: datetime.date
    due_date: datetime.date
    amount: Decimal  # whole dollars
    product_lines: str  # descriptive, e.g. "Industrial Parts, Custom Machining"


@dataclass
class Receipt:
    """A cash receipt from a customer, referencing the originating invoice."""

    receipt_id: str  # e.g. "RCT-PC-2025-02-001"
    invoice_id: str
    customer_id: str
    customer_name: str
    entity_code: str
    receipt_date: datetime.date
    amount: Decimal  # whole dollars
    deposit_reference: str  # links to bank deposit category


def generate_invoices(
    revenue_records: list[MonthlyRevenue],
    years: list[int] | None = None,
) -> list[Invoice]:
    """Generate customer-level invoices from monthly revenue.

    Each month × entity produces one invoice per customer (proportional to
    revenue share), creating the first link in the sales-to-cash chain.
    """
    if years is None:
        years = [2023, 2024, 2025]

    rev_lookup = _entity_monthly_revenue(revenue_records)

    # Build entity → product-line-names mapping for invoice descriptions
    pl_by_entity: dict[str, list[str]] = {}
    for rec in revenue_records:
        pl_by_entity.setdefault(rec.entity_code, [])
        if rec.product_line not in pl_by_entity[rec.entity_code]:
            pl_by_entity[rec.entity_code].append(rec.product_line)
    for ec in sorted(pl_by_entity):
        pl_by_entity[ec].sort()

    invoices: list[Invoice] = []
    inv_seq = 0  # Global sequence counter for unique invoice IDs
    for entity_code in sorted({"PC", "AM", "DS"}):
        for year in sorted(years):
            for month in range(1, 13):
                entity_rev = rev_lookup.get((entity_code, year, month), Decimal(0))
                if entity_rev == 0:
                    continue

                entity_custs = [c for c in CUSTOMERS if c.entity_code == entity_code]
                for cust in entity_custs:
                    cust_rev = (entity_rev * cust.revenue_share).quantize(
                        Decimal("1"), rounding=ROUND_HALF_UP
                    )
                    if cust_rev <= 0:
                        continue

                    inv_seq += 1

                    issue_date = datetime.date(year, month, 15)
                    due_date = issue_date + datetime.timedelta(days=cust.dso)

                    invoices.append(Invoice(
                        invoice_id=f"INV-{entity_code}-{year}-{month:02d}-{inv_seq:04d}",
                        customer_id=cust.id,
                        customer_name=cust.name,
                        entity_code=entity_code,
                        year=year,
                        month=month,
                        issue_date=issue_date,
                        due_date=due_date,
                        amount=cust_rev,
                        product_lines=", ".join(pl_by_entity.get(entity_code, [])),
                    ))

    invoices.sort(key=lambda i: (i.year, i.month, i.entity_code, i.customer_id))
    return invoices


def generate_receipts(
    invoices: list[Invoice],
    collections: list[MonthlyCollection] | None = None,
) -> list[Receipt]:
    """Generate receipt records linking invoices to cash collections.

    Each invoice produces a single receipt when payment is expected based
    on the customer's DSO.  The receipt date is ``issue_date + DSO``
    (clamped to month-end on the 25th for GL posting consistency).  This
    creates a clean Invoice → Receipt → Bank-deposit chain.

    The ``collections`` parameter is accepted for API compatibility but
    not used; receipts are derived directly from invoice timing.
    """
    # Build customer DSO lookup
    cust_dso: dict[str, int] = {c.id: c.dso for c in CUSTOMERS}

    receipts: list[Receipt] = []
    receipt_seq: dict[str, int] = {}  # per-entity sequence counter

    for inv in invoices:
        dso = cust_dso.get(inv.customer_id, 30)
        expected_payment = inv.issue_date + datetime.timedelta(days=dso)

        # Receipt is posted on the 25th of the payment month
        pay_year = expected_payment.year
        pay_month = expected_payment.month
        receipt_date = datetime.date(pay_year, pay_month, 25)

        seq_key = inv.entity_code
        seq = receipt_seq.get(seq_key, 0) + 1
        receipt_seq[seq_key] = seq

        receipts.append(Receipt(
            receipt_id=f"RCT-{inv.entity_code}-{pay_year}-{pay_month:02d}-{seq:03d}",
            invoice_id=inv.invoice_id,
            customer_id=inv.customer_id,
            customer_name=inv.customer_name,
            entity_code=inv.entity_code,
            receipt_date=receipt_date,
            amount=inv.amount,
            deposit_reference=f"DEP-{inv.entity_code}-{pay_year}-{pay_month:02d}",
        ))

    receipts.sort(key=lambda r: (r.receipt_date, r.entity_code, r.receipt_id))
    return receipts


# ── Validation helpers ───────────────────────────────────────────────────────

def validate_ar_equals_gl(
    aging: list[ARAgingEntry],
    ledger: Ledger,
    entity_code: str,
    as_of_date: datetime.date | None = None,
) -> tuple[Decimal, Decimal, bool]:
    """Compare AR aging total to GL 1100 balance for an entity.

    Returns (aging_total, gl_balance, match).
    """
    aging_total = sum(e.total for e in aging if e.entity_code == entity_code)
    balances = ledger.balance_by_account(entity_code, as_of_date=as_of_date)
    gl_balance = balances.get("1100", Decimal(0))
    return aging_total, gl_balance, aging_total == gl_balance
