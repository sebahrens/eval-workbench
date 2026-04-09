"""Cascade Industries bank transaction model (TC-02 of prompt.md).

Generates 340 bank transactions for December 2025, matching GL cash detail,
and computes the gold standard reconciliation.

Key gold-standard values from the spec:
- Adjusted bank balance = adjusted book balance = $4,287,331
- Bank confirmation letter balance = $4,312,117

Reconciliation structure:
  Bank ending balance:          $4,312,117 (= confirmation)
    + Deposits in transit (2):  + DIT
    - Outstanding checks (4):   - OC
    = Adjusted bank balance:    $4,287,331

  GL ending balance:            G
    + Interest earned (not in GL): + I
    - Bank charges (not in GL):    - S
    = Adjusted book balance:    $4,287,331

The model generates matched transactions between bank and GL with:
- Different descriptions (cryptic bank vs. company descriptions)
- Date differences (±2 business days for some transactions)
- 4 outstanding checks totaling OC (in GL but not on bank statement)
- 2 deposits in transit totaling DIT (in GL but not on bank statement)
- Bank charges + interest (on bank statement only, not in GL)
"""

from __future__ import annotations

import datetime
import random
from dataclasses import dataclass, field
from decimal import Decimal

from generator.model.gl import JournalEntry, JournalEntryLine, Ledger

# ── Gold-standard targets ──────────────────────────────────────────────────

ADJUSTED_BALANCE = Decimal("4287331")
CONFIRMATION_BALANCE = Decimal("4312117")

# ── Reconciling items ─────────────────────────────────────────────────────
# These are the specific items that bridge the bank balance to adjusted balance.

# 4 outstanding checks (in GL, not on bank statement)
OUTSTANDING_CHECKS: tuple[tuple[str, Decimal, datetime.date], ...] = (
    ("CHK-12847 — Vendor: Portland Steel Supply", Decimal("15230"), datetime.date(2025, 12, 29)),
    ("CHK-12851 — Vendor: Pacific Northwest Electric", Decimal("8475"), datetime.date(2025, 12, 30)),
    ("CHK-12853 — Vendor: Cascade Office Products", Decimal("3920"), datetime.date(2025, 12, 30)),
    ("CHK-12855 — Vendor: Metro Freight Lines", Decimal("12825"), datetime.date(2025, 12, 31)),
)

TOTAL_OUTSTANDING_CHECKS = sum(oc[1] for oc in OUTSTANDING_CHECKS)
assert TOTAL_OUTSTANDING_CHECKS == Decimal("40450")

# 2 deposits in transit (in GL, not on bank statement)
DEPOSITS_IN_TRANSIT: tuple[tuple[str, Decimal, datetime.date], ...] = (
    ("DEP — Customer: Northstar Manufacturing", Decimal("11427"), datetime.date(2025, 12, 30)),
    ("DEP — Customer: Western Allied Components", Decimal("4237"), datetime.date(2025, 12, 31)),
)

TOTAL_DEPOSITS_IN_TRANSIT = sum(d[1] for d in DEPOSITS_IN_TRANSIT)
assert TOTAL_DEPOSITS_IN_TRANSIT == Decimal("15664")

# Verify: bank ending + DIT - OC = adjusted
assert CONFIRMATION_BALANCE + TOTAL_DEPOSITS_IN_TRANSIT - TOTAL_OUTSTANDING_CHECKS == ADJUSTED_BALANCE

# Bank-only items (interest and charges, not in GL)
BANK_INTEREST = Decimal("4567")  # Interest earned, bank recorded but company hasn't
BANK_SERVICE_CHARGES = Decimal("657")  # Monthly service charge

# GL ending balance: adjusted - interest + charges
GL_ENDING_BALANCE = ADJUSTED_BALANCE - BANK_INTEREST + BANK_SERVICE_CHARGES
assert GL_ENDING_BALANCE == Decimal("4283421")

# Bank statement starting balance (Dec 1, 2025)
BANK_STARTING_BALANCE = Decimal("3856442")


# ── Transaction categories for realistic generation ──────────────────────

@dataclass(frozen=True)
class TransactionTemplate:
    """Template for generating bank/GL transaction pairs."""

    category: str
    bank_desc_prefix: str  # Cryptic bank description prefix
    gl_desc_prefix: str  # Company GL description
    min_amount: int
    max_amount: int
    is_debit: bool  # True = money out (checks, payments), False = money in


_INFLOW_TEMPLATES: tuple[TransactionTemplate, ...] = (
    TransactionTemplate("customer_payment", "ACH CR", "Customer Payment —", 5000, 85000, False),
    TransactionTemplate("wire_in", "WR CR", "Wire Transfer Received —", 20000, 150000, False),
    TransactionTemplate("deposit", "DEP", "Bank Deposit", 8000, 45000, False),
    TransactionTemplate("ach_collection", "ACH CR COLL", "ACH Collection —", 3000, 25000, False),
    TransactionTemplate("refund", "CR MEMO", "Vendor Refund —", 500, 5000, False),
)

_OUTFLOW_TEMPLATES: tuple[TransactionTemplate, ...] = (
    TransactionTemplate("vendor_payment", "ACH DB", "Vendor Payment —", 2000, 40000, True),
    TransactionTemplate("check", "CHK#", "Check —", 1000, 20000, True),
    TransactionTemplate("wire_out", "WR DB", "Wire Transfer —", 10000, 80000, True),
    TransactionTemplate("payroll", "ACH DB PAYROLL", "Payroll — Bi-weekly", 120000, 180000, True),
    TransactionTemplate("tax_payment", "ACH DB TAX", "Tax Payment —", 5000, 50000, True),
    TransactionTemplate("utility", "ACH DB UTIL", "Utility Payment —", 800, 4000, True),
    TransactionTemplate("insurance", "ACH DB INS", "Insurance Premium —", 3000, 15000, True),
    TransactionTemplate("lease", "ACH DB LEASE", "Lease Payment —", 5000, 25000, True),
)

# Vendor/customer names for realistic descriptions
_VENDOR_NAMES: tuple[str, ...] = (
    "Portland Steel Supply", "Pacific Northwest Electric", "Cascade Office Products",
    "Metro Freight Lines", "Northwest Industrial Gas", "Columbia River Packaging",
    "Willamette Valley Paper", "Oregon Chemical Supply", "Puget Sound Logistics",
    "Mountain View Components", "Evergreen Fabrication", "Silver Falls Hardware",
    "Crater Lake Materials", "Hood River Distribution", "Deschutes Manufacturing",
    "Rogue Valley Services", "Umpqua Tech Supply", "Klamath Equipment",
    "Baker City Metals", "Bend Industrial Supply",
)

_CUSTOMER_NAMES: tuple[str, ...] = (
    "Northstar Manufacturing", "Western Allied Components", "Pacific Rim Industries",
    "Summit Engineering Corp", "Redwood Precision", "Granite State Manufacturing",
    "Blue Ridge Automotive", "Prairie Wind Energy", "Golden Gate Defense",
    "Lakeshore Electronics", "Pinnacle Aerospace", "Ironwood Fabrication",
    "Sterling Heavy Industries", "Coastal Marine Systems", "Ridgeline Power",
    "Timberline Construction", "Copper Mountain Mining", "Clearwater Technologies",
    "Sunstone Semiconductor", "Emerald City Robotics",
)


@dataclass
class BankTransaction:
    """A single transaction on the bank statement."""

    date: datetime.date
    description: str  # Bank's cryptic format
    amount: Decimal  # Positive = credit (inflow), negative = debit (outflow)
    running_balance: Decimal
    # Internal tracking (not output to bank statement)
    category: str = ""
    matched: bool = True  # Whether this appears on both bank and GL


@dataclass
class GLCashEntry:
    """A single line in the GL cash detail for account 1010."""

    date: datetime.date
    reference: str  # e.g., "JE-2025-1234" or "CHK-12847"
    description: str  # Company's description format
    debit: Decimal  # Money in
    credit: Decimal  # Money out
    running_balance: Decimal
    # Internal tracking
    category: str = ""
    matched: bool = True  # Whether this appears on both bank and GL


@dataclass
class BankModel:
    """The complete bank data model for TC-02.

    Contains all generated transactions plus the gold standard reconciliation.
    """

    bank_transactions: list[BankTransaction] = field(default_factory=list)
    gl_entries: list[GLCashEntry] = field(default_factory=list)
    bank_starting_balance: Decimal = BANK_STARTING_BALANCE
    bank_ending_balance: Decimal = CONFIRMATION_BALANCE
    gl_starting_balance: Decimal = Decimal("0")
    gl_ending_balance: Decimal = GL_ENDING_BALANCE
    confirmation_balance: Decimal = CONFIRMATION_BALANCE
    adjusted_balance: Decimal = ADJUSTED_BALANCE
    total_outstanding_checks: Decimal = TOTAL_OUTSTANDING_CHECKS
    total_deposits_in_transit: Decimal = TOTAL_DEPOSITS_IN_TRANSIT
    bank_interest: Decimal = BANK_INTEREST
    bank_charges: Decimal = BANK_SERVICE_CHARGES


def _business_days_offset(
    date: datetime.date,
    offset: int,
) -> datetime.date:
    """Shift a date by ``offset`` business days (skipping weekends)."""
    direction = 1 if offset >= 0 else -1
    remaining = abs(offset)
    current = date
    while remaining > 0:
        current += datetime.timedelta(days=direction)
        if current.weekday() < 5:  # Mon-Fri
            remaining -= 1
    return current


def generate_bank_model(rng: random.Random) -> BankModel:
    """Generate 340 bank transactions for December 2025 and matching GL entries.

    The generator produces:
    - ~334 matched transactions (on both bank statement and GL)
    - 4 outstanding checks (GL only)
    - 2 deposits in transit (GL only)
    - Bank interest + service charges (bank only)

    340 = matched bank transactions + bank-only items (interest, charges)
    GL entries = matched transactions + outstanding checks + deposits in transit

    Returns a fully populated BankModel with gold standard reconciliation data.
    """
    model = BankModel()

    # We need 340 bank statement rows.
    # 2 bank-only items (interest, charges) → 338 matched transactions.
    n_matched = 338
    # ── Step 1: Generate matched transaction amounts ─────────────────────
    # We need the net of matched transactions to produce the correct ending balance.
    # net_matched = bank_ending - bank_starting - interest + charges
    net_matched = (
        CONFIRMATION_BALANCE
        - BANK_STARTING_BALANCE
        - BANK_INTEREST
        + BANK_SERVICE_CHARGES
    )
    # net_matched = 4312117 - 3856442 - 4567 + 657 = 451765

    # Generate individual transaction amounts that sum to net_matched.
    # Mix of inflows and outflows for a realistic operating account.
    matched_amounts: list[Decimal] = []
    matched_categories: list[str] = []
    matched_is_debit: list[bool] = []

    # Structured transactions first: 2 payrolls (bi-weekly, 2 in December)
    payroll_amounts = [Decimal("156832"), Decimal("161447")]
    for pa in payroll_amounts:
        matched_amounts.append(-pa)  # outflow
        matched_categories.append("payroll")
        matched_is_debit.append(True)

    # Generate remaining transactions
    remaining_count = n_matched - len(matched_amounts)
    remaining_net = net_matched - sum(matched_amounts)

    # Target: ~60% inflows by count, ~40% outflows by count
    n_inflows = int(remaining_count * 0.55)
    n_outflows = remaining_count - n_inflows

    # Generate inflow amounts
    inflow_total = Decimal("0")
    inflows: list[tuple[Decimal, str]] = []
    for i in range(n_inflows):
        tmpl = _INFLOW_TEMPLATES[rng.randint(0, len(_INFLOW_TEMPLATES) - 1)]
        amount = Decimal(str(rng.randint(tmpl.min_amount, tmpl.max_amount)))
        inflows.append((amount, tmpl.category))
        inflow_total += amount

    # Generate outflow amounts
    outflow_total = Decimal("0")
    outflows: list[tuple[Decimal, str]] = []
    for i in range(n_outflows):
        tmpl = _OUTFLOW_TEMPLATES[rng.randint(0, len(_OUTFLOW_TEMPLATES) - 1)]
        if tmpl.category == "payroll":
            tmpl = _OUTFLOW_TEMPLATES[0]  # Skip extra payroll, use vendor payment
        amount = Decimal(str(rng.randint(tmpl.min_amount, tmpl.max_amount)))
        outflows.append((amount, tmpl.category))
        outflow_total += amount

    # Current net from generated amounts (excluding payroll already added)
    generated_net = inflow_total - outflow_total

    # Adjust the last inflow to make everything balance
    adjustment = remaining_net - generated_net
    if inflows:
        old_amount, cat = inflows[-1]
        inflows[-1] = (old_amount + adjustment, cat)

    # Add inflows to matched lists
    for amount, cat in inflows:
        matched_amounts.append(amount)  # positive = inflow
        matched_categories.append(cat)
        matched_is_debit.append(False)

    # Add outflows to matched lists
    for amount, cat in outflows:
        matched_amounts.append(-amount)  # negative = outflow
        matched_categories.append(cat)
        matched_is_debit.append(True)

    # Verify net
    assert sum(matched_amounts) == net_matched, (
        f"Net mismatch: {sum(matched_amounts)} != {net_matched}"
    )

    # ── Step 2: Assign dates across December 2025 ───────────────────────
    # Spread transactions across business days in December.
    dec_start = datetime.date(2025, 12, 1)
    dec_end = datetime.date(2025, 12, 31)

    business_days: list[datetime.date] = []
    d = dec_start
    while d <= dec_end:
        if d.weekday() < 5:  # Mon-Fri
            business_days.append(d)
        d += datetime.timedelta(days=1)

    # Assign each matched transaction a bank date
    bank_dates: list[datetime.date] = []
    for i in range(n_matched):
        # Payrolls on specific dates (15th and 31st or nearest business day)
        if i == 0:
            bank_dates.append(datetime.date(2025, 12, 15))  # Payroll 1
        elif i == 1:
            bank_dates.append(datetime.date(2025, 12, 31))  # Payroll 2
        else:
            bank_dates.append(business_days[rng.randint(0, len(business_days) - 1)])

    # Sort by date for chronological bank statement
    combined = list(zip(bank_dates, matched_amounts, matched_categories, matched_is_debit))
    combined.sort(key=lambda x: (x[0], x[2], str(x[1])))
    bank_dates = [c[0] for c in combined]
    matched_amounts = [c[1] for c in combined]
    matched_categories = [c[2] for c in combined]
    matched_is_debit = [c[3] for c in combined]

    # ── Step 3: Build bank statement rows ────────────────────────────────
    running = BANK_STARTING_BALANCE

    # Determine where to insert bank-only items (interest and charges)
    # Interest: typically credited on last business day
    # Charges: typically on last business day
    interest_inserted = False

    vendor_idx = 0
    customer_idx = 0
    je_counter = 2025_12_001

    for i in range(n_matched):
        amount = matched_amounts[i]
        category = matched_categories[i]
        is_debit = matched_is_debit[i]
        bank_date = bank_dates[i]

        # Insert bank-only items before Dec 31 transactions
        if bank_date >= datetime.date(2025, 12, 31) and not interest_inserted:
            # Interest credit
            running += BANK_INTEREST
            model.bank_transactions.append(BankTransaction(
                date=datetime.date(2025, 12, 31),
                description="INT CR MONTHLY INTEREST",
                amount=BANK_INTEREST,
                running_balance=running,
                category="interest",
                matched=False,
            ))
            interest_inserted = True

            # Service charge
            running -= BANK_SERVICE_CHARGES
            model.bank_transactions.append(BankTransaction(
                date=datetime.date(2025, 12, 31),
                description="SVC CHG MONTHLY MAINTENANCE FEE",
                amount=-BANK_SERVICE_CHARGES,
                running_balance=running,
                category="service_charge",
                matched=False,
            ))
        # Build bank description (cryptic)
        bank_desc = _make_bank_description(
            category, is_debit, rng,
            vendor_names=_VENDOR_NAMES,
            customer_names=_CUSTOMER_NAMES,
            vendor_idx=vendor_idx,
            customer_idx=customer_idx,
        )
        if is_debit and category in ("vendor_payment", "check"):
            vendor_idx = (vendor_idx + 1) % len(_VENDOR_NAMES)
        elif not is_debit and category in ("customer_payment", "wire_in", "ach_collection"):
            customer_idx = (customer_idx + 1) % len(_CUSTOMER_NAMES)

        running += amount
        model.bank_transactions.append(BankTransaction(
            date=bank_date,
            description=bank_desc,
            amount=amount,
            running_balance=running,
            category=category,
            matched=True,
        ))

        # Build corresponding GL entry with possible date offset
        offset_days = rng.choice([0, 0, 0, 0, 0, -1, -1, 1, 1, -2, 2])
        gl_date = _business_days_offset(bank_date, offset_days)
        # Clamp to December
        if gl_date < dec_start:
            gl_date = dec_start
        if gl_date > dec_end:
            gl_date = dec_end

        gl_desc = _make_gl_description(
            category, is_debit, rng,
            vendor_names=_VENDOR_NAMES,
            customer_names=_CUSTOMER_NAMES,
            vendor_idx=(vendor_idx - 1) % len(_VENDOR_NAMES) if is_debit else vendor_idx,
            customer_idx=(customer_idx - 1) % len(_CUSTOMER_NAMES) if not is_debit else customer_idx,
        )

        ref = f"JE-{je_counter}"
        je_counter += 1

        if is_debit:
            model.gl_entries.append(GLCashEntry(
                date=gl_date,
                reference=ref,
                description=gl_desc,
                debit=Decimal("0"),
                credit=abs(amount),
                running_balance=Decimal("0"),  # Computed later
                category=category,
                matched=True,
            ))
        else:
            model.gl_entries.append(GLCashEntry(
                date=gl_date,
                reference=ref,
                description=gl_desc,
                debit=amount,
                credit=Decimal("0"),
                running_balance=Decimal("0"),  # Computed later
                category=category,
                matched=True,
            ))

    # Insert bank-only items if not yet inserted (all transactions were before Dec 31)
    if not interest_inserted:
        running += BANK_INTEREST
        model.bank_transactions.append(BankTransaction(
            date=datetime.date(2025, 12, 31),
            description="INT CR MONTHLY INTEREST",
            amount=BANK_INTEREST,
            running_balance=running,
            category="interest",
            matched=False,
        ))
        running -= BANK_SERVICE_CHARGES
        model.bank_transactions.append(BankTransaction(
            date=datetime.date(2025, 12, 31),
            description="SVC CHG MONTHLY MAINTENANCE FEE",
            amount=-BANK_SERVICE_CHARGES,
            running_balance=running,
            category="service_charge",
            matched=False,
        ))

    # Verify bank ending balance
    assert running == CONFIRMATION_BALANCE, (
        f"Bank ending mismatch: {running} != {CONFIRMATION_BALANCE}"
    )

    # ── Step 4: Add outstanding checks to GL ─────────────────────────────
    for desc, amount, date in OUTSTANDING_CHECKS:
        ref = f"JE-{je_counter}"
        je_counter += 1
        model.gl_entries.append(GLCashEntry(
            date=date,
            reference=ref,
            description=desc,
            debit=Decimal("0"),
            credit=amount,
            running_balance=Decimal("0"),  # Computed later
            category="outstanding_check",
            matched=False,
        ))

    # ── Step 5: Add deposits in transit to GL ────────────────────────────
    for desc, amount, date in DEPOSITS_IN_TRANSIT:
        ref = f"JE-{je_counter}"
        je_counter += 1
        model.gl_entries.append(GLCashEntry(
            date=date,
            reference=ref,
            description=desc,
            debit=amount,
            credit=Decimal("0"),
            running_balance=Decimal("0"),  # Computed later
            category="deposit_in_transit",
            matched=False,
        ))

    # ── Step 6: Sort GL entries by date and compute running balance ───────
    model.gl_entries.sort(key=lambda e: (e.date, e.reference))

    # GL starting balance: ending - net of all entries
    gl_net = sum(e.debit - e.credit for e in model.gl_entries)
    model.gl_starting_balance = GL_ENDING_BALANCE - gl_net

    gl_running = model.gl_starting_balance
    for entry in model.gl_entries:
        gl_running += entry.debit - entry.credit
        entry.running_balance = gl_running

    # Verify GL ending
    assert gl_running == GL_ENDING_BALANCE, (
        f"GL ending mismatch: {gl_running} != {GL_ENDING_BALANCE}"
    )

    # ── Step 7: Verify transaction counts ────────────────────────────────
    assert len(model.bank_transactions) == 340, (
        f"Bank transaction count: {len(model.bank_transactions)} != 340"
    )

    return model


def _make_bank_description(
    category: str,
    is_debit: bool,
    rng: random.Random,
    *,
    vendor_names: tuple[str, ...],
    customer_names: tuple[str, ...],
    vendor_idx: int,
    customer_idx: int,
) -> str:
    """Generate a cryptic bank-style description."""
    if category == "payroll":
        return f"ACH DB PAYROLL CASCADE IND {rng.randint(1000, 9999)}"
    if category == "customer_payment":
        name = customer_names[customer_idx % len(customer_names)]
        abbr = name[:12].upper().replace(" ", "")
        return f"ACH CR {abbr} {rng.randint(100000, 999999)}"
    if category == "wire_in":
        name = customer_names[customer_idx % len(customer_names)]
        abbr = name[:10].upper().replace(" ", "")
        return f"WR CR {abbr} REF{rng.randint(10000, 99999)}"
    if category == "deposit":
        return f"DEP {rng.randint(1000, 9999)}"
    if category == "ach_collection":
        name = customer_names[customer_idx % len(customer_names)]
        abbr = name[:10].upper().replace(" ", "")
        return f"ACH CR COLL {abbr}"
    if category == "refund":
        name = vendor_names[vendor_idx % len(vendor_names)]
        abbr = name[:12].upper().replace(" ", "")
        return f"CR MEMO {abbr}"
    if category == "vendor_payment":
        name = vendor_names[vendor_idx % len(vendor_names)]
        abbr = name[:12].upper().replace(" ", "")
        return f"ACH DB {abbr} {rng.randint(100000, 999999)}"
    if category == "check":
        return f"CHK#{rng.randint(12700, 12850)}"
    if category == "wire_out":
        name = vendor_names[vendor_idx % len(vendor_names)]
        abbr = name[:10].upper().replace(" ", "")
        return f"WR DB {abbr} REF{rng.randint(10000, 99999)}"
    if category == "tax_payment":
        agencies = ["IRS", "OR DOR", "TX COMP", "IL DOR"]
        return f"ACH DB TAX {rng.choice(agencies)} {rng.randint(1000, 9999)}"
    if category == "utility":
        providers = ["PGE", "NW NATGAS", "PORTLND WTR", "COMED"]
        return f"ACH DB UTIL {rng.choice(providers)}"
    if category == "insurance":
        return f"ACH DB INS HARTFORD {rng.randint(100, 999)}"
    if category == "lease":
        return f"ACH DB LEASE PMT {rng.randint(100, 999)}"
    return f"{'DB' if is_debit else 'CR'} MISC {rng.randint(1000, 9999)}"


def _make_gl_description(
    category: str,
    is_debit: bool,
    rng: random.Random,
    *,
    vendor_names: tuple[str, ...],
    customer_names: tuple[str, ...],
    vendor_idx: int,
    customer_idx: int,
) -> str:
    """Generate a company-style GL description."""
    if category == "payroll":
        return "Payroll — Bi-weekly"
    if category == "customer_payment":
        name = customer_names[customer_idx % len(customer_names)]
        return f"Customer Payment — {name}"
    if category == "wire_in":
        name = customer_names[customer_idx % len(customer_names)]
        return f"Wire Transfer Received — {name}"
    if category == "deposit":
        return "Bank Deposit"
    if category == "ach_collection":
        name = customer_names[customer_idx % len(customer_names)]
        return f"ACH Collection — {name}"
    if category == "refund":
        name = vendor_names[vendor_idx % len(vendor_names)]
        return f"Vendor Refund — {name}"
    if category == "vendor_payment":
        name = vendor_names[vendor_idx % len(vendor_names)]
        return f"Vendor Payment — {name}"
    if category == "check":
        name = vendor_names[vendor_idx % len(vendor_names)]
        return f"Check — {name}"
    if category == "wire_out":
        name = vendor_names[vendor_idx % len(vendor_names)]
        return f"Wire Transfer — {name}"
    if category == "tax_payment":
        return "Tax Payment — Federal/State"
    if category == "utility":
        return "Utility Payment"
    if category == "insurance":
        return "Insurance Premium — Hartford"
    if category == "lease":
        return "Lease Payment"
    return f"{'Cash Disbursement' if is_debit else 'Cash Receipt'}"


def post_bank_to_gl(ledger: Ledger, model: BankModel) -> None:
    """Post GL cash entries to the ledger as journal entries.

    Posts only entries for account 1010 (Cash — Operating). The offsetting
    entries go to suspense-like accounts based on category:
    - customer_payment/wire_in/deposit/ach_collection → DR 1010, CR 1100 (A/R)
    - vendor_payment/check/wire_out → DR 2010 (A/P), CR 1010
    - payroll → DR 2030 (Accrued Payroll), CR 1010
    - tax_payment → DR 2070 (Tax Payable), CR 1010
    - utility/insurance/lease → DR 6xxx (various expense), CR 1010
    - outstanding_check → DR 2010, CR 1010
    - deposit_in_transit → DR 1010, CR 1100
    """
    _OFFSET_ACCOUNTS: dict[str, str] = {
        "customer_payment": "1100",
        "wire_in": "1100",
        "deposit": "1100",
        "ach_collection": "1100",
        "refund": "2010",
        "vendor_payment": "2010",
        "check": "2010",
        "wire_out": "2010",
        "payroll": "2030",
        "tax_payment": "2070",
        "utility": "6110",
        "insurance": "6140",
        "lease": "6100",
        "outstanding_check": "2010",
        "deposit_in_transit": "1100",
    }

    for entry in model.gl_entries:
        offset_acct = _OFFSET_ACCOUNTS.get(entry.category, "2150")
        amount = entry.debit if entry.debit > 0 else entry.credit

        if entry.debit > 0:
            # Cash inflow: DR 1010, CR offset
            lines = (
                JournalEntryLine(account="1010", debit=amount, credit=Decimal("0"),
                                 memo=entry.description),
                JournalEntryLine(account=offset_acct, debit=Decimal("0"), credit=amount,
                                 memo=entry.description),
            )
        else:
            # Cash outflow: DR offset, CR 1010
            lines = (
                JournalEntryLine(account=offset_acct, debit=amount, credit=Decimal("0"),
                                 memo=entry.description),
                JournalEntryLine(account="1010", debit=Decimal("0"), credit=amount,
                                 memo=entry.description),
            )

        je = JournalEntry(
            date=entry.date,
            entity_code="CI",  # Parent entity — operating account
            description=entry.description,
            lines=lines,
        )
        ledger.post(je)


# ── Validation helpers ──────────────────────────────────────────────────

def validate_reconciliation(model: BankModel) -> list[str]:
    """Validate that the model's reconciliation matches gold standard values.

    Returns a list of errors (empty = valid).
    """
    errors: list[str] = []

    # Bank side
    bank_ending = model.bank_transactions[-1].running_balance
    if bank_ending != CONFIRMATION_BALANCE:
        errors.append(
            f"Bank ending balance {bank_ending} != confirmation {CONFIRMATION_BALANCE}"
        )

    adjusted_bank = bank_ending + model.total_deposits_in_transit - model.total_outstanding_checks
    if adjusted_bank != ADJUSTED_BALANCE:
        errors.append(
            f"Adjusted bank balance {adjusted_bank} != target {ADJUSTED_BALANCE}"
        )

    # Book side
    gl_ending = model.gl_entries[-1].running_balance
    if gl_ending != GL_ENDING_BALANCE:
        errors.append(
            f"GL ending balance {gl_ending} != target {GL_ENDING_BALANCE}"
        )

    adjusted_book = gl_ending + model.bank_interest - model.bank_charges
    if adjusted_book != ADJUSTED_BALANCE:
        errors.append(
            f"Adjusted book balance {adjusted_book} != target {ADJUSTED_BALANCE}"
        )

    # Count checks
    if len(model.bank_transactions) != 340:
        errors.append(
            f"Bank transaction count {len(model.bank_transactions)} != 340"
        )

    oc_count = sum(1 for e in model.gl_entries if e.category == "outstanding_check")
    if oc_count != 4:
        errors.append(f"Outstanding check count {oc_count} != 4")

    dit_count = sum(1 for e in model.gl_entries if e.category == "deposit_in_transit")
    if dit_count != 2:
        errors.append(f"Deposit in transit count {dit_count} != 2")

    return errors
