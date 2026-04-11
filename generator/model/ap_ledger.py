"""Cascade Industries 52K-row AP transaction ledger for TC-13 forensic analysis.

Generates ``ap_transactions_fy2025.csv`` with 52,000 rows.  Seven anomaly
categories are planted at deterministic transaction IDs so the gold standard
can reference them by ID.

Anomaly categories (per prompt.md TC-13):
1. Duplicate payments (4 exact + 4 near-dup, total exposure $127,340)
2. Benford's-law violation — 35 txns clustered $9,900-$9,999
3. Round-number payments — 12 txns to "Pacific Consulting Group"
4. Temporal anomalies — 15 weekend/holiday + 8 invoice-after-payment
5. Vendor anomalies — 2 similar-name vendors + 1 employee-address vendor
6. Split transactions — 3 sets split below approval threshold
7. Approver anomalies — 1 self-approved + 1 single-approver cost center

Feeds: TC-13 gold standard.
"""

from __future__ import annotations

import datetime
import random
from dataclasses import dataclass
from decimal import Decimal

from generator.model.ap import VENDORS, compute_vendor_annual_purchases
from generator.model.employees import Employee
from generator.model.revenue import MonthlyRevenue

# ── Data structures ─────────────────────────────────────────────────────────

@dataclass(frozen=True)
class APTransaction:
    """A single AP transaction row."""

    transaction_id: str
    date: datetime.date
    vendor_id: str
    vendor_name: str
    amount: Decimal
    description: str
    approver: str
    cost_center: str
    payment_method: str
    invoice_number: str

    # Metadata for gold standard (not written to CSV)
    anomaly_type: str | None = None
    anomaly_detail: str | None = None


@dataclass
class VendorSummary:
    """Per-vendor reconciliation between AP ledger and AP aging/subledger."""

    vendor_id: str
    vendor_name: str
    entity_code: str
    ledger_total: Decimal       # Sum of AP ledger transactions for this vendor
    target_purchases: Decimal   # Annual purchase target from revenue model
    anomaly_total: Decimal      # Sum of anomaly transactions for this vendor
    normal_total: Decimal       # Sum of normal transactions for this vendor


@dataclass
class APLedgerResult:
    """Return value of :func:`generate_ap_ledger`."""

    transactions: list[APTransaction]
    anomaly_index: dict[str, list[str]]  # anomaly_type → [transaction_id, ...]
    vendor_summaries: list[VendorSummary]  # Per-vendor reconciliation


# ── US federal holidays for FY2025 ─────────────────────────────────────────

_US_HOLIDAYS_2025 = frozenset([
    datetime.date(2025, 1, 1),   # New Year's Day
    datetime.date(2025, 1, 20),  # MLK Day
    datetime.date(2025, 2, 17),  # Presidents' Day
    datetime.date(2025, 5, 26),  # Memorial Day
    datetime.date(2025, 7, 4),   # Independence Day
    datetime.date(2025, 9, 1),   # Labor Day
    datetime.date(2025, 10, 13), # Columbus Day
    datetime.date(2025, 11, 11), # Veterans Day
    datetime.date(2025, 11, 27), # Thanksgiving
    datetime.date(2025, 12, 25), # Christmas
])


def _is_non_business_day(d: datetime.date) -> bool:
    return d.weekday() >= 5 or d in _US_HOLIDAYS_2025


# ── Helpers ─────────────────────────────────────────────────────────────────

_PAYMENT_METHODS = ("ACH", "Check", "Wire")
_PAYMENT_METHOD_WEIGHTS = (0.60, 0.30, 0.10)

# Anomaly vendor IDs (not in the main VENDORS tuple — forensic additions).
_PACIFIC_CONSULTING_ID = "VEND-050"
_PACIFIC_CONSULTING_NAME = "Pacific Consulting Group"
_JKL_LLC_ID = "VEND-051"
_JKL_LLC_NAME = "JKL Services LLC"
_JKL_INC_ID = "VEND-052"
_JKL_INC_NAME = "JKL Services Inc."
_EMPLOYEE_VENDOR_ID = "VEND-053"
_EMPLOYEE_VENDOR_NAME = "Willow Creek Consulting"

# Target employee whose home address matches a vendor.
_TARGET_EMPLOYEE_ID = "PC-0342"

# Approval threshold used for split-transaction anomalies.
_APPROVAL_THRESHOLD = Decimal("10000")

# Cost centers from the employee model (PC entity).
_COST_CENTERS = ["2100", "2200", "2300", "2400", "2600", "2700"]

# Single-approver cost center (anomaly 7b).
_SINGLE_APPROVER_CC = "2600"


def _random_date_fy2025(rng: random.Random) -> datetime.date:
    """Return a random business day in FY2025 (2025-01-01 to 2025-12-31)."""
    start = datetime.date(2025, 1, 1)
    end = datetime.date(2025, 12, 31)
    days_range = (end - start).days
    while True:
        d = start + datetime.timedelta(days=rng.randint(0, days_range))
        if not _is_non_business_day(d):
            return d


def _random_weekend_or_holiday(rng: random.Random) -> datetime.date:
    """Return a random weekend or holiday date in FY2025."""
    start = datetime.date(2025, 1, 1)
    end = datetime.date(2025, 12, 31)
    days_range = (end - start).days
    while True:
        d = start + datetime.timedelta(days=rng.randint(0, days_range))
        if _is_non_business_day(d):
            return d


def _make_invoice_number(rng: random.Random, vendor_id: str) -> str:
    """Generate a plausible invoice number."""
    prefix = vendor_id.replace("VEND-", "INV-")
    seq = rng.randint(10000, 99999)
    return f"{prefix}-{seq}"


def _pick_payment_method(rng: random.Random) -> str:
    return rng.choices(_PAYMENT_METHODS, weights=_PAYMENT_METHOD_WEIGHTS, k=1)[0]


def _pick_approver(rng: random.Random, employees: list[Employee]) -> str:
    """Pick a random active employee name as approver."""
    active = [e for e in employees if e.termination_date is None]
    return rng.choice(active).name


def _pick_cost_center(rng: random.Random) -> str:
    return rng.choice(_COST_CENTERS)


# ── Anomaly builders ────────────────────────────────────────────────────────

def _build_duplicate_anomalies(
    rng: random.Random,
    employees: list[Employee],
    next_id: int,
) -> tuple[list[APTransaction], int]:
    """Anomaly 1: 4 exact duplicates + 4 near-duplicates.

    Total exposure = $127,340.
    Each exact dup is a pair (original + duplicate).
    Each near-dup is a pair (original + near-dup with 1-digit invoice diff).
    """
    txns: list[APTransaction] = []

    # We need 8 pairs total (4 exact + 4 near-dup).
    # Target: total exposure $127,340 across all 8 duplicate instances.
    # Exposure = sum of the duplicated amounts (the second copy in each pair).
    exact_amounts = [
        Decimal("18500.00"),
        Decimal("12340.00"),
        Decimal("22750.00"),
        Decimal("9080.00"),
    ]  # sum = $62,670

    near_dup_amounts = [
        Decimal("24500.00"),
        Decimal("15670.00"),
        Decimal("14200.00"),
        Decimal("10300.00"),
    ]  # sum = $64,670  → total exposure = $127,340

    vendor_pool = [v for v in VENDORS[:8]]  # use first 8 vendors

    # Exact duplicates
    for i, amount in enumerate(exact_amounts):
        vendor = vendor_pool[i]
        date = _random_date_fy2025(rng)
        approver = _pick_approver(rng, employees)
        cc = _pick_cost_center(rng)
        pm = _pick_payment_method(rng)
        inv = _make_invoice_number(rng, vendor.id)

        base = APTransaction(
            transaction_id=f"APT-{next_id:06d}",
            date=date,
            vendor_id=vendor.id,
            vendor_name=vendor.name,
            amount=amount,
            description=f"Payment for services — {vendor.name}",
            approver=approver,
            cost_center=cc,
            payment_method=pm,
            invoice_number=inv,
            anomaly_type="exact_duplicate",
            anomaly_detail=f"Exact duplicate pair {i+1} of 4",
        )
        next_id += 1

        dup = APTransaction(
            transaction_id=f"APT-{next_id:06d}",
            date=date + datetime.timedelta(days=rng.randint(1, 5)),
            vendor_id=vendor.id,
            vendor_name=vendor.name,
            amount=amount,
            description=f"Payment for services — {vendor.name}",
            approver=approver,
            cost_center=cc,
            payment_method=pm,
            invoice_number=inv,  # same invoice = exact duplicate
            anomaly_type="exact_duplicate",
            anomaly_detail=f"Exact duplicate pair {i+1} of 4",
        )
        next_id += 1
        txns.extend([base, dup])

    # Near-duplicates (invoice differs by 1 digit)
    for i, amount in enumerate(near_dup_amounts):
        vendor = vendor_pool[4 + i]
        date = _random_date_fy2025(rng)
        approver = _pick_approver(rng, employees)
        cc = _pick_cost_center(rng)
        pm = _pick_payment_method(rng)
        inv = _make_invoice_number(rng, vendor.id)

        base = APTransaction(
            transaction_id=f"APT-{next_id:06d}",
            date=date,
            vendor_id=vendor.id,
            vendor_name=vendor.name,
            amount=amount,
            description=f"Vendor invoice — {vendor.name}",
            approver=approver,
            cost_center=cc,
            payment_method=pm,
            invoice_number=inv,
            anomaly_type="near_duplicate",
            anomaly_detail=f"Near-duplicate pair {i+1} of 4 — original invoice",
        )
        next_id += 1

        # Mutate one digit in the invoice number
        inv_chars = list(inv)
        # Change the last digit
        old_digit = inv_chars[-1]
        new_digit = str((int(old_digit) + 1) % 10)
        inv_chars[-1] = new_digit
        near_inv = "".join(inv_chars)

        dup = APTransaction(
            transaction_id=f"APT-{next_id:06d}",
            date=date + datetime.timedelta(days=rng.randint(1, 7)),
            vendor_id=vendor.id,
            vendor_name=vendor.name,
            amount=amount,
            description=f"Vendor invoice — {vendor.name}",
            approver=approver,
            cost_center=cc,
            payment_method=pm,
            invoice_number=near_inv,
            anomaly_type="near_duplicate",
            anomaly_detail=f"Near-duplicate pair {i+1} of 4 — typo invoice",
        )
        next_id += 1
        txns.extend([base, dup])

    return txns, next_id


def _build_benford_anomalies(
    rng: random.Random,
    employees: list[Employee],
    next_id: int,
) -> tuple[list[APTransaction], int]:
    """Anomaly 2: 35 transactions between $9,900-$9,999 (just below $10K)."""
    txns: list[APTransaction] = []
    for i in range(35):
        amount = Decimal(str(rng.randint(9900_00, 9999_99))) / Decimal("100")
        vendor = rng.choice(VENDORS)
        txns.append(APTransaction(
            transaction_id=f"APT-{next_id:06d}",
            date=_random_date_fy2025(rng),
            vendor_id=vendor.id,
            vendor_name=vendor.name,
            amount=amount,
            description=f"Consulting services — {vendor.name}",
            approver=_pick_approver(rng, employees),
            cost_center=_pick_cost_center(rng),
            payment_method=_pick_payment_method(rng),
            invoice_number=_make_invoice_number(rng, vendor.id),
            anomaly_type="benford_violation",
            anomaly_detail=f"Just-below-threshold txn {i+1} of 35 ($9,900-$9,999)",
        ))
        next_id += 1
    return txns, next_id


def _build_round_number_anomalies(
    rng: random.Random,
    employees: list[Employee],
    next_id: int,
) -> tuple[list[APTransaction], int]:
    """Anomaly 3: 12 round-number payments to Pacific Consulting Group."""
    txns: list[APTransaction] = []
    round_amounts = [
        Decimal("5000.00"), Decimal("5000.00"), Decimal("5000.00"),
        Decimal("10000.00"), Decimal("10000.00"), Decimal("10000.00"),
        Decimal("10000.00"),
        Decimal("25000.00"), Decimal("25000.00"), Decimal("25000.00"),
        Decimal("25000.00"), Decimal("25000.00"),
    ]
    for i, amount in enumerate(round_amounts):
        txns.append(APTransaction(
            transaction_id=f"APT-{next_id:06d}",
            date=_random_date_fy2025(rng),
            vendor_id=_PACIFIC_CONSULTING_ID,
            vendor_name=_PACIFIC_CONSULTING_NAME,
            amount=amount,
            description=f"Advisory services — {_PACIFIC_CONSULTING_NAME}",
            approver=_pick_approver(rng, employees),
            cost_center=_pick_cost_center(rng),
            payment_method=_pick_payment_method(rng),
            invoice_number=_make_invoice_number(rng, _PACIFIC_CONSULTING_ID),
            anomaly_type="round_number",
            anomaly_detail=f"Round payment {i+1} of 12 to shell-company-style vendor",
        ))
        next_id += 1
    return txns, next_id


def _build_temporal_anomalies(
    rng: random.Random,
    employees: list[Employee],
    next_id: int,
) -> tuple[list[APTransaction], int]:
    """Anomaly 4: 15 weekend/holiday approvals + 8 invoice-after-payment."""
    txns: list[APTransaction] = []

    # 15 weekend/holiday transactions
    for i in range(15):
        vendor = rng.choice(VENDORS)
        amount = Decimal(str(rng.randint(500_00, 50000_00))) / Decimal("100")
        txns.append(APTransaction(
            transaction_id=f"APT-{next_id:06d}",
            date=_random_weekend_or_holiday(rng),
            vendor_id=vendor.id,
            vendor_name=vendor.name,
            amount=amount,
            description=f"Supplies — {vendor.name}",
            approver=_pick_approver(rng, employees),
            cost_center=_pick_cost_center(rng),
            payment_method=_pick_payment_method(rng),
            invoice_number=_make_invoice_number(rng, vendor.id),
            anomaly_type="weekend_holiday",
            anomaly_detail=f"Weekend/holiday approval {i+1} of 15",
        ))
        next_id += 1

    # 8 invoice-after-payment (invoice_date > payment_date encoded in description)
    for i in range(8):
        vendor = rng.choice(VENDORS)
        amount = Decimal(str(rng.randint(1000_00, 30000_00))) / Decimal("100")
        payment_date = _random_date_fy2025(rng)
        # Invoice date is 3-15 days AFTER the payment date
        invoice_date = payment_date + datetime.timedelta(days=rng.randint(3, 15))
        if invoice_date > datetime.date(2025, 12, 31):
            invoice_date = datetime.date(2025, 12, 31)
        inv_num = _make_invoice_number(rng, vendor.id)
        txns.append(APTransaction(
            transaction_id=f"APT-{next_id:06d}",
            date=payment_date,
            vendor_id=vendor.id,
            vendor_name=vendor.name,
            amount=amount,
            description=(
                f"Invoice {inv_num} dated {invoice_date.isoformat()} — "
                f"{vendor.name}"
            ),
            approver=_pick_approver(rng, employees),
            cost_center=_pick_cost_center(rng),
            payment_method=_pick_payment_method(rng),
            invoice_number=inv_num,
            anomaly_type="invoice_after_payment",
            anomaly_detail=(
                f"Invoice date {invoice_date} after payment date {payment_date} "
                f"(pair {i+1} of 8)"
            ),
        ))
        next_id += 1

    return txns, next_id


def _build_vendor_anomalies(
    rng: random.Random,
    employees: list[Employee],
    next_id: int,
) -> tuple[list[APTransaction], int]:
    """Anomaly 5: 2 similar-name P.O. Box vendors + 1 employee-address vendor."""
    txns: list[APTransaction] = []

    # JKL Services LLC — 8 transactions
    for i in range(8):
        amount = Decimal(str(rng.randint(2000_00, 15000_00))) / Decimal("100")
        txns.append(APTransaction(
            transaction_id=f"APT-{next_id:06d}",
            date=_random_date_fy2025(rng),
            vendor_id=_JKL_LLC_ID,
            vendor_name=_JKL_LLC_NAME,
            amount=amount,
            description=f"Professional services — {_JKL_LLC_NAME}",
            approver=_pick_approver(rng, employees),
            cost_center=_pick_cost_center(rng),
            payment_method=_pick_payment_method(rng),
            invoice_number=_make_invoice_number(rng, _JKL_LLC_ID),
            anomaly_type="similar_name_vendor",
            anomaly_detail=f"JKL Services LLC txn {i+1} (P.O. Box vendor)",
        ))
        next_id += 1

    # JKL Services Inc. — 6 transactions
    for i in range(6):
        amount = Decimal(str(rng.randint(2000_00, 15000_00))) / Decimal("100")
        txns.append(APTransaction(
            transaction_id=f"APT-{next_id:06d}",
            date=_random_date_fy2025(rng),
            vendor_id=_JKL_INC_ID,
            vendor_name=_JKL_INC_NAME,
            amount=amount,
            description=f"Professional services — {_JKL_INC_NAME}",
            approver=_pick_approver(rng, employees),
            cost_center=_pick_cost_center(rng),
            payment_method=_pick_payment_method(rng),
            invoice_number=_make_invoice_number(rng, _JKL_INC_ID),
            anomaly_type="similar_name_vendor",
            anomaly_detail=f"JKL Services Inc. txn {i+1} (P.O. Box vendor)",
        ))
        next_id += 1

    # Employee-address vendor — 5 transactions, vendor address = PC-0342's home
    for i in range(5):
        amount = Decimal(str(rng.randint(1500_00, 8000_00))) / Decimal("100")
        txns.append(APTransaction(
            transaction_id=f"APT-{next_id:06d}",
            date=_random_date_fy2025(rng),
            vendor_id=_EMPLOYEE_VENDOR_ID,
            vendor_name=_EMPLOYEE_VENDOR_NAME,
            amount=amount,
            description=f"Consulting — {_EMPLOYEE_VENDOR_NAME}",
            approver=_pick_approver(rng, employees),
            cost_center=_pick_cost_center(rng),
            payment_method=_pick_payment_method(rng),
            invoice_number=_make_invoice_number(rng, _EMPLOYEE_VENDOR_ID),
            anomaly_type="employee_address_vendor",
            anomaly_detail=(
                f"Vendor at employee {_TARGET_EMPLOYEE_ID} home address "
                f"(txn {i+1} of 5)"
            ),
        ))
        next_id += 1

    return txns, next_id


def _build_split_anomalies(
    rng: random.Random,
    employees: list[Employee],
    next_id: int,
) -> tuple[list[APTransaction], int]:
    """Anomaly 6: 3 sets of split transactions below $10K threshold."""
    txns: list[APTransaction] = []

    splits = [
        (Decimal("7200.00"), Decimal("7300.00")),   # $14,500 split
        (Decimal("4800.00"), Decimal("4950.00")),   # $9,750 split
        (Decimal("6100.00"), Decimal("5900.00")),   # $12,000 split
    ]

    for i, (amt_a, amt_b) in enumerate(splits):
        vendor = VENDORS[10 + i]  # different vendor for each
        date_a = _random_date_fy2025(rng)
        date_b = date_a + datetime.timedelta(days=1)  # consecutive days
        if date_b.weekday() >= 5:
            date_b += datetime.timedelta(days=7 - date_b.weekday())
        approver = _pick_approver(rng, employees)
        cc = _pick_cost_center(rng)

        txns.append(APTransaction(
            transaction_id=f"APT-{next_id:06d}",
            date=date_a,
            vendor_id=vendor.id,
            vendor_name=vendor.name,
            amount=amt_a,
            description=f"Equipment purchase part A — {vendor.name}",
            approver=approver,
            cost_center=cc,
            payment_method=_pick_payment_method(rng),
            invoice_number=_make_invoice_number(rng, vendor.id),
            anomaly_type="split_transaction",
            anomaly_detail=f"Split set {i+1} of 3 — part A",
        ))
        next_id += 1

        txns.append(APTransaction(
            transaction_id=f"APT-{next_id:06d}",
            date=date_b,
            vendor_id=vendor.id,
            vendor_name=vendor.name,
            amount=amt_b,
            description=f"Equipment purchase part B — {vendor.name}",
            approver=approver,
            cost_center=cc,
            payment_method=_pick_payment_method(rng),
            invoice_number=_make_invoice_number(rng, vendor.id),
            anomaly_type="split_transaction",
            anomaly_detail=f"Split set {i+1} of 3 — part B",
        ))
        next_id += 1

    return txns, next_id


def _build_approver_anomalies(
    rng: random.Random,
    employees: list[Employee],
    next_id: int,
) -> tuple[list[APTransaction], int]:
    """Anomaly 7: 1 self-approved + 1 cost center with single approver (95%)."""
    txns: list[APTransaction] = []

    # 7a: Self-approved reimbursement
    active = [e for e in employees if e.termination_date is None]
    self_approver = next(
        e for e in active if e.entity_code == "PC" and e.department == "Finance"
    )
    amount = Decimal(str(rng.randint(500_00, 3000_00))) / Decimal("100")
    txns.append(APTransaction(
        transaction_id=f"APT-{next_id:06d}",
        date=_random_date_fy2025(rng),
        vendor_id="REIMB",
        vendor_name=self_approver.name,
        amount=amount,
        description=f"Employee reimbursement — {self_approver.name}",
        approver=self_approver.name,  # same person = self-approved
        cost_center=self_approver.cost_center,
        payment_method="ACH",
        invoice_number=f"REIMB-{rng.randint(10000, 99999)}",
        anomaly_type="self_approved",
        anomaly_detail=(
            f"Self-approved reimbursement by {self_approver.name} "
            f"({self_approver.employee_id})"
        ),
    ))
    next_id += 1

    # 7b: Single approver for cost center 2600 (Warehouse)
    # Generate 40 transactions for this CC, 38 with same approver (95%)
    single_approver = next(
        e for e in active
        if e.entity_code == "PC" and e.cost_center == _SINGLE_APPROVER_CC
    )
    for i in range(40):
        vendor = rng.choice(VENDORS[:10])  # PC vendors
        amount = Decimal(str(rng.randint(200_00, 8000_00))) / Decimal("100")
        if i < 38:  # 38/40 = 95%
            approver_name = single_approver.name
        else:
            approver_name = _pick_approver(rng, employees)

        txns.append(APTransaction(
            transaction_id=f"APT-{next_id:06d}",
            date=_random_date_fy2025(rng),
            vendor_id=vendor.id,
            vendor_name=vendor.name,
            amount=amount,
            description=f"Warehouse supplies — {vendor.name}",
            approver=approver_name,
            cost_center=_SINGLE_APPROVER_CC,
            payment_method=_pick_payment_method(rng),
            invoice_number=_make_invoice_number(rng, vendor.id),
            anomaly_type="single_approver_cc",
            anomaly_detail=(
                f"CC {_SINGLE_APPROVER_CC} single-approver pattern — "
                f"txn {i+1} of 40 "
                f"({'dominant approver' if i < 38 else 'other approver'})"
            ),
        ))
        next_id += 1

    return txns, next_id


# ── Normal transaction generator ────────────────────────────────────────────

def _build_normal_transactions(
    rng: random.Random,
    employees: list[Employee],
    count: int,
    next_id: int,
    vendor_targets: dict[str, Decimal] | None = None,
    anomaly_by_vendor: dict[str, Decimal] | None = None,
) -> tuple[list[APTransaction], int]:
    """Generate ``count`` normal (non-anomalous) AP transactions.

    When *vendor_targets* is supplied, transactions are distributed across
    vendors proportionally to their purchase targets, and amounts are scaled
    so each vendor's normal transactions sum to the target minus any anomaly
    amounts already attributed to that vendor.  This ties the 52K-row ledger
    to the AP aging / subledger / GL.
    """
    txns: list[APTransaction] = []

    if vendor_targets is None:
        # Legacy path: uniform random vendor, unscaled amounts.
        for _ in range(count):
            vendor = rng.choice(VENDORS)
            raw = rng.lognormvariate(7.5, 1.2)
            amount = Decimal(str(round(raw, 2)))
            if amount < Decimal("50.00"):
                amount = Decimal("50.00") + Decimal(str(rng.randint(0, 200))) / Decimal("100")
            if amount > Decimal("100000.00"):
                amount = Decimal(str(rng.randint(5000_00, 100000_00))) / Decimal("100")

            descriptions = [
                f"Materials purchase — {vendor.name}",
                f"Service payment — {vendor.name}",
                f"Monthly retainer — {vendor.name}",
                f"Parts order — {vendor.name}",
                f"Maintenance services — {vendor.name}",
                f"Equipment rental — {vendor.name}",
                f"Freight charges — {vendor.name}",
                f"Utilities — {vendor.name}",
            ]

            txns.append(APTransaction(
                transaction_id=f"APT-{next_id:06d}",
                date=_random_date_fy2025(rng),
                vendor_id=vendor.id,
                vendor_name=vendor.name,
                amount=amount,
                description=rng.choice(descriptions),
                approver=_pick_approver(rng, employees),
                cost_center=_pick_cost_center(rng),
                payment_method=_pick_payment_method(rng),
                invoice_number=_make_invoice_number(rng, vendor.id),
            ))
            next_id += 1

        return txns, next_id

    # ── Vendor-targeted path: tie to AP subledger ──────────────────────
    anom = anomaly_by_vendor or {}

    # Compute per-vendor normal targets (total target minus anomaly amounts).
    vendor_normal_targets: dict[str, Decimal] = {}
    total_target = Decimal(0)
    for vid, target in sorted(vendor_targets.items()):
        normal_target = target - anom.get(vid, Decimal(0))
        if normal_target < Decimal(0):
            normal_target = Decimal(0)
        vendor_normal_targets[vid] = normal_target
        total_target += normal_target

    # Allocate transaction count per vendor proportional to target.
    vendor_list = sorted(vendor_normal_targets.keys())
    vendor_counts: dict[str, int] = {}
    remaining = count
    for i, vid in enumerate(vendor_list):
        if i == len(vendor_list) - 1:
            vendor_counts[vid] = remaining
        else:
            share = vendor_normal_targets[vid] / total_target if total_target > 0 else Decimal(0)
            n = int(count * float(share))
            n = max(n, 20)  # at least 20 txns per vendor
            vendor_counts[vid] = n
            remaining -= n

    if remaining < 0:
        # Redistribute excess from the last vendor
        last = vendor_list[-1]
        vendor_counts[last] = max(vendor_counts[last], 20)

    # Build vendor lookup for names.
    vendor_by_id = {v.id: v for v in VENDORS}

    descriptions_templates = [
        "Materials purchase",
        "Service payment",
        "Monthly retainer",
        "Parts order",
        "Maintenance services",
        "Equipment rental",
        "Freight charges",
        "Utilities",
    ]

    for vid in vendor_list:
        vendor = vendor_by_id[vid]
        n = vendor_counts[vid]
        target = vendor_normal_targets[vid]

        if n == 0 or target <= 0:
            continue

        # Generate raw log-normal amounts, then scale to match target.
        raw_amounts: list[float] = []
        for _ in range(n):
            raw = rng.lognormvariate(7.5, 1.2)
            raw = max(raw, 50.0)
            raw = min(raw, 100000.0)
            raw_amounts.append(raw)

        raw_total = sum(raw_amounts)
        scale = float(target) / raw_total if raw_total > 0 else 1.0

        # Scale and quantize to cents.
        scaled: list[Decimal] = []
        running_target = target
        for j, raw in enumerate(raw_amounts):
            if j == n - 1:
                # Last transaction absorbs rounding remainder.
                amt = running_target
            else:
                amt = Decimal(str(round(raw * scale, 2)))
                # Clamp to at least $10.00 to avoid trivial rows.
                if amt < Decimal("10.00"):
                    amt = Decimal("10.00")
                running_target -= amt
            scaled.append(amt)

        # If rounding pushed the last amount negative, redistribute.
        if scaled[-1] < Decimal("10.00"):
            deficit = Decimal("10.00") - scaled[-1]
            scaled[-1] = Decimal("10.00")
            # Trim from the largest transaction.
            max_idx = max(range(len(scaled) - 1), key=lambda k: scaled[k])
            scaled[max_idx] -= deficit

        for amt in scaled:
            desc_template = rng.choice(descriptions_templates)
            txns.append(APTransaction(
                transaction_id=f"APT-{next_id:06d}",
                date=_random_date_fy2025(rng),
                vendor_id=vendor.id,
                vendor_name=vendor.name,
                amount=amt,
                description=f"{desc_template} — {vendor.name}",
                approver=_pick_approver(rng, employees),
                cost_center=_pick_cost_center(rng),
                payment_method=_pick_payment_method(rng),
                invoice_number=_make_invoice_number(rng, vendor.id),
            ))
            next_id += 1

    return txns, next_id


# ── Public API ──────────────────────────────────────────────────────────────

def generate_ap_ledger(
    rng: random.Random,
    employees: list[Employee],
    total_rows: int = 52_000,
    revenue_records: list[MonthlyRevenue] | None = None,
) -> APLedgerResult:
    """Generate the 52K-row AP transaction ledger with planted anomalies.

    Args:
        rng: Seeded random.Random for determinism.
        employees: Full employee roster (needed for approver names and
            employee-address matching).
        total_rows: Target total transaction count (default 52,000).
        revenue_records: Monthly revenue records.  When supplied, normal
            transaction amounts are scaled so each vendor's total ties to
            the AP subledger purchase volumes (procure-to-pay lifecycle).

    Returns:
        APLedgerResult with transactions sorted by transaction_id,
        an anomaly index mapping anomaly type to transaction IDs, and
        per-vendor summaries tying the ledger to AP aging.
    """
    all_txns: list[APTransaction] = []
    next_id = 1

    # Build anomalies first so they get deterministic low IDs.
    dup_txns, next_id = _build_duplicate_anomalies(rng, employees, next_id)
    all_txns.extend(dup_txns)

    benford_txns, next_id = _build_benford_anomalies(rng, employees, next_id)
    all_txns.extend(benford_txns)

    round_txns, next_id = _build_round_number_anomalies(rng, employees, next_id)
    all_txns.extend(round_txns)

    temporal_txns, next_id = _build_temporal_anomalies(rng, employees, next_id)
    all_txns.extend(temporal_txns)

    vendor_txns, next_id = _build_vendor_anomalies(rng, employees, next_id)
    all_txns.extend(vendor_txns)

    split_txns, next_id = _build_split_anomalies(rng, employees, next_id)
    all_txns.extend(split_txns)

    approver_txns, next_id = _build_approver_anomalies(rng, employees, next_id)
    all_txns.extend(approver_txns)

    # Compute per-vendor anomaly totals (only for vendors in the VENDORS tuple).
    anomaly_by_vendor: dict[str, Decimal] = {}
    vendor_ids = {v.id for v in VENDORS}
    for txn in all_txns:
        if txn.vendor_id in vendor_ids:
            anomaly_by_vendor[txn.vendor_id] = (
                anomaly_by_vendor.get(txn.vendor_id, Decimal(0)) + txn.amount
            )

    # Compute vendor purchase targets when revenue data is available.
    vendor_targets: dict[str, Decimal] | None = None
    if revenue_records is not None:
        vendor_targets = compute_vendor_annual_purchases(revenue_records, year=2025)

    # Fill remaining rows with normal transactions.
    anomaly_count = len(all_txns)
    normal_count = total_rows - anomaly_count
    assert normal_count > 0, (
        f"Anomalies ({anomaly_count}) exceed total_rows ({total_rows})"
    )

    normal_txns, next_id = _build_normal_transactions(
        rng, employees, normal_count, next_id,
        vendor_targets=vendor_targets,
        anomaly_by_vendor=anomaly_by_vendor,
    )
    all_txns.extend(normal_txns)

    # Sort by transaction_id for deterministic output.
    all_txns.sort(key=lambda t: t.transaction_id)

    # Build anomaly index.
    anomaly_index: dict[str, list[str]] = {}
    for txn in all_txns:
        if txn.anomaly_type:
            anomaly_index.setdefault(txn.anomaly_type, []).append(txn.transaction_id)

    # Build per-vendor summaries for reconciliation.
    vendor_by_id = {v.id: v for v in VENDORS}
    vendor_ledger_totals: dict[str, Decimal] = {}
    vendor_normal_totals: dict[str, Decimal] = {}
    vendor_anomaly_totals: dict[str, Decimal] = {}
    for txn in all_txns:
        if txn.vendor_id in vendor_ids:
            vendor_ledger_totals[txn.vendor_id] = (
                vendor_ledger_totals.get(txn.vendor_id, Decimal(0)) + txn.amount
            )
            if txn.anomaly_type:
                vendor_anomaly_totals[txn.vendor_id] = (
                    vendor_anomaly_totals.get(txn.vendor_id, Decimal(0)) + txn.amount
                )
            else:
                vendor_normal_totals[txn.vendor_id] = (
                    vendor_normal_totals.get(txn.vendor_id, Decimal(0)) + txn.amount
                )

    vendor_summaries: list[VendorSummary] = []
    for vid in sorted(vendor_ids):
        v = vendor_by_id[vid]
        target = vendor_targets.get(vid, Decimal(0)) if vendor_targets else Decimal(0)
        vendor_summaries.append(VendorSummary(
            vendor_id=vid,
            vendor_name=v.name,
            entity_code=v.entity_code,
            ledger_total=vendor_ledger_totals.get(vid, Decimal(0)),
            target_purchases=target,
            anomaly_total=vendor_anomaly_totals.get(vid, Decimal(0)),
            normal_total=vendor_normal_totals.get(vid, Decimal(0)),
        ))

    return APLedgerResult(
        transactions=all_txns,
        anomaly_index=anomaly_index,
        vendor_summaries=vendor_summaries,
    )
