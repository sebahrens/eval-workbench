"""Cascade Industries accounts payable subledger + aging.

Generates vendor-level AP aging (current/30/60/90 buckets) derived from
estimated purchase volumes.  Posts vendor invoice and payment entries to GL
so that account 2010 (AP Trade) balances to the AP aging total at each
year-end.

Key constraints:
- AP aging totals must equal GL 2010 balance (acceptance criterion).
- Vendor purchases estimated as a fraction of entity revenue (proxy for
  materials + services flowing through trade payables).
- Payment terms vary by vendor (Net 15 to Net 60).

Feeds: TC-13 (forensic AP), TC-14 (cash flow forecast).
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal

from generator.model.gl import JournalEntry, JournalEntryLine, Ledger
from generator.model.revenue import MonthlyRevenue

# ── Vendor definitions ───────────────────────────────────────────────────────
# Hardcoded for determinism.  Purchase shares sum to 1.000 per entity.

@dataclass(frozen=True)
class Vendor:
    """A trade vendor in the AP subledger."""

    id: str
    name: str
    entity_code: str
    purchase_share: Decimal  # fraction of entity annual purchases
    payment_terms: int  # days (Net 30, Net 45, etc.)


# Annual purchases through AP as a fraction of entity revenue.
# PC: materials-heavy mfg (67 % of $95 M ≈ $63.7 M).
# AM: materials + R&D supplies (50 % of $65 M = $32.5 M).
# DS: high pass-through logistics (84 % of $40 M = $33.6 M).
_PURCHASE_RATES: dict[str, Decimal] = {
    "PC": Decimal("0.67"),
    "AM": Decimal("0.50"),
    "DS": Decimal("0.84"),
}

VENDORS: tuple[Vendor, ...] = (
    # ── Precision Components ──────────────────────────────────────────────
    Vendor("VEND-001", "Portland Steel Supply", "PC", Decimal("0.250"), 45),
    Vendor("VEND-002", "Cascade Metal Works", "PC", Decimal("0.150"), 30),
    Vendor("VEND-003", "Pacific Northwest Electric", "PC", Decimal("0.120"), 30),
    Vendor("VEND-004", "Precision Tooling Inc", "PC", Decimal("0.100"), 45),
    Vendor("VEND-005", "Industrial Fastener Co", "PC", Decimal("0.080"), 30),
    Vendor("VEND-006", "Northwest Hydraulics", "PC", Decimal("0.080"), 45),
    Vendor("VEND-007", "Columbia Valley Metals", "PC", Decimal("0.070"), 30),
    Vendor("VEND-008", "Summit Machine Parts", "PC", Decimal("0.060"), 30),
    Vendor("VEND-009", "Heritage Welding Supply", "PC", Decimal("0.050"), 15),
    Vendor("VEND-010", "Evergreen Industrial", "PC", Decimal("0.040"), 15),
    # ── Advanced Materials ────────────────────────────────────────────────
    Vendor("VEND-011", "ChemSource International", "AM", Decimal("0.220"), 45),
    Vendor("VEND-012", "Nano Materials Corp", "AM", Decimal("0.150"), 60),
    Vendor("VEND-013", "Advanced Polymer Supply", "AM", Decimal("0.180"), 30),
    Vendor("VEND-014", "Lab Equipment Direct", "AM", Decimal("0.120"), 30),
    Vendor("VEND-015", "Spectrum Analytics", "AM", Decimal("0.100"), 45),
    Vendor("VEND-016", "Coastal Chemical Co", "AM", Decimal("0.090"), 30),
    Vendor("VEND-017", "BioTech Instruments", "AM", Decimal("0.080"), 60),
    Vendor("VEND-018", "Sierra Research Supply", "AM", Decimal("0.060"), 30),
    # ── Distribution Services ─────────────────────────────────────────────
    Vendor("VEND-019", "Fleet Maintenance Corp", "DS", Decimal("0.180"), 30),
    Vendor("VEND-020", "Fuel Systems America", "DS", Decimal("0.150"), 15),
    Vendor("VEND-021", "Warehouse Equipment Inc", "DS", Decimal("0.160"), 45),
    Vendor("VEND-022", "Continental Trucking Parts", "DS", Decimal("0.140"), 30),
    Vendor("VEND-023", "Pallet & Packaging Co", "DS", Decimal("0.120"), 30),
    Vendor("VEND-024", "National Fork Lift", "DS", Decimal("0.100"), 45),
    Vendor("VEND-025", "Climate Control Systems", "DS", Decimal("0.080"), 30),
    Vendor("VEND-026", "Safety Equipment Direct", "DS", Decimal("0.070"), 15),
)

for _ec in ("PC", "AM", "DS"):
    _total = sum(v.purchase_share for v in VENDORS if v.entity_code == _ec)
    assert _total == Decimal("1.000"), f"Vendor shares for {_ec} sum to {_total}"


# ── AP aging data structures ─────────────────────────────────────────────────

@dataclass
class APAgingEntry:
    """Year-end AP aging for one vendor."""

    vendor_id: str
    vendor_name: str
    entity_code: str
    payment_terms: int
    current: Decimal        # 0–30 days
    days_30: Decimal        # 31–60 days
    days_60: Decimal        # 61–90 days
    days_90_plus: Decimal   # 90+ days

    @property
    def total(self) -> Decimal:
        return self.current + self.days_30 + self.days_60 + self.days_90_plus


@dataclass
class MonthlyAPFlow:
    """Monthly AP purchases and payments for one entity."""

    year: int
    month: int
    entity_code: str
    purchases: Decimal  # new vendor invoices received
    payments: Decimal   # cash paid to vendors


# ── Internal helpers ─────────────────────────────────────────────────────────

def _fraction_unpaid(age_days: int, payment_terms: int) -> float:
    """Fraction of a vendor invoice still unpaid based on age and terms.

    Same linear-transition model as AR: 100 % when age < terms − 15,
    0 % when age > terms + 15.
    """
    low = payment_terms - 15
    high = payment_terms + 15
    if age_days <= low:
        return 1.0
    if age_days >= high:
        return 0.0
    return (high - age_days) / 30.0


def _ap_bucket(age_days: int) -> str:
    """Map invoice age (days) to AP aging bucket field name."""
    if age_days <= 30:
        return "current"
    if age_days <= 60:
        return "days_30"
    if age_days <= 90:
        return "days_60"
    return "days_90_plus"


def _entity_monthly_revenue(
    revenue_records: list[MonthlyRevenue],
) -> dict[tuple[str, int, int], Decimal]:
    """Build lookup (entity_code, year, month) → total revenue (whole dollars)."""
    lookup: dict[tuple[str, int, int], Decimal] = {}
    for rec in revenue_records:
        key = (rec.entity_code, rec.year, rec.month)
        rev = rec.revenue.quantize(Decimal("1"), rounding=ROUND_HALF_UP)
        lookup[key] = lookup.get(key, Decimal(0)) + rev
    return lookup


def _monthly_purchases(
    entity_code: str,
    year: int,
    month: int,
    rev_lookup: dict[tuple[str, int, int], Decimal],
) -> Decimal:
    """Compute total vendor purchases for an entity in a given month."""
    entity_rev = rev_lookup.get((entity_code, year, month), Decimal(0))
    rate = _PURCHASE_RATES.get(entity_code, Decimal(0))
    return (entity_rev * rate).quantize(Decimal("1"), rounding=ROUND_HALF_UP)


def _vendor_outstanding(
    vendor: Vendor,
    ref_year: int,
    ref_month: int,
    rev_lookup: dict[tuple[str, int, int], Decimal],
) -> tuple[Decimal, dict[str, Decimal]]:
    """Compute outstanding AP for a vendor at month-end.

    Returns (total, {bucket_name: amount}).
    """
    buckets: dict[str, Decimal] = {
        "current": Decimal(0),
        "days_30": Decimal(0),
        "days_60": Decimal(0),
        "days_90_plus": Decimal(0),
    }

    for lookback in range(4):  # 4 months is enough for Net 60 terms
        m = ref_month - lookback
        y = ref_year
        if m <= 0:
            m += 12
            y -= 1

        entity_purchases = _monthly_purchases(
            vendor.entity_code, y, m, rev_lookup
        )
        if entity_purchases == 0:
            continue

        vendor_purchases = (entity_purchases * vendor.purchase_share).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP
        )

        age_days = 16 + lookback * 30

        frac = _fraction_unpaid(age_days, vendor.payment_terms)
        if frac <= 0:
            continue

        unpaid = (vendor_purchases * Decimal(str(frac))).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP
        )
        bucket = _ap_bucket(age_days)
        buckets[bucket] += unpaid

    total = sum(buckets.values())
    return total, buckets


def _entity_ap_at_month_end(
    entity_code: str,
    year: int,
    month: int,
    rev_lookup: dict[tuple[str, int, int], Decimal],
) -> Decimal:
    """Compute total AP balance for an entity at a given month-end."""
    total = Decimal(0)
    for vendor in VENDORS:
        if vendor.entity_code == entity_code:
            v_total, _ = _vendor_outstanding(vendor, year, month, rev_lookup)
            total += v_total
    return total


# ── Public API ───────────────────────────────────────────────────────────────

def generate_ap_aging(
    revenue_records: list[MonthlyRevenue],
    year: int = 2025,
) -> list[APAgingEntry]:
    """Compute year-end AP aging by vendor for the given fiscal year.

    Returns a sorted list of APAgingEntry (entity_code, vendor_id).
    """
    rev_lookup = _entity_monthly_revenue(revenue_records)
    entries: list[APAgingEntry] = []

    for vendor in VENDORS:
        _, buckets = _vendor_outstanding(vendor, year, 12, rev_lookup)
        entries.append(APAgingEntry(
            vendor_id=vendor.id,
            vendor_name=vendor.name,
            entity_code=vendor.entity_code,
            payment_terms=vendor.payment_terms,
            **buckets,
        ))

    entries.sort(key=lambda e: (e.entity_code, e.vendor_id))
    return entries


def generate_ap_flows(
    revenue_records: list[MonthlyRevenue],
    years: list[int] | None = None,
) -> list[MonthlyAPFlow]:
    """Compute monthly AP purchases and payments by entity.

    Payments = purchases − ΔAP for each month.
    """
    if years is None:
        years = [2023, 2024, 2025]

    rev_lookup = _entity_monthly_revenue(revenue_records)
    flows: list[MonthlyAPFlow] = []

    for entity_code in sorted({"PC", "AM", "DS"}):
        prev_ap = Decimal(0)

        for year in sorted(years):
            for month in range(1, 13):
                purchases = _monthly_purchases(
                    entity_code, year, month, rev_lookup
                )

                ap_end = _entity_ap_at_month_end(
                    entity_code, year, month, rev_lookup
                )

                delta_ap = ap_end - prev_ap
                payments = purchases - delta_ap

                if payments < 0:
                    payments = Decimal(0)

                flows.append(MonthlyAPFlow(
                    year=year,
                    month=month,
                    entity_code=entity_code,
                    purchases=purchases,
                    payments=payments,
                ))

                prev_ap = ap_end

    flows.sort(key=lambda f: (f.year, f.month, f.entity_code))
    return flows


# ── GL posting ───────────────────────────────────────────────────────────────

def post_ap_to_gl(
    ledger: Ledger,
    flows: list[MonthlyAPFlow],
) -> None:
    """Post AP purchase and payment entries.

    Purchases: DR 1220 (Finished Goods inventory), CR 2010 (AP Trade).
    Payments:  DR 2010 (AP Trade), CR 1010 (Cash).
    """
    for flow in flows:
        date = datetime.date(flow.year, flow.month, 20)

        if flow.purchases > 0:
            purchase_entry = JournalEntry(
                date=date,
                entity_code=flow.entity_code,
                description=(
                    f"Vendor purchases {flow.year}-{flow.month:02d}"
                ),
                lines=(
                    JournalEntryLine(
                        account="1220",
                        debit=flow.purchases,
                        credit=Decimal(0),
                        memo="Inventory from vendor invoices",
                    ),
                    JournalEntryLine(
                        account="2010",
                        debit=Decimal(0),
                        credit=flow.purchases,
                        memo="Trade payables — vendor invoices",
                    ),
                ),
            )
            ledger.post(purchase_entry)

        if flow.payments > 0:
            payment_entry = JournalEntry(
                date=date,
                entity_code=flow.entity_code,
                description=(
                    f"Vendor payments {flow.year}-{flow.month:02d}"
                ),
                lines=(
                    JournalEntryLine(
                        account="2010",
                        debit=flow.payments,
                        credit=Decimal(0),
                        memo="Trade payables — vendor payments",
                    ),
                    JournalEntryLine(
                        account="1010",
                        debit=Decimal(0),
                        credit=flow.payments,
                        memo="Cash paid to vendors",
                    ),
                ),
            )
            ledger.post(payment_entry)


# ── Validation helpers ───────────────────────────────────────────────────────

def validate_ap_equals_gl(
    aging: list[APAgingEntry],
    ledger: Ledger,
    entity_code: str,
    as_of_date: datetime.date | None = None,
) -> tuple[Decimal, Decimal, bool]:
    """Compare AP aging total to GL 2010 balance for an entity.

    GL 2010 has a credit-normal balance, so the GL balance is negated
    (balance_by_account returns debit − credit, which is negative for
    credit-normal accounts).

    Returns (aging_total, gl_balance, match).
    """
    aging_total = sum(e.total for e in aging if e.entity_code == entity_code)
    balances = ledger.balance_by_account(entity_code, as_of_date=as_of_date)
    # 2010 is credit-normal: balance_by_account returns debit - credit → negative.
    gl_raw = balances.get("2010", Decimal(0))
    gl_balance = -gl_raw  # flip to positive for comparison
    return aging_total, gl_balance, aging_total == gl_balance
