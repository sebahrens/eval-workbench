"""Cascade Industries lease portfolio and ASC 842 computations (TC-04, TC-06).

Generates 15 leases across all entities with:
- Lessee/lessor names, commencement dates, terms, base rent, escalation
- Renewal options, purchase options, termination provisions
- 2 leases qualifying for short-term exemption (≤12 months remaining)
- 3 leases with amendments that materially change terms
- ROU asset and lease liability computations per ASC 842
- GL postings: operating lease expense (straight-line), finance lease
  amortisation + interest, and balance sheet entries

The ASC 842 lease adjustments create a temporary difference for TC-06:
book lease expense (straight-line) differs from the cash rent paid,
producing a ROU asset / lease liability gap on the balance sheet.

Determinism: uses only the passed ``rng``; no unordered sets or wall-clock reads.
"""

from __future__ import annotations

import datetime
import random
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal
from enum import Enum

from generator.model.gl import JournalEntry, JournalEntryLine, Ledger

# ── Enums ───────────────────────────────────────────────────────────────────

class LeaseType(Enum):
    OPERATING = "operating"
    FINANCE = "finance"


class EscalationType(Enum):
    FIXED_PCT = "fixed_pct"
    CPI = "cpi"
    STEPPED = "stepped"
    NONE = "none"


# ── Dataclasses ─────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class Amendment:
    """A material amendment to a lease that changes key terms."""

    effective_date: datetime.date
    description: str
    # Amended fields (None = unchanged from original)
    new_monthly_rent: Decimal | None = None
    new_term_months: int | None = None
    new_escalation_pct: Decimal | None = None


@dataclass(frozen=True)
class Lease:
    """A single lease in the Cascade Industries portfolio."""

    lease_id: str
    entity_code: str
    lessee: str            # Always a Cascade entity
    lessor: str            # External landlord/lessor name
    description: str       # What's being leased
    commencement_date: datetime.date
    term_months: int       # Original contractual term
    monthly_base_rent: Decimal
    escalation_type: EscalationType
    escalation_pct: Decimal        # Annual escalation % (0 if none/CPI)
    escalation_steps: tuple[Decimal, ...] | None  # For stepped escalation
    renewal_option_months: int     # 0 = no renewal option
    renewal_rent_increase_pct: Decimal  # Rent increase if renewed
    purchase_option: bool
    purchase_option_price: Decimal | None  # None if no purchase option
    termination_provision: str     # Free-text description
    lease_type: LeaseType          # Operating or finance
    short_term_exempt: bool        # True if ≤12 months remaining
    amendments: tuple[Amendment, ...]  # Material amendments (may be empty)

    # Pre-computed ASC 842 values at commencement (or amendment effective date)
    discount_rate: Decimal         # Incremental borrowing rate
    rou_asset_initial: Decimal     # Initial ROU asset
    lease_liability_initial: Decimal  # Initial lease liability

    @property
    def effective_monthly_rent(self) -> Decimal:
        """Monthly rent after applying last amendment, if any."""
        for amend in reversed(self.amendments):
            if amend.new_monthly_rent is not None:
                return amend.new_monthly_rent
        return self.monthly_base_rent

    @property
    def effective_term_months(self) -> int:
        """Term after applying last amendment, if any."""
        for amend in reversed(self.amendments):
            if amend.new_term_months is not None:
                return amend.new_term_months
        return self.term_months

    @property
    def end_date(self) -> datetime.date:
        """Lease end date based on commencement + effective term."""
        months = self.effective_term_months
        year = self.commencement_date.year + (self.commencement_date.month - 1 + months) // 12
        month = (self.commencement_date.month - 1 + months) % 12 + 1
        # Last day of the ending month
        if month == 12:
            return datetime.date(year, 12, 31)
        return datetime.date(year, month, 1) - datetime.timedelta(days=1)


# ── Discount rate helper ────────────────────────────────────────────────────

# Cascade's incremental borrowing rate varies by term length.
# Short leases get a lower rate; longer leases get higher.
_IBR_BY_TERM: list[tuple[int, Decimal]] = [
    (24, Decimal("0.045")),   # ≤2yr
    (60, Decimal("0.050")),   # ≤5yr
    (120, Decimal("0.055")),  # ≤10yr
    (999, Decimal("0.060")),  # >10yr
]


def _ibr_for_term(months: int) -> Decimal:
    for threshold, rate in _IBR_BY_TERM:
        if months <= threshold:
            return rate
    return Decimal("0.060")


# ── PV of lease payments ───────────────────────────────────────────────────

def _pv_lease_payments(
    monthly_rent: Decimal,
    term_months: int,
    annual_rate: Decimal,
) -> Decimal:
    """Present value of an annuity of monthly payments.

    Uses monthly discounting: r = annual_rate / 12.
    PV = rent × [(1 - (1+r)^-n) / r]
    """
    if term_months <= 0:
        return Decimal(0)
    r = annual_rate / Decimal(12)
    if r == 0:
        return monthly_rent * term_months
    factor = (Decimal(1) - (Decimal(1) + r) ** (-term_months)) / r
    return (monthly_rent * factor).quantize(Decimal("1"), rounding=ROUND_HALF_UP)


# ── Lease templates ─────────────────────────────────────────────────────────
# (entity_code, lessor, description, term_months, monthly_rent, escalation,
#  esc_pct, renewal_months, purchase_option, lease_type, short_term,
#  amendment_specs)

# Amendment spec: (months_after_commence, new_rent_delta, new_term_change, desc)
_AmendSpec = tuple[int, int | None, int | None, str]

_LEASE_TEMPLATES: list[dict] = [
    # ── Precision Components (PC) — several long-term leases (§1.2) ────
    dict(
        entity_code="PC", lessor="Pacific Northwest Realty Partners",
        description="Manufacturing facility — Building A",
        term_months=120, monthly_rent=45_000,
        escalation=EscalationType.FIXED_PCT, esc_pct="0.030",
        renewal_months=60, purchase=False, lease_type=LeaseType.OPERATING,
        short_term=False, amendments=[],
    ),
    dict(
        entity_code="PC", lessor="Columbia Property Trust",
        description="Warehouse & shipping dock",
        term_months=84, monthly_rent=28_000,
        escalation=EscalationType.FIXED_PCT, esc_pct="0.025",
        renewal_months=36, purchase=False, lease_type=LeaseType.OPERATING,
        short_term=False,
        # Amendment: landlord expanded the space, rent increased
        amendments=[(36, 6_000, 24, "Expanded warehouse footprint by 4,000 sq ft; rent increased and term extended")],
    ),
    dict(
        entity_code="PC", lessor="Willamette Equipment Leasing",
        description="CNC milling center — Haas VF-4SS",
        term_months=60, monthly_rent=8_500,
        escalation=EscalationType.NONE, esc_pct="0",
        renewal_months=0, purchase=True, purchase_price=85_000,
        lease_type=LeaseType.FINANCE, short_term=False, amendments=[],
    ),
    dict(
        entity_code="PC", lessor="Evergreen Realty Group",
        description="Administrative office — Suite 200",
        term_months=36, monthly_rent=12_000,
        escalation=EscalationType.STEPPED, esc_pct="0",
        stepped_rents=(Decimal("12000"), Decimal("12600"), Decimal("13200")),
        renewal_months=24, purchase=False, lease_type=LeaseType.OPERATING,
        short_term=False, amendments=[],
    ),
    dict(
        entity_code="PC", lessor="NW Fleet Services",
        description="Delivery fleet — 3 box trucks",
        term_months=48, monthly_rent=4_200,
        escalation=EscalationType.NONE, esc_pct="0",
        renewal_months=0, purchase=False, lease_type=LeaseType.OPERATING,
        short_term=False, amendments=[],
    ),
    # ── Advanced Materials (AM) — growing, newer leases ────────────────
    dict(
        entity_code="AM", lessor="Lone Star Commercial Properties",
        description="R&D laboratory — Building 7",
        term_months=96, monthly_rent=38_000,
        escalation=EscalationType.CPI, esc_pct="0",
        renewal_months=48, purchase=False, lease_type=LeaseType.OPERATING,
        short_term=False, amendments=[],
    ),
    dict(
        entity_code="AM", lessor="Texas Industrial Realty",
        description="Production facility — West Wing",
        term_months=120, monthly_rent=52_000,
        escalation=EscalationType.FIXED_PCT, esc_pct="0.035",
        renewal_months=60, purchase=False, lease_type=LeaseType.OPERATING,
        short_term=False,
        # Amendment: converted open storage to clean room, major rent increase
        amendments=[(
            48, 15_000, 0,
            "Converted 2,500 sq ft to ISO Class 7 clean room; rent increased to reflect TI improvements",
        )],
    ),
    dict(
        entity_code="AM", lessor="Austin Materials Testing LLC",
        description="Electron microscope — FEI Titan Themis",
        term_months=60, monthly_rent=12_000,
        escalation=EscalationType.NONE, esc_pct="0",
        renewal_months=0, purchase=True, purchase_price=320_000,
        lease_type=LeaseType.FINANCE, short_term=False, amendments=[],
    ),
    dict(
        entity_code="AM", lessor="Hill Country Office Partners",
        description="Office suite — 3rd floor",
        term_months=10, monthly_rent=6_500,
        escalation=EscalationType.NONE, esc_pct="0",
        renewal_months=0, purchase=False, lease_type=LeaseType.OPERATING,
        # SHORT-TERM EXEMPT: 10-month lease, no renewal, no purchase option
        short_term=True, amendments=[],
    ),
    # ── Distribution Services (DS) — asset-light, but some leases ─────
    dict(
        entity_code="DS", lessor="Midwest Logistics Properties",
        description="Distribution center — Main warehouse",
        term_months=84, monthly_rent=65_000,
        escalation=EscalationType.FIXED_PCT, esc_pct="0.028",
        renewal_months=36, purchase=False, lease_type=LeaseType.OPERATING,
        short_term=False,
        # Amendment: reduced footprint after operational review
        amendments=[(
            60, -12_000, -12,
            "Returned 8,000 sq ft of underused racking area; rent reduced and term shortened",
        )],
    ),
    dict(
        entity_code="DS", lessor="Lakeside Equipment Rental",
        description="Forklift fleet — 8 units",
        term_months=48, monthly_rent=6_800,
        escalation=EscalationType.NONE, esc_pct="0",
        renewal_months=12, purchase=False, lease_type=LeaseType.OPERATING,
        short_term=False, amendments=[],
    ),
    dict(
        entity_code="DS", lessor="Great Lakes Realty Trust",
        description="Satellite staging facility — Elk Grove",
        term_months=60, monthly_rent=22_000,
        escalation=EscalationType.FIXED_PCT, esc_pct="0.020",
        renewal_months=24, purchase=False, lease_type=LeaseType.OPERATING,
        short_term=False, amendments=[],
    ),
    dict(
        entity_code="DS", lessor="Prairie State Fleet Management",
        description="Refrigerated truck — single unit",
        term_months=9, monthly_rent=3_800,
        escalation=EscalationType.NONE, esc_pct="0",
        renewal_months=0, purchase=False, lease_type=LeaseType.OPERATING,
        # SHORT-TERM EXEMPT: 9-month lease
        short_term=True, amendments=[],
    ),
    # ── Parent (CI) — corporate HQ ────────────────────────────────────
    dict(
        entity_code="CI", lessor="Cascade Tower Management LLC",
        description="Corporate headquarters — floors 8-10",
        term_months=120, monthly_rent=85_000,
        escalation=EscalationType.FIXED_PCT, esc_pct="0.030",
        renewal_months=60, purchase=False, lease_type=LeaseType.OPERATING,
        short_term=False, amendments=[],
    ),
    dict(
        entity_code="CI", lessor="Pacific Data Centers Inc.",
        description="Colocation — primary data center rack space",
        term_months=60, monthly_rent=15_000,
        escalation=EscalationType.FIXED_PCT, esc_pct="0.025",
        renewal_months=24, purchase=False, lease_type=LeaseType.OPERATING,
        short_term=False, amendments=[],
    ),
]


# ── Lease generation ────────────────────────────────────────────────────────

def _commencement_date(rng: random.Random, entity_code: str, term_months: int) -> datetime.date:
    """Deterministic commencement date.

    Long leases start earlier (2019–2022); short-term leases start in
    mid-2025 so they're still active at year-end 2025.
    """
    if term_months <= 12:
        # Short-term: commence between Mar–Jun 2025
        month = rng.randint(3, 6)
        return datetime.date(2025, month, 1)
    elif term_months <= 60:
        # Medium: 2021–2023
        year = rng.randint(2021, 2023)
        month = rng.randint(1, 12)
        return datetime.date(year, month, 1)
    else:
        # Long: 2019–2022
        year = rng.randint(2019, 2022)
        month = rng.randint(1, 12)
        return datetime.date(year, month, 1)


def generate_leases(rng: random.Random) -> list[Lease]:
    """Generate the 15-lease portfolio deterministically.

    Returns a list sorted by lease_id.
    """
    leases: list[Lease] = []

    for idx, tmpl in enumerate(_LEASE_TEMPLATES, start=1):
        lease_id = f"LS-{idx:03d}"
        entity_code = tmpl["entity_code"]
        term_months = tmpl["term_months"]

        commence = _commencement_date(rng, entity_code, term_months)
        monthly_rent = Decimal(str(tmpl["monthly_rent"]))
        esc_pct = Decimal(tmpl["esc_pct"])

        # Build amendments
        amendments: list[Amendment] = []
        for amend_spec in tmpl.get("amendments", []):
            months_after, rent_delta, term_delta, desc = amend_spec
            eff_date = datetime.date(
                commence.year + (commence.month - 1 + months_after) // 12,
                (commence.month - 1 + months_after) % 12 + 1,
                1,
            )
            new_rent = (monthly_rent + Decimal(str(rent_delta))) if rent_delta is not None else None
            new_term = (term_months + term_delta) if term_delta is not None else None
            amendments.append(Amendment(
                effective_date=eff_date,
                description=desc,
                new_monthly_rent=new_rent,
                new_term_months=new_term,
            ))

        # Determine effective values (after amendments)
        eff_rent = monthly_rent
        eff_term = term_months
        for amend in amendments:
            if amend.new_monthly_rent is not None:
                eff_rent = amend.new_monthly_rent
            if amend.new_term_months is not None:
                eff_term = amend.new_term_months

        # Discount rate
        discount_rate = _ibr_for_term(eff_term)

        # ASC 842: compute initial ROU asset and lease liability
        if tmpl.get("short_term", False):
            # Short-term exempt: no ROU or liability recognised
            rou_initial = Decimal(0)
            liability_initial = Decimal(0)
        else:
            liability_initial = _pv_lease_payments(eff_rent, eff_term, discount_rate)
            # ROU asset = liability + any prepaid rent (none for our leases)
            rou_initial = liability_initial

        # Stepped escalation rents
        stepped = tmpl.get("stepped_rents")

        # Purchase option
        has_purchase = tmpl.get("purchase", False)
        purchase_price = Decimal(str(tmpl["purchase_price"])) if has_purchase else None

        # Termination provision text
        if term_months <= 12:
            termination = "Lease expires at term end; no early termination."
        elif has_purchase:
            termination = "Lessee may terminate upon exercise of purchase option or at term end."
        else:
            notice_days = rng.choice([90, 120, 180])
            penalty_months = rng.choice([3, 6])
            termination = f"Early termination with {notice_days}-day notice and {penalty_months}-month penalty."

        # Renewal rent increase (small random variation)
        renewal_months = tmpl.get("renewal_months", 0)
        if renewal_months > 0:
            renewal_increase = Decimal(str(rng.choice(["0.03", "0.04", "0.05"])))
        else:
            renewal_increase = Decimal(0)

        lease = Lease(
            lease_id=lease_id,
            entity_code=entity_code,
            lessee=_entity_legal_name(entity_code),
            lessor=tmpl["lessor"],
            description=tmpl["description"],
            commencement_date=commence,
            term_months=term_months,
            monthly_base_rent=monthly_rent,
            escalation_type=tmpl["escalation"],
            escalation_pct=esc_pct,
            escalation_steps=tuple(stepped) if stepped else None,
            renewal_option_months=renewal_months,
            renewal_rent_increase_pct=renewal_increase,
            purchase_option=has_purchase,
            purchase_option_price=purchase_price,
            termination_provision=termination,
            lease_type=tmpl["lease_type"],
            short_term_exempt=tmpl.get("short_term", False),
            amendments=tuple(amendments),
            discount_rate=discount_rate,
            rou_asset_initial=rou_initial,
            lease_liability_initial=liability_initial,
        )
        leases.append(lease)

    leases.sort(key=lambda ls: ls.lease_id)
    return leases


_ENTITY_NAMES: dict[str, str] = {
    "CI": "Cascade Industries, Inc.",
    "PC": "Cascade Precision Components LLC",
    "AM": "Cascade Advanced Materials, Inc.",
    "DS": "Cascade Distribution Services LLC",
}


def _entity_legal_name(code: str) -> str:
    return _ENTITY_NAMES[code]


# ── ASC 842 schedule computations ──────────────────────────────────────────

@dataclass
class LeaseScheduleRow:
    """One year of ASC 842 schedule data for a single lease."""

    lease_id: str
    entity_code: str
    year: int
    lease_type: LeaseType
    # Balance sheet
    rou_asset_beg: Decimal
    rou_asset_end: Decimal
    lease_liability_beg: Decimal
    lease_liability_end: Decimal
    # Income statement
    lease_expense: Decimal       # Operating: straight-line; Finance: amortisation
    interest_expense: Decimal    # Finance leases only (operating = 0)
    # Cash
    cash_paid: Decimal           # Actual rent paid during the year


def compute_lease_schedules(
    leases: list[Lease],
    years: list[int] | None = None,
) -> list[LeaseScheduleRow]:
    """Compute annual ASC 842 schedules for all non-exempt leases.

    For operating leases:
    - Lease expense = total undiscounted payments / term (straight-line)
    - Liability reduction = cash paid - interest on liability
    - Interest on liability = beg liability × monthly rate × 12
    - ROU amortisation = lease expense - interest

    For finance leases:
    - Amortisation = ROU asset / term (straight-line)
    - Interest = beg liability × rate
    - Liability reduction = cash paid - interest
    """
    if years is None:
        years = [2023, 2024, 2025]

    rows: list[LeaseScheduleRow] = []

    for lease in leases:
        if lease.short_term_exempt:
            # Short-term leases: expense = cash paid, no ROU/liability
            for year in sorted(years):
                months_active = _months_active_in_year(lease, year)
                if months_active <= 0:
                    continue
                cash = lease.monthly_base_rent * months_active
                rows.append(LeaseScheduleRow(
                    lease_id=lease.lease_id,
                    entity_code=lease.entity_code,
                    year=year,
                    lease_type=lease.lease_type,
                    rou_asset_beg=Decimal(0),
                    rou_asset_end=Decimal(0),
                    lease_liability_beg=Decimal(0),
                    lease_liability_end=Decimal(0),
                    lease_expense=cash,
                    interest_expense=Decimal(0),
                    cash_paid=cash,
                ))
            continue

        # Compute the full monthly amortisation schedule, then aggregate by year.
        monthly_rate = lease.discount_rate / Decimal(12)
        eff_term = lease.effective_term_months
        eff_rent = lease.effective_monthly_rent
        total_payments = eff_rent * eff_term

        # Straight-line expense (operating) or amortisation (finance)
        if eff_term > 0:
            straight_line_monthly = (total_payments / Decimal(eff_term)).quantize(
                Decimal("0.01"), rounding=ROUND_HALF_UP
            )
        else:
            straight_line_monthly = Decimal(0)

        liability = lease.lease_liability_initial
        rou = lease.rou_asset_initial

        # Walk month by month from commencement through end
        for year in sorted(years):
            months_active = _months_active_in_year(lease, year)
            if months_active <= 0:
                # Lease not active this year — skip
                continue

            rou_beg = rou
            liab_beg = liability

            year_expense = Decimal(0)
            year_interest = Decimal(0)
            year_cash = Decimal(0)

            for _m in range(months_active):
                # Interest on opening liability for the month
                interest = (liability * monthly_rate).quantize(
                    Decimal("0.01"), rounding=ROUND_HALF_UP
                )

                if lease.lease_type == LeaseType.OPERATING:
                    # Operating: expense is straight-line
                    expense = straight_line_monthly
                    # Liability: accrue interest then reduce by (cash - interest)
                    liability = liability + interest - eff_rent
                    # ROU: reduce by (expense - interest) to keep balance sheet balanced
                    rou_reduction = expense - interest
                    rou = rou - rou_reduction
                    year_interest += Decimal(0)  # Op leases: interest embedded in expense
                else:
                    # Finance: amortise ROU + separate interest
                    amort = straight_line_monthly
                    rou = rou - amort
                    liability = liability + interest - eff_rent
                    expense = amort
                    year_interest += interest

                year_expense += expense
                year_cash += eff_rent

            # Prevent negative balances from rounding
            rou = max(Decimal(0), rou)
            liability = max(Decimal(0), liability)

            year_expense = year_expense.quantize(Decimal("1"), rounding=ROUND_HALF_UP)
            year_interest = year_interest.quantize(Decimal("1"), rounding=ROUND_HALF_UP)
            year_cash = year_cash.quantize(Decimal("1"), rounding=ROUND_HALF_UP)

            rows.append(LeaseScheduleRow(
                lease_id=lease.lease_id,
                entity_code=lease.entity_code,
                year=year,
                lease_type=lease.lease_type,
                rou_asset_beg=rou_beg.quantize(Decimal("1"), rounding=ROUND_HALF_UP),
                rou_asset_end=rou.quantize(Decimal("1"), rounding=ROUND_HALF_UP),
                lease_liability_beg=liab_beg.quantize(Decimal("1"), rounding=ROUND_HALF_UP),
                lease_liability_end=liability.quantize(Decimal("1"), rounding=ROUND_HALF_UP),
                lease_expense=year_expense,
                interest_expense=year_interest,
                cash_paid=year_cash,
            ))

    rows.sort(key=lambda r: (r.lease_id, r.year))
    return rows


def _months_active_in_year(lease: Lease, year: int) -> int:
    """Number of months a lease is active during a calendar year."""
    year_start = datetime.date(year, 1, 1)
    year_end = datetime.date(year, 12, 31)

    lease_start = lease.commencement_date
    lease_end = lease.end_date

    if lease_end < year_start or lease_start > year_end:
        return 0

    eff_start = max(lease_start, year_start)
    eff_end = min(lease_end, year_end)

    # Count months: from eff_start.month to eff_end.month inclusive
    return (eff_end.year - eff_start.year) * 12 + eff_end.month - eff_start.month + 1


# ── TC-06 interface: ASC 842 temporary difference ──────────────────────────

def asc842_temp_difference(
    schedules: list[LeaseScheduleRow],
    year: int,
    entity_code: str | None = None,
) -> Decimal:
    """ASC 842 temporary difference for a given year.

    The temp difference is the net of (ROU asset - Lease liability) at year end.
    Under ASC 842, this net is typically a small debit (the ROU asset exceeds
    the liability early on for operating leases) creating a deferred tax asset
    or liability depending on direction.

    For tax purposes, lease payments are deductible when paid (i.e. the cash
    amount), so the book-tax difference = book lease expense - cash paid.
    Cumulative difference drives the DTA/DTL.

    Returns: book_expense - cash_paid for the year (positive = book > cash = DTA).
    """
    total = Decimal(0)
    for row in schedules:
        if row.year != year:
            continue
        if entity_code is not None and row.entity_code != entity_code:
            continue
        # Book expense vs. cash: the difference is the temporary difference
        total += row.lease_expense + row.interest_expense - row.cash_paid
    return total.quantize(Decimal("1"), rounding=ROUND_HALF_UP)


# ── GL posting ──────────────────────────────────────────────────────────────

def post_leases_to_gl(
    ledger: Ledger,
    leases: list[Lease],
    schedules: list[LeaseScheduleRow],
    years: list[int] | None = None,
) -> None:
    """Post ASC 842 lease entries to the GL.

    1. At commencement (or amendment date): recognise ROU asset and liability.
    2. Annually: post lease expense (operating) or amortisation + interest (finance).
    3. Annually: reduce liability and ROU per schedule.

    Short-term exempt leases: expense = cash paid to 6100 (Rent Expense).
    """
    if years is None:
        years = [2023, 2024, 2025]

    # Post initial recognition entries (commencement or amendment)
    for lease in leases:
        if lease.short_term_exempt:
            continue

        # Initial recognition: DR ROU asset, CR Lease liability
        recognise_date = lease.commencement_date
        rou_acct = "1600" if lease.lease_type == LeaseType.OPERATING else "1610"
        liab_acct = "2310" if lease.lease_type == LeaseType.OPERATING else "2330"

        # Only post if commencement is within or before our year range
        if recognise_date.year <= max(years):
            entry = JournalEntry(
                date=recognise_date,
                entity_code=lease.entity_code,
                description=f"ASC 842 initial recognition — {lease.description}",
                lines=(
                    JournalEntryLine(
                        account=rou_acct,
                        debit=lease.rou_asset_initial,
                        credit=Decimal(0),
                        memo=f"ROU asset — {lease.lease_id}",
                    ),
                    JournalEntryLine(
                        account=liab_acct,
                        debit=Decimal(0),
                        credit=lease.lease_liability_initial,
                        memo=f"Lease liability — {lease.lease_id}",
                    ),
                ),
            )
            ledger.post(entry)

    # Post annual schedule entries
    schedule_by_key: dict[tuple[str, int], LeaseScheduleRow] = {
        (row.lease_id, row.year): row for row in schedules
    }

    for lease in leases:
        for year in sorted(years):
            key = (lease.lease_id, year)
            row = schedule_by_key.get(key)
            if row is None:
                continue

            date = datetime.date(year, 12, 31)

            if lease.short_term_exempt:
                # Simple rent expense
                if row.cash_paid > 0:
                    entry = JournalEntry(
                        date=date,
                        entity_code=lease.entity_code,
                        description=f"Short-term lease rent — {lease.description} FY{year}",
                        lines=(
                            JournalEntryLine(
                                account="6100",
                                debit=row.cash_paid,
                                credit=Decimal(0),
                                memo=f"Short-term lease expense — {lease.lease_id}",
                            ),
                            JournalEntryLine(
                                account="1010",
                                debit=Decimal(0),
                                credit=row.cash_paid,
                                memo=f"Rent payment — {lease.lease_id}",
                            ),
                        ),
                    )
                    ledger.post(entry)
                continue

            # Non-exempt leases: post expense, liability reduction, ROU amortisation
            rou_acct = "1600" if lease.lease_type == LeaseType.OPERATING else "1610"
            rou_amort_acct = "1620" if lease.lease_type == LeaseType.OPERATING else "1630"
            liab_noncurrent = "2310" if lease.lease_type == LeaseType.OPERATING else "2330"

            lines: list[JournalEntryLine] = []

            if lease.lease_type == LeaseType.OPERATING:
                # DR Lease expense (6100), CR Cash, and adjust ROU/liability
                if row.lease_expense > 0:
                    lines.append(JournalEntryLine(
                        account="6100", debit=row.lease_expense, credit=Decimal(0),
                        memo=f"Operating lease expense — {lease.lease_id}",
                    ))
                if row.cash_paid > 0:
                    lines.append(JournalEntryLine(
                        account="1010", debit=Decimal(0), credit=row.cash_paid,
                        memo=f"Lease payment — {lease.lease_id}",
                    ))
                # Balance: expense vs cash difference goes to ROU/liability adjustment
                diff = row.cash_paid - row.lease_expense
                if diff > 0:
                    # Cash > expense: reduce liability more than ROU
                    lines.append(JournalEntryLine(
                        account=liab_noncurrent, debit=diff, credit=Decimal(0),
                        memo=f"Lease liability reduction — {lease.lease_id}",
                    ))
                elif diff < 0:
                    lines.append(JournalEntryLine(
                        account=liab_noncurrent, debit=Decimal(0), credit=abs(diff),
                        memo=f"Lease liability accrual — {lease.lease_id}",
                    ))

                # ROU amortisation (contra-asset)
                rou_amort = row.rou_asset_beg - row.rou_asset_end
                if rou_amort > 0:
                    lines.append(JournalEntryLine(
                        account=rou_amort_acct, debit=Decimal(0), credit=rou_amort,
                        memo=f"ROU amortisation — {lease.lease_id}",
                    ))

            else:
                # Finance lease: separate amortisation + interest
                if row.lease_expense > 0:
                    lines.append(JournalEntryLine(
                        account="6210", debit=row.lease_expense, credit=Decimal(0),
                        memo=f"Finance lease amortisation — {lease.lease_id}",
                    ))
                    lines.append(JournalEntryLine(
                        account=rou_amort_acct, debit=Decimal(0), credit=row.lease_expense,
                        memo=f"ROU amortisation — {lease.lease_id}",
                    ))
                if row.interest_expense > 0:
                    lines.append(JournalEntryLine(
                        account="7020", debit=row.interest_expense, credit=Decimal(0),
                        memo=f"Finance lease interest — {lease.lease_id}",
                    ))
                if row.cash_paid > 0:
                    # Cash payment reduces liability
                    principal_paid = row.cash_paid - row.interest_expense
                    if principal_paid > 0:
                        lines.append(JournalEntryLine(
                            account=liab_noncurrent, debit=principal_paid, credit=Decimal(0),
                            memo=f"Lease liability principal — {lease.lease_id}",
                        ))
                    lines.append(JournalEntryLine(
                        account="1010", debit=Decimal(0), credit=row.cash_paid,
                        memo=f"Lease payment — {lease.lease_id}",
                    ))

            # Filter out zero-amount lines and lines with zero debit AND zero credit
            lines = [ln for ln in lines if ln.debit > 0 or ln.credit > 0]

            if not lines:
                continue

            # Balance check — add rounding adjustment if needed
            total_dr = sum(ln.debit for ln in lines)
            total_cr = sum(ln.credit for ln in lines)
            diff = total_dr - total_cr
            if diff > 0:
                lines.append(JournalEntryLine(
                    account=rou_amort_acct, debit=Decimal(0), credit=diff,
                    memo=f"Rounding adjustment — {lease.lease_id}",
                ))
            elif diff < 0:
                lines.append(JournalEntryLine(
                    account=rou_amort_acct, debit=abs(diff), credit=Decimal(0),
                    memo=f"Rounding adjustment — {lease.lease_id}",
                ))

            entry = JournalEntry(
                date=date,
                entity_code=lease.entity_code,
                description=f"ASC 842 — {lease.description} FY{year}",
                lines=tuple(lines),
            )
            ledger.post(entry)
