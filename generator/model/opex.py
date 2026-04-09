"""Cascade Industries operating expense generators (SG&A, payroll, R&D).

Generates monthly operating expenses by entity for FY2023–FY2025. Payroll is
derived from the employee roster; R&D spend for Advanced Materials targets
~12% of its revenue. Posts journal entries to the GL.

Key constraints from the spec:
- Advanced Materials R&D expense ≈ 12% of its revenue.
- SG&A distributed by department weight across entities.
- Payroll is the dominant opex category, derived from actual employee salaries.
"""

from __future__ import annotations

import datetime
import random
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal

from generator.model.employees import Employee
from generator.model.gl import JournalEntry, JournalEntryLine, Ledger

# ── Monthly allocation weights (mirrors revenue seasonality loosely) ─────────
# Opex is flatter than revenue but still has some Q4 bump from bonuses/travel.

_MONTHLY_WEIGHTS: tuple[float, ...] = (
    0.078, 0.078, 0.080,   # Q1 — 23.6%
    0.082, 0.083, 0.083,   # Q2 — 24.8%
    0.083, 0.083, 0.082,   # Q3 — 24.8%
    0.084, 0.086, 0.098,   # Q4 — 26.8% (Dec bump: bonuses, year-end)
)

assert abs(sum(_MONTHLY_WEIGHTS) - 1.0) < 1e-9, "Monthly weights must sum to 1.0"


# ── Opex category definitions ────────────────────────────────────────────────
# Each category maps to a GL account and has an annual budget as a fraction
# of revenue (except payroll, which is derived from the employee roster).

@dataclass(frozen=True)
class OpexCategory:
    """A non-payroll operating expense category."""

    name: str
    gl_account: str  # 6xxx GL account
    pct_of_revenue: dict[str, float]  # entity_code → annual spend as % of revenue


# Non-payroll, non-R&D opex categories with entity-specific rates.
# These rates are calibrated so total opex (payroll + these + R&D) is realistic
# for a mid-market manufacturer.
OPEX_CATEGORIES: tuple[OpexCategory, ...] = (
    OpexCategory("Benefits", "6030", {"PC": 0.022, "AM": 0.025, "DS": 0.020}),
    OpexCategory("Payroll Taxes", "6040", {"PC": 0.012, "AM": 0.013, "DS": 0.011}),
    OpexCategory("Rent Expense", "6100", {"PC": 0.008, "AM": 0.006, "DS": 0.010}),
    OpexCategory("Utilities", "6110", {"PC": 0.005, "AM": 0.004, "DS": 0.006}),
    OpexCategory("Office Supplies", "6120", {"PC": 0.001, "AM": 0.001, "DS": 0.001}),
    OpexCategory("Telecom", "6130", {"PC": 0.001, "AM": 0.002, "DS": 0.001}),
    OpexCategory("Insurance", "6140", {"PC": 0.004, "AM": 0.003, "DS": 0.003}),
    OpexCategory("Travel", "6150", {"PC": 0.003, "AM": 0.005, "DS": 0.002}),
    OpexCategory("Meals & Entertainment", "6160", {"PC": 0.002, "AM": 0.003, "DS": 0.001}),
    OpexCategory("Legal Fees", "6170", {"PC": 0.002, "AM": 0.003, "DS": 0.001}),
    OpexCategory("Audit Fees", "6180", {"PC": 0.001, "AM": 0.001, "DS": 0.001}),
    OpexCategory("Consulting", "6190", {"PC": 0.002, "AM": 0.003, "DS": 0.001}),
    OpexCategory("Advertising", "6200", {"PC": 0.003, "AM": 0.004, "DS": 0.002}),
    OpexCategory("Bad Debt", "6230", {"PC": 0.003, "AM": 0.002, "DS": 0.002}),
    OpexCategory("Repairs & Maintenance", "6240", {"PC": 0.005, "AM": 0.003, "DS": 0.004}),
    OpexCategory("Shipping", "6250", {"PC": 0.004, "AM": 0.002, "DS": 0.008}),
    OpexCategory("Software", "6260", {"PC": 0.002, "AM": 0.003, "DS": 0.001}),
    OpexCategory("Training", "6270", {"PC": 0.001, "AM": 0.002, "DS": 0.001}),
    OpexCategory("Warranty", "6400", {"PC": 0.004, "AM": 0.003, "DS": 0.000}),
)


# ── R&D account mapping ────────────────────────────────────────────────────────
# R&D is only for Advanced Materials. We split the 12% target across
# granular 63xx accounts for TC-08 (Section 41 credit study).
# Fractions of total R&D spend by account.

_RD_ACCOUNT_SPLITS: tuple[tuple[str, float], ...] = (
    ("6300", 0.45),  # R&D Salaries & Wages (derived from roster)
    ("6310", 0.18),  # Contract Research
    ("6320", 0.10),  # Supplies & Materials
    ("6330", 0.08),  # Equipment Depreciation (cross-posted from PPE)
    ("6340", 0.06),  # Software & Tools
    ("6350", 0.05),  # Testing & Certification
    ("6360", 0.04),  # Travel
    ("6370", 0.04),  # Other
)

assert abs(sum(s for _, s in _RD_ACCOUNT_SPLITS) - 1.0) < 1e-9


# ── Data records ──────────────────────────────────────────────────────────────

@dataclass
class MonthlyOpex:
    """A single month's operating expense for one entity × category."""

    year: int
    month: int
    entity_code: str
    category: str
    gl_account: str
    amount: Decimal  # Positive


# ── Payroll computation ──────────────────────────────────────────────────────

def _compute_annual_payroll(
    employees: list[Employee],
    entity_code: str,
    year: int,
    rd_salary_cap: Decimal | None = None,
) -> dict[str, Decimal]:
    """Compute annual payroll cost by GL account for an entity/year.

    Returns {"6010": admin_payroll, "6020": sales_payroll, "6300": rd_payroll}.
    Active employees in the given year contribute their annual_salary prorated
    by months active in that year.

    If ``rd_salary_cap`` is provided, the 6300 (R&D salaries) total is capped
    at that amount and any excess is reclassified to 6010 (admin salaries).
    This ensures the R&D-eligible roster doesn't push total R&D spend above
    the 12% target.
    """
    year_start = datetime.date(year, 1, 1)
    year_end = datetime.date(year, 12, 31)

    payroll: dict[str, Decimal] = {}

    for emp in employees:
        if emp.entity_code != entity_code:
            continue

        # Determine months active in this year.
        active_start = max(emp.hire_date, year_start)
        active_end = year_end
        if emp.termination_date is not None:
            active_end = min(emp.termination_date, year_end)

        if active_start > active_end:
            continue

        # Prorate: count months (rounded to nearest month).
        months_active = (active_end.month - active_start.month + 1 +
                         12 * (active_end.year - active_start.year))
        months_active = max(1, min(12, months_active))

        prorated = Decimal(emp.annual_salary) * Decimal(months_active) / Decimal(12)
        prorated = prorated.quantize(Decimal("1"), rounding=ROUND_HALF_UP)

        # Route to appropriate GL account.
        if emp.department in ("R&D", "Engineering") and emp.is_rd_eligible:
            acct = "6300"  # R&D Salaries
        elif emp.department == "Sales":
            acct = "6020"  # Sales Salaries
        else:
            acct = "6010"  # Admin Salaries

        payroll[acct] = payroll.get(acct, Decimal(0)) + prorated

    # Cap R&D salary if needed (excess → admin salaries).
    if rd_salary_cap is not None and payroll.get("6300", Decimal(0)) > rd_salary_cap:
        excess = payroll["6300"] - rd_salary_cap
        payroll["6300"] = rd_salary_cap
        payroll["6010"] = payroll.get("6010", Decimal(0)) + excess

    return payroll


# ── Stock-based compensation ─────────────────────────────────────────────────
# Permanent difference for TC-06. Roughly 0.8% of revenue for AM, less for others.

_SBC_RATES: dict[str, float] = {
    "PC": 0.004,
    "AM": 0.008,
    "DS": 0.003,
}


# ── Main generator ────────────────────────────────────────────────────────────

def generate_opex(
    rng: random.Random,
    employees: list[Employee],
    revenue_by_entity_year: dict[tuple[str, int], Decimal],
    years: list[int] | None = None,
) -> list[MonthlyOpex]:
    """Generate monthly operating expense records for all entities and years.

    Args:
        rng: Seeded random.Random for determinism.
        employees: Full employee roster from generate_employees().
        revenue_by_entity_year: {(entity_code, year): annual_revenue} for
            computing percentage-of-revenue opex categories.
        years: Fiscal years to generate (default: [2023, 2024, 2025]).

    Returns:
        Sorted list of MonthlyOpex records.
    """
    if years is None:
        years = [2023, 2024, 2025]

    records: list[MonthlyOpex] = []

    for entity_code in sorted({"PC", "AM", "DS"}):
        for year in sorted(years):
            annual_rev = revenue_by_entity_year.get(
                (entity_code, year), Decimal(0)
            )
            if annual_rev <= 0:
                continue

            # ── Payroll ──────────────────────────────────────────────────
            # For AM, cap R&D salary at 45% of the 12% R&D target so that
            # the roster doesn't push total R&D spend above the target.
            rd_cap = None
            if entity_code == "AM":
                rd_target = (annual_rev * Decimal("0.12")).quantize(
                    Decimal("1"), rounding=ROUND_HALF_UP
                )
                rd_cap = (rd_target * Decimal("0.45")).quantize(
                    Decimal("1"), rounding=ROUND_HALF_UP
                )

            payroll_by_acct = _compute_annual_payroll(
                employees, entity_code, year, rd_salary_cap=rd_cap
            )
            for acct, annual_amount in sorted(payroll_by_acct.items()):
                records.extend(
                    _spread_to_months(rng, year, entity_code, f"Payroll ({acct})",
                                      acct, annual_amount)
                )

            # ── Non-payroll opex categories ──────────────────────────────
            for cat in OPEX_CATEGORIES:
                rate = cat.pct_of_revenue.get(entity_code, 0.0)
                if rate <= 0:
                    continue
                annual_amount = (annual_rev * Decimal(str(rate))).quantize(
                    Decimal("1"), rounding=ROUND_HALF_UP
                )
                records.extend(
                    _spread_to_months(rng, year, entity_code, cat.name,
                                      cat.gl_account, annual_amount)
                )

            # ── R&D (Advanced Materials only) ────────────────────────────
            if entity_code == "AM":
                # Salary portion (6300) already posted via payroll above,
                # capped at 45% of the 12% target. Non-salary R&D accounts
                # get the remaining 55%.
                non_salary_rd = rd_target - rd_cap  # type: ignore[operator]
                non_salary_splits = [(a, s) for a, s in _RD_ACCOUNT_SPLITS if a != "6300"]
                total_ns_weight = sum(s for _, s in non_salary_splits)

                for acct, share in non_salary_splits:
                    acct_amount = (non_salary_rd * Decimal(str(share / total_ns_weight))).quantize(
                        Decimal("1"), rounding=ROUND_HALF_UP
                    )
                    records.extend(
                        _spread_to_months(rng, year, entity_code,
                                          f"R&D ({acct})", acct, acct_amount)
                    )

            # ── Stock-based compensation (permanent difference) ──────────
            sbc_rate = _SBC_RATES.get(entity_code, 0.0)
            if sbc_rate > 0:
                sbc_amount = (annual_rev * Decimal(str(sbc_rate))).quantize(
                    Decimal("1"), rounding=ROUND_HALF_UP
                )
                records.extend(
                    _spread_to_months(rng, year, entity_code, "Stock Comp",
                                      "6050", sbc_amount)
                )

    records.sort(key=lambda r: (r.year, r.month, r.entity_code, r.category))
    return records


def _spread_to_months(
    rng: random.Random,
    year: int,
    entity_code: str,
    category: str,
    gl_account: str,
    annual_amount: Decimal,
) -> list[MonthlyOpex]:
    """Spread an annual amount across 12 months using weights + perturbation."""
    if annual_amount <= 0:
        return []

    # Perturb weights ±2% for realism
    raw = [w * (1 + rng.uniform(-0.02, 0.02)) for w in _MONTHLY_WEIGHTS]
    total_w = sum(raw)
    normed = [w / total_w for w in raw]

    records: list[MonthlyOpex] = []
    allocated = Decimal(0)

    for i, mw in enumerate(normed):
        if i == 11:
            amount = annual_amount - allocated
        else:
            amount = (annual_amount * Decimal(str(mw))).quantize(
                Decimal("1"), rounding=ROUND_HALF_UP
            )
            allocated += amount

        records.append(MonthlyOpex(
            year=year,
            month=i + 1,
            entity_code=entity_code,
            category=category,
            gl_account=gl_account,
            amount=amount,
        ))

    return records


# ── GL posting ────────────────────────────────────────────────────────────────

def post_opex_to_gl(
    ledger: Ledger,
    records: list[MonthlyOpex],
) -> None:
    """Post operating expense records as journal entries.

    DR: opex GL account
    CR: Accrued Expenses (2020) or Accrued Payroll (2030) for payroll items.
    """
    for rec in records:
        if rec.amount <= 0:
            continue

        date = datetime.date(rec.year, rec.month, 28)  # End-of-month posting

        # Payroll-related accounts credit Accrued Payroll; others credit Accrued Expenses.
        cr_account = "2030" if rec.gl_account in ("6010", "6020", "6300") else "2020"

        entry = JournalEntry(
            date=date,
            entity_code=rec.entity_code,
            description=f"Opex — {rec.category} {rec.year}-{rec.month:02d}",
            lines=(
                JournalEntryLine(
                    account=rec.gl_account,
                    debit=rec.amount,
                    credit=Decimal(0),
                    memo=rec.category,
                ),
                JournalEntryLine(
                    account=cr_account,
                    debit=Decimal(0),
                    credit=rec.amount,
                    memo=rec.category,
                ),
            ),
        )
        ledger.post(entry)


# ── Validation helpers ────────────────────────────────────────────────────────

def validate_rd_spend(
    records: list[MonthlyOpex],
    am_revenue: Decimal,
    year: int = 2025,
) -> Decimal:
    """Return total R&D spend for AM in a given year as a fraction of revenue.

    Includes all 63xx accounts plus any payroll routed to 6300.
    """
    rd_total = sum(
        r.amount for r in records
        if r.entity_code == "AM"
        and r.year == year
        and r.gl_account.startswith("63")
    )
    if am_revenue <= 0:
        return Decimal(0)
    return rd_total / am_revenue
