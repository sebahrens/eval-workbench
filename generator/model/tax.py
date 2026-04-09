"""Cascade Industries tax provision model (ASC 740) for TC-06.

Computes the income tax provision for FY2023–FY2025 including:
- Pre-tax book income from the GL
- Permanent differences (M&E 50%, tax-exempt interest, stock comp excess, fines)
- Temporary differences (depreciation, ASC 842, warranty, inventory, bonuses, bad debt)
- DTA/DTL rollforward
- Federal 21% + blended state 6.2%
- R&D credit (Section 41)
- Effective tax rate reconciliation

Target effective tax rate ≈ 24.8% for FY2025 per prompt.md TC-06.

Determinism: all inputs come from existing model modules; no randomness needed here.
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from generator.model.gl import Ledger
    from generator.model.leases import LeaseScheduleRow
    from generator.model.opex import MonthlyOpex
    from generator.model.ppe import FixedAsset


# ── Tax rates ────────────────────────────────────────────────────────────────

FEDERAL_RATE = Decimal("0.21")
STATE_RATE = Decimal("0.062")
# Combined rate: federal applies to pre-state income, state is deductible for federal
# Effective combined = federal × (1 - state) + state = 0.21 × 0.938 + 0.062 = 0.25898
# But for provision purposes we compute separately.


# ── Permanent differences ────────────────────────────────────────────────────

@dataclass(frozen=True)
class PermanentDifference:
    """A permanent book-tax difference."""

    description: str
    amount: Decimal  # Positive = increases taxable income; negative = decreases


@dataclass(frozen=True)
class TemporaryDifference:
    """A temporary book-tax difference for a given year."""

    description: str
    book_amount: Decimal  # Book basis amount for the year
    tax_amount: Decimal   # Tax basis amount for the year

    @property
    def difference(self) -> Decimal:
        """Current year difference (positive = taxable > book = reduces DTA or increases DTL)."""
        return self.tax_amount - self.book_amount

    @property
    def cumulative_type(self) -> str:
        """Whether this creates a DTA or DTL when cumulative difference is positive."""
        return "DTL" if self.difference > 0 else "DTA"


@dataclass(frozen=True)
class DeferredTaxItem:
    """A deferred tax asset or liability for a specific temporary difference."""

    description: str
    cumulative_difference: Decimal  # Cumulative book-tax difference (positive = taxable > book)
    deferred_tax: Decimal           # Positive = DTL, negative = DTA
    item_type: str                  # "DTA" or "DTL"


@dataclass
class TaxProvision:
    """Complete tax provision for a single year."""

    year: int
    pre_tax_book_income: Decimal

    # Permanent differences
    permanent_differences: list[PermanentDifference]
    total_permanent: Decimal

    # Temporary differences (current year changes)
    temporary_differences: list[TemporaryDifference]
    total_temporary_change: Decimal

    # Taxable income
    taxable_income: Decimal

    # Current provision
    federal_current: Decimal
    state_current: Decimal
    rd_credit: Decimal
    total_current: Decimal

    # Deferred provision (change in DTA/DTL from prior year)
    deferred_items: list[DeferredTaxItem]
    total_deferred: Decimal

    # Prior year DTA/DTL for rollforward
    prior_dta_total: Decimal
    prior_dtl_total: Decimal
    current_dta_total: Decimal
    current_dtl_total: Decimal

    # Total provision and ETR
    total_provision: Decimal
    effective_tax_rate: Decimal

    # Rate reconciliation items
    rate_reconciliation: list[tuple[str, Decimal, Decimal]]  # (description, amount, rate impact)


# ── Permanent difference computation ─────────────────────────────────────────

def _compute_permanent_differences(
    opex_records: list[MonthlyOpex],
    year: int,
) -> list[PermanentDifference]:
    """Compute permanent differences from the opex model.

    Permanent differences:
    1. Meals & Entertainment: 50% non-deductible (increases taxable income)
    2. Tax-exempt interest: reduces taxable income
    3. Stock-based compensation excess: book expense exceeds tax deduction
    4. Fines & penalties: 100% non-deductible
    """
    perms: list[PermanentDifference] = []

    # 1. Meals & Entertainment (GL 6160) — 50% non-deductible
    me_total = sum(
        r.amount for r in opex_records
        if r.year == year and r.gl_account == "6160"
    )
    if me_total > 0:
        me_nondeductible = (me_total * Decimal("0.50")).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP
        )
        perms.append(PermanentDifference(
            description="Meals & entertainment (50% non-deductible)",
            amount=me_nondeductible,
        ))

    # 2. Tax-exempt interest income (GL 7080)
    # This is a credit-balance account, so it appears as income on books
    # but is excluded from taxable income → decreases taxable income.
    # We generate a fixed amount per year (from municipal bond portfolio).
    # FY2025: $180K; FY2024: $170K; FY2023: $160K
    _tax_exempt_by_year = {2023: Decimal("160000"), 2024: Decimal("170000"), 2025: Decimal("180000")}
    tax_exempt = _tax_exempt_by_year.get(year, Decimal(0))
    if tax_exempt > 0:
        perms.append(PermanentDifference(
            description="Tax-exempt municipal bond interest",
            amount=-tax_exempt,  # Negative = decreases taxable income
        ))

    # 3. Stock-based compensation excess deduction
    # Book expense recorded in GL 6050. For tax, only the intrinsic value at
    # exercise is deductible. The excess (book > tax) is a permanent addback.
    # We model the tax deduction as 70% of book expense (30% permanent).
    sbc_total = sum(
        r.amount for r in opex_records
        if r.year == year and r.gl_account == "6050"
    )
    if sbc_total > 0:
        sbc_permanent = (sbc_total * Decimal("0.30")).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP
        )
        perms.append(PermanentDifference(
            description="Stock compensation excess (book > tax deduction)",
            amount=sbc_permanent,
        ))

    # 4. Fines & penalties (GL 6280) — 100% non-deductible
    fines_total = sum(
        r.amount for r in opex_records
        if r.year == year and r.gl_account == "6280"
    )
    if fines_total > 0:
        perms.append(PermanentDifference(
            description="Fines & penalties (non-deductible)",
            amount=fines_total,
        ))

    return perms


# ── Temporary difference computation ─────────────────────────────────────────

def _compute_temporary_differences(
    assets: list[FixedAsset],
    lease_schedules: list[LeaseScheduleRow],
    opex_records: list[MonthlyOpex],
    year: int,
) -> list[TemporaryDifference]:
    """Compute temporary differences for a given year.

    Temporary differences:
    1. Depreciation (MACRS vs. straight-line) — DTL
    2. ASC 842 lease adjustments — DTA or DTL
    3. Warranty reserve (book accrual vs. tax deduction when paid) — DTA
    4. Inventory reserve (book LCM write-down vs. tax basis) — DTA
    5. Accrued bonuses (book accrual vs. tax deduction when paid) — DTA
    6. Bad debt reserve (book provision vs. tax write-off when uncollectible) — DTA
    """
    temps: list[TemporaryDifference] = []

    # 1. Depreciation: book (straight-line) vs tax (MACRS)
    book_depr = sum(a.book_depr_for_year(year) for a in assets)
    tax_depr = sum(a.tax_depr_for_year(year) for a in assets)
    temps.append(TemporaryDifference(
        description="Depreciation (MACRS vs. straight-line)",
        book_amount=book_depr,
        tax_amount=tax_depr,
    ))

    # 2. ASC 842 lease adjustments
    # Book: straight-line lease expense; Tax: cash rent deduction
    # Import here to avoid circular import at module level
    # Book: straight-line lease expense; Tax: cash rent deduction
    # For our TemporaryDifference: book_amount = lease expense, tax_amount = cash paid
    total_lease_expense = sum(
        r.lease_expense + r.interest_expense for r in lease_schedules if r.year == year
    )
    total_cash_paid = sum(
        r.cash_paid for r in lease_schedules if r.year == year
    )
    temps.append(TemporaryDifference(
        description="ASC 842 lease adjustments",
        book_amount=total_lease_expense,
        tax_amount=total_cash_paid,
    ))

    # 3. Warranty reserve — book accrues warranty expense (6400);
    # tax only deducts when claims are actually paid.
    # Model: book accrual = opex 6400; tax deduction = 85% of accrual (15% still in reserve)
    warranty_book = sum(
        r.amount for r in opex_records
        if r.year == year and r.gl_account == "6400"
    )
    warranty_tax = (warranty_book * Decimal("0.85")).quantize(
        Decimal("1"), rounding=ROUND_HALF_UP
    )
    temps.append(TemporaryDifference(
        description="Warranty reserve",
        book_amount=warranty_book,
        tax_amount=warranty_tax,
    ))

    # 4. Inventory reserve — book LCM write-down (5050 contains inventory adjustments);
    # tax doesn't allow until actual disposal.
    # Model: 40% of inventory adjustments remain as a book-only reserve.
    # The opex model doesn't generate standalone inventory reserve entries,
    # so we use a fixed schedule per year.
    _inv_reserve_change = {2023: Decimal("120000"), 2024: Decimal("135000"), 2025: Decimal("150000")}
    inv_book = _inv_reserve_change.get(year, Decimal(0))
    inv_tax = Decimal(0)  # Tax: no deduction until disposal
    temps.append(TemporaryDifference(
        description="Inventory obsolescence reserve",
        book_amount=inv_book,
        tax_amount=inv_tax,
    ))

    # 5. Accrued bonuses — book accrues in current year; tax deducts when paid (next year).
    # Model: total bonus accrual = sum of December payroll bump across entities × 15%.
    # Use a fixed schedule that aligns with the opex model's year-end bump.
    _bonus_accrual = {2023: Decimal("1800000"), 2024: Decimal("1950000"), 2025: Decimal("2100000")}
    bonus_book = _bonus_accrual.get(year, Decimal(0))
    bonus_tax = Decimal(0)  # Tax: deducted when paid (next year)
    temps.append(TemporaryDifference(
        description="Accrued bonuses",
        book_amount=bonus_book,
        tax_amount=bonus_tax,
    ))

    # 6. Bad debt reserve — book provision (6230); tax deducts specific write-offs.
    # Model: tax deduction = 60% of book provision (rest stays in reserve).
    bad_debt_book = sum(
        r.amount for r in opex_records
        if r.year == year and r.gl_account == "6230"
    )
    bad_debt_tax = (bad_debt_book * Decimal("0.60")).quantize(
        Decimal("1"), rounding=ROUND_HALF_UP
    )
    temps.append(TemporaryDifference(
        description="Bad debt reserve",
        book_amount=bad_debt_book,
        tax_amount=bad_debt_tax,
    ))

    return temps


# ── R&D credit computation ───────────────────────────────────────────────────

def _compute_rd_credit(
    opex_records: list[MonthlyOpex],
    year: int,
) -> Decimal:
    """Compute Section 41 R&D tax credit.

    Simplified: 20% of qualified research expenses (QREs) above a base amount.
    QREs = all 63xx account expenses for Advanced Materials.
    Base = 50% of average of prior 3 years' QREs (or current year if no history).
    For simplicity in our model, credit ≈ 6.5% of current-year QREs.
    """
    qres = sum(
        r.amount for r in opex_records
        if r.year == year
        and r.entity_code == "AM"
        and r.gl_account.startswith("63")
    )
    # Simplified credit rate: ~6.5% of QREs
    credit = (qres * Decimal("0.065")).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    return credit


# ── Main provision computation ───────────────────────────────────────────────

def compute_tax_provision(
    pre_tax_book_income: Decimal,
    opex_records: list[MonthlyOpex],
    assets: list[FixedAsset],
    lease_schedules: list[LeaseScheduleRow],
    year: int,
    prior_provision: TaxProvision | None = None,
) -> TaxProvision:
    """Compute the full ASC 740 tax provision for a given year.

    Args:
        pre_tax_book_income: Pre-tax income from the trial balance.
        opex_records: All MonthlyOpex records (all years needed for rollforward).
        assets: All FixedAsset records.
        lease_schedules: All LeaseScheduleRow records.
        year: The fiscal year to compute.
        prior_provision: Prior year's TaxProvision (for DTA/DTL rollforward).
            None for first year (FY2023).

    Returns:
        Complete TaxProvision for the year.
    """
    # ── Permanent differences ────────────────────────────────────────
    perms = _compute_permanent_differences(opex_records, year)
    total_permanent = sum(p.amount for p in perms)

    # ── Temporary differences ────────────────────────────────────────
    temps = _compute_temporary_differences(assets, lease_schedules, opex_records, year)
    # Total temporary change = sum of (tax - book) for each difference
    total_temporary_change = sum(t.difference for t in temps)

    # ── Taxable income ───────────────────────────────────────────────
    taxable_income = pre_tax_book_income + total_permanent + total_temporary_change

    # ── Current provision ────────────────────────────────────────────
    # State tax is deductible for federal purposes.
    state_current = (taxable_income * STATE_RATE).quantize(
        Decimal("1"), rounding=ROUND_HALF_UP
    )
    federal_taxable = taxable_income - state_current
    federal_current = (federal_taxable * FEDERAL_RATE).quantize(
        Decimal("1"), rounding=ROUND_HALF_UP
    )

    # R&D credit
    rd_credit = _compute_rd_credit(opex_records, year)

    total_current = federal_current + state_current - rd_credit

    # ── Deferred provision (change in DTA/DTL) ───────────────────────
    # Build cumulative temporary differences through current year.
    # For each temp diff category, compute cumulative position.
    combined_rate = FEDERAL_RATE + STATE_RATE - (FEDERAL_RATE * STATE_RATE)
    # Actually: use simple combined rate for deferred items
    # Combined = federal + state × (1 - federal) ... but for simplicity:
    combined_rate = (FEDERAL_RATE + STATE_RATE * (Decimal(1) - FEDERAL_RATE)).quantize(
        Decimal("0.0001"), rounding=ROUND_HALF_UP
    )

    deferred_items: list[DeferredTaxItem] = []
    for temp in temps:
        # Cumulative difference: we need to know prior year cumulative.
        # For year 1, cumulative = current year difference.
        # For later years, we build from prior provision's deferred items.
        current_diff = temp.difference  # tax - book for this year
        prior_cumulative = Decimal(0)

        if prior_provision is not None:
            # Find matching item in prior year
            for prior_item in prior_provision.deferred_items:
                if prior_item.description == temp.description:
                    # Reverse-engineer cumulative from deferred_tax / combined_rate
                    if combined_rate != 0:
                        prior_cumulative = (prior_item.deferred_tax / combined_rate).quantize(
                            Decimal("1"), rounding=ROUND_HALF_UP
                        )
                    break

        cumulative = prior_cumulative + current_diff
        deferred_tax = (cumulative * combined_rate).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP
        )

        item_type = "DTL" if deferred_tax > 0 else "DTA"

        deferred_items.append(DeferredTaxItem(
            description=temp.description,
            cumulative_difference=cumulative,
            deferred_tax=deferred_tax,
            item_type=item_type,
        ))

    # Compute deferred provision = change in net deferred tax from prior year
    current_dta = sum(abs(d.deferred_tax) for d in deferred_items if d.item_type == "DTA")
    current_dtl = sum(d.deferred_tax for d in deferred_items if d.item_type == "DTL")

    prior_dta = prior_provision.current_dta_total if prior_provision else Decimal(0)
    prior_dtl = prior_provision.current_dtl_total if prior_provision else Decimal(0)

    # Deferred expense = increase in DTL + decrease in DTA (net)
    # Or equivalently: change in (DTL - DTA) = deferred tax expense
    total_deferred = (current_dtl - current_dta) - (prior_dtl - prior_dta)

    # ── Total provision ──────────────────────────────────────────────
    total_provision = total_current + total_deferred

    # ── Effective tax rate ────────────────────────────────────────────
    if pre_tax_book_income != 0:
        effective_rate = (total_provision / pre_tax_book_income).quantize(
            Decimal("0.0001"), rounding=ROUND_HALF_UP
        )
    else:
        effective_rate = Decimal(0)

    # ── Rate reconciliation ──────────────────────────────────────────
    statutory = FEDERAL_RATE
    recon: list[tuple[str, Decimal, Decimal]] = []

    # Start with federal statutory rate
    statutory_tax = (pre_tax_book_income * statutory).quantize(
        Decimal("1"), rounding=ROUND_HALF_UP
    )
    recon.append(("Federal statutory rate (21%)", statutory_tax, statutory))

    # State taxes (net of federal benefit)
    state_net = (state_current * (Decimal(1) - FEDERAL_RATE)).quantize(
        Decimal("1"), rounding=ROUND_HALF_UP
    )
    state_rate_impact = (state_net / pre_tax_book_income).quantize(
        Decimal("0.0001"), rounding=ROUND_HALF_UP
    ) if pre_tax_book_income != 0 else Decimal(0)
    recon.append(("State taxes, net of federal benefit", state_net, state_rate_impact))

    # Permanent differences
    for perm in perms:
        impact = (perm.amount * combined_rate).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP
        )
        rate_impact = (impact / pre_tax_book_income).quantize(
            Decimal("0.0001"), rounding=ROUND_HALF_UP
        ) if pre_tax_book_income != 0 else Decimal(0)
        recon.append((perm.description, impact, rate_impact))

    # R&D credit
    rd_rate_impact = (-rd_credit / pre_tax_book_income).quantize(
        Decimal("0.0001"), rounding=ROUND_HALF_UP
    ) if pre_tax_book_income != 0 else Decimal(0)
    recon.append(("R&D tax credit (Section 41)", -rd_credit, rd_rate_impact))

    # Deferred tax provision (temporary differences)
    if total_deferred != 0:
        deferred_rate_impact = (total_deferred / pre_tax_book_income).quantize(
            Decimal("0.0001"), rounding=ROUND_HALF_UP
        ) if pre_tax_book_income != 0 else Decimal(0)
        recon.append(("Deferred tax provision (temporary differences)", total_deferred, deferred_rate_impact))

    # Rounding / other to tie out exactly
    recon_subtotal = sum(item[1] for item in recon)
    rounding_diff = total_provision - recon_subtotal
    if rounding_diff != 0:
        rounding_rate = (rounding_diff / pre_tax_book_income).quantize(
            Decimal("0.0001"), rounding=ROUND_HALF_UP
        ) if pre_tax_book_income != 0 else Decimal(0)
        recon.append(("Other / rounding", rounding_diff, rounding_rate))

    return TaxProvision(
        year=year,
        pre_tax_book_income=pre_tax_book_income,
        permanent_differences=perms,
        total_permanent=total_permanent,
        temporary_differences=temps,
        total_temporary_change=total_temporary_change,
        taxable_income=taxable_income,
        federal_current=federal_current,
        state_current=state_current,
        rd_credit=rd_credit,
        total_current=total_current,
        deferred_items=deferred_items,
        total_deferred=total_deferred,
        prior_dta_total=prior_dta,
        prior_dtl_total=prior_dtl,
        current_dta_total=current_dta,
        current_dtl_total=current_dtl,
        total_provision=total_provision,
        effective_tax_rate=effective_rate,
        rate_reconciliation=recon,
    )


# ── Multi-year convenience ───────────────────────────────────────────────────

def compute_provisions_multi_year(
    pre_tax_by_year: dict[int, Decimal],
    opex_records: list[MonthlyOpex],
    assets: list[FixedAsset],
    lease_schedules: list[LeaseScheduleRow],
    years: list[int] | None = None,
) -> dict[int, TaxProvision]:
    """Compute tax provisions for multiple years with DTA/DTL rollforward.

    Args:
        pre_tax_by_year: {year: pre_tax_book_income}.
        opex_records: All MonthlyOpex records.
        assets: All FixedAsset records.
        lease_schedules: All LeaseScheduleRow records.
        years: Years to compute (default: [2023, 2024, 2025]).

    Returns:
        {year: TaxProvision}.
    """
    if years is None:
        years = [2023, 2024, 2025]

    provisions: dict[int, TaxProvision] = {}
    prior: TaxProvision | None = None

    for year in sorted(years):
        pti = pre_tax_by_year.get(year, Decimal(0))
        prov = compute_tax_provision(
            pre_tax_book_income=pti,
            opex_records=opex_records,
            assets=assets,
            lease_schedules=lease_schedules,
            year=year,
            prior_provision=prior,
        )
        provisions[year] = prov
        prior = prov

    return provisions


# ── GL posting ───────────────────────────────────────────────────────────────

def post_tax_provision_to_gl(
    ledger: "Ledger",
    provisions: dict[int, TaxProvision],
) -> None:
    """Post tax provision entries to the GL for each year.

    Current provision:
      DR: Federal Income Tax Expense — Current (8010)
      DR: State Income Tax Expense — Current (8020)
      CR: R&D Tax Credit (8050)
      CR: Income Tax Payable — Federal (2070)
      CR: Income Tax Payable — State (2075)

    Deferred provision:
      DR/CR: Federal Income Tax Expense — Deferred (8030)
      DR/CR: State Income Tax Expense — Deferred (8040)
      DR/CR: Deferred Tax Asset / Liability accounts
    """
    import datetime

    from generator.model.gl import JournalEntry, JournalEntryLine

    for year in sorted(provisions.keys()):
        prov = provisions[year]
        date = datetime.date(year, 12, 31)

        # ── Current provision entry ──────────────────────────────────
        current_lines: list[JournalEntryLine] = []

        if prov.federal_current > 0:
            current_lines.append(JournalEntryLine(
                account="8010", debit=prov.federal_current, credit=Decimal(0),
                memo="Federal income tax expense — current",
            ))
            current_lines.append(JournalEntryLine(
                account="2070", debit=Decimal(0), credit=prov.federal_current,
                memo="Federal income tax payable",
            ))

        if prov.state_current > 0:
            current_lines.append(JournalEntryLine(
                account="8020", debit=prov.state_current, credit=Decimal(0),
                memo="State income tax expense — current",
            ))
            current_lines.append(JournalEntryLine(
                account="2075", debit=Decimal(0), credit=prov.state_current,
                memo="State income tax payable",
            ))

        if prov.rd_credit > 0:
            current_lines.append(JournalEntryLine(
                account="8050", debit=Decimal(0), credit=prov.rd_credit,
                memo="R&D tax credit (Section 41)",
            ))
            current_lines.append(JournalEntryLine(
                account="2070", debit=prov.rd_credit, credit=Decimal(0),
                memo="R&D credit offset against federal tax payable",
            ))

        if current_lines:
            entry = JournalEntry(
                date=date,
                entity_code="CI",  # Tax provision is at consolidated (parent) level
                description=f"Current income tax provision FY{year}",
                lines=tuple(current_lines),
            )
            ledger.post(entry)

        # ── Deferred provision entry ─────────────────────────────────
        # Net deferred = change in DTL - change in DTA
        if prov.total_deferred != 0:
            deferred_lines: list[JournalEntryLine] = []

            # Split deferred expense: ~77% federal, ~23% state (proportional to rates)
            fed_share = (FEDERAL_RATE / (FEDERAL_RATE + STATE_RATE)).quantize(
                Decimal("0.01"), rounding=ROUND_HALF_UP
            )
            fed_deferred = (prov.total_deferred * fed_share).quantize(
                Decimal("1"), rounding=ROUND_HALF_UP
            )
            state_deferred = prov.total_deferred - fed_deferred

            if fed_deferred > 0:
                deferred_lines.append(JournalEntryLine(
                    account="8030", debit=fed_deferred, credit=Decimal(0),
                    memo="Deferred federal income tax expense",
                ))
            elif fed_deferred < 0:
                deferred_lines.append(JournalEntryLine(
                    account="8030", debit=Decimal(0), credit=abs(fed_deferred),
                    memo="Deferred federal income tax benefit",
                ))

            if state_deferred > 0:
                deferred_lines.append(JournalEntryLine(
                    account="8040", debit=state_deferred, credit=Decimal(0),
                    memo="Deferred state income tax expense",
                ))
            elif state_deferred < 0:
                deferred_lines.append(JournalEntryLine(
                    account="8040", debit=Decimal(0), credit=abs(state_deferred),
                    memo="Deferred state income tax benefit",
                ))

            # Balance: offset to DTA/DTL accounts
            if prov.total_deferred > 0:
                # Increase in net DTL
                deferred_lines.append(JournalEntryLine(
                    account="2410", debit=Decimal(0), credit=prov.total_deferred,
                    memo="Increase in deferred tax liability",
                ))
            else:
                # Increase in DTA (net)
                deferred_lines.append(JournalEntryLine(
                    account="1910", debit=abs(prov.total_deferred), credit=Decimal(0),
                    memo="Increase in deferred tax asset",
                ))

            if deferred_lines:
                entry = JournalEntry(
                    date=date,
                    entity_code="CI",
                    description=f"Deferred income tax provision FY{year}",
                    lines=tuple(deferred_lines),
                )
                ledger.post(entry)


# ── Validation helpers ───────────────────────────────────────────────────────

def validate_provision(prov: TaxProvision) -> list[str]:
    """Validate internal consistency of a tax provision.

    Returns a list of error messages (empty = all good).
    """
    errors: list[str] = []

    # 1. Current + deferred = total provision
    expected_total = prov.total_current + prov.total_deferred
    if expected_total != prov.total_provision:
        errors.append(
            f"FY{prov.year}: current ({prov.total_current}) + deferred ({prov.total_deferred}) "
            f"!= total ({prov.total_provision})"
        )

    # 2. Rate reconciliation should approximate total provision
    recon_total = sum(item[1] for item in prov.rate_reconciliation)
    diff = abs(recon_total - prov.total_provision)
    # Allow ±$5K tolerance for rounding
    if diff > 5000:
        errors.append(
            f"FY{prov.year}: rate reconciliation total ({recon_total}) "
            f"differs from provision ({prov.total_provision}) by ${diff}"
        )

    # 3. ETR sanity check (should be between 15% and 35% for a normal C-corp)
    if prov.pre_tax_book_income > 0:
        if prov.effective_tax_rate < Decimal("0.15") or prov.effective_tax_rate > Decimal("0.35"):
            errors.append(
                f"FY{prov.year}: ETR of {prov.effective_tax_rate:.2%} is outside "
                f"expected range (15%–35%)"
            )

    return errors
