"""European IAS 12 tax provision model for TC-06-EU.

Computes a multi-jurisdiction income tax provision for the Cascade Europe
Holdings B.V. group (CE/CP/CM/CD) under IAS 12 Income Taxes.

Jurisdictions:
- NL (25.8%): Cascade Europe Holdings B.V. — holding company
- DE (29.9%): Cascade Präzisionsteile GmbH — manufacturing
- FR (25.0%): Cascade Matériaux Avancés SAS — R&D / materials
- UK (25.0%): Cascade Distribution Services Ltd — distribution (GBP)

Key differences from ASC 740 (TC-06):
- No valuation allowance — IAS 12 uses "probable" recoverability assessment
- No indefinite reversal exception for subsidiary earnings
- Multi-jurisdiction current tax instead of federal+state
- Weighted group statutory rate for ETR reconciliation
- CIR as tax credit (not book income adjustment)
- GBP→EUR translation (IAS 21)
- Pillar Two not applicable (revenue < €750M)

Determinism: all inputs are fixed constants per the design bead (synth-data-eu.6).
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal

# ── Tax rates by jurisdiction ───────────────────────────────────────────────

# NL: 25.8% (profits > €200k)
NL_RATE = Decimal("0.258")

# DE composite: KSt 15% + SolZ 5.5% on KSt + GewSt at Munich 490% Hebesatz
# KSt = 15%, SolZ = 15% × 5.5% = 0.825%, GewSt = 3.5% × 490% / 100 = 17.15%
# But GewSt deductible from its own base: effective = 3.5% × 490/100 / (1 + 3.5% × 490/100)
# Simpler: standard composite = 15.825% + 14.075% = 29.9% (widely used benchmark)
DE_KST_RATE = Decimal("0.15")
DE_SOLZ_RATE = Decimal("0.055")  # 5.5% of KSt
DE_GEWST_HEBESATZ = Decimal("490")  # Munich current rate
DE_GEWST_MESSZAHL = Decimal("0.035")  # 3.5% base rate
DE_RATE = Decimal("0.299")  # Composite

# Stale Hebesatz for planted error (ERR-EU-002)
DE_GEWST_HEBESATZ_STALE = Decimal("480")  # old Munich rate
DE_RATE_STALE = Decimal("0.2958")  # 15.825% + ~13.755% (approx with old Hebesatz)

# FR: 25% (IS rate 2025)
FR_RATE = Decimal("0.25")

# UK: 25% (main rate for profits > £250k)
UK_RATE = Decimal("0.25")

# GBP/EUR average rate for IAS 21 translation
GBP_EUR_AVG = Decimal("1.17")

# ── Entity financial data (deterministic, from EU company profile) ──────────

# Revenue targets from profile
_REVENUE = {
    "CE": Decimal("22000000"),   # Holding — management fees + loan interest
    "CP": Decimal("45000000"),   # Manufacturing
    "CM": Decimal("32000000"),   # R&D / materials
    "CD": Decimal("21000000"),   # Distribution (translated from £18M at 1.17)
}

# Gross margins
_GROSS_MARGIN = {
    "CE": Decimal("0.85"),  # Holding company — mostly fee income, low cost base
    "CP": Decimal("0.35"),
    "CM": Decimal("0.52"),
    "CD": Decimal("0.18"),
}

# Operating expense ratio (of revenue)
_OPEX_RATIO = {
    "CE": Decimal("0.12"),  # Lean holding
    "CP": Decimal("0.22"),
    "CM": Decimal("0.28"),
    "CD": Decimal("0.14"),
}


@dataclass(frozen=True)
class EntityPermanentDifference:
    """A permanent book-tax difference for a specific entity."""
    entity_code: str
    description: str
    amount_eur: Decimal  # Positive = increases taxable; negative = decreases


@dataclass(frozen=True)
class EntityTemporaryDifference:
    """A temporary book-tax difference for a specific entity."""
    entity_code: str
    description: str
    book_amount_eur: Decimal
    tax_amount_eur: Decimal

    @property
    def difference(self) -> Decimal:
        return self.tax_amount_eur - self.book_amount_eur


@dataclass(frozen=True)
class EntityDeferredTaxItem:
    """Deferred tax asset/liability for one entity's temporary difference."""
    entity_code: str
    description: str
    opening_eur: Decimal
    movement_eur: Decimal
    closing_eur: Decimal
    item_type: str  # "DTA" or "DTL"


@dataclass(frozen=True)
class EntityCurrentTax:
    """Current tax computation for one entity."""
    entity_code: str
    jurisdiction: str
    pre_tax_book_income_eur: Decimal
    permanent_adjustments_eur: Decimal
    temporary_adjustments_eur: Decimal
    taxable_income_eur: Decimal
    statutory_rate: Decimal
    gross_tax_eur: Decimal
    tax_credits_eur: Decimal  # CIR, Forschungszulage, RDEC
    current_tax_eur: Decimal  # gross_tax - credits
    # For UK: original GBP amounts before translation
    current_tax_local: Decimal | None = None
    local_currency: str = "EUR"


@dataclass(frozen=True)
class RateReconItem:
    """One line of the consolidated ETR reconciliation."""
    description: str
    amount_eur: Decimal
    rate_impact: Decimal


@dataclass
class EuropeanTaxProvision:
    """Complete IAS 12 tax provision for the Cascade Europe group."""
    year: int

    # Per-entity data
    entity_pre_tax: dict[str, Decimal]  # entity_code -> pre-tax book income EUR
    consolidated_pre_tax: Decimal

    # Permanent differences
    permanent_differences: list[EntityPermanentDifference]

    # Temporary differences
    temporary_differences: list[EntityTemporaryDifference]

    # Current tax by entity
    current_tax_by_entity: list[EntityCurrentTax]
    total_current_tax: Decimal

    # Deferred tax rollforward
    deferred_items: list[EntityDeferredTaxItem]
    total_deferred_opening: Decimal
    total_deferred_movement: Decimal
    total_deferred_closing: Decimal

    # ETR reconciliation
    weighted_statutory_rate: Decimal
    effective_tax_rate: Decimal
    rate_reconciliation: list[RateReconItem]

    # Total provision
    total_provision: Decimal

    # Pillar Two
    pillar_two_applicable: bool


def _rd(d: Decimal) -> Decimal:
    """Round to whole EUR."""
    return d.quantize(Decimal("1"), rounding=ROUND_HALF_UP)


def compute_eu_pre_tax_income(year: int) -> dict[str, Decimal]:
    """Compute deterministic pre-tax book income per entity.

    Uses revenue, gross margin, and opex ratio from the EU company profile.
    Growth: FY2024→FY2025 = +8%.
    """
    growth = Decimal("1.08") if year == 2025 else Decimal("1.00")

    result = {}
    for entity in ("CE", "CP", "CM", "CD"):
        rev = _rd(_REVENUE[entity] * growth)
        gross_profit = _rd(rev * _GROSS_MARGIN[entity])
        opex = _rd(rev * _OPEX_RATIO[entity])
        result[entity] = gross_profit - opex
    return result


def compute_eu_permanent_differences(
    year: int,
    entity_income: dict[str, Decimal],
) -> list[EntityPermanentDifference]:
    """Compute permanent differences per the design bead."""
    perms: list[EntityPermanentDifference] = []

    # CE: Non-deductible supervisory board fees
    perms.append(EntityPermanentDifference(
        entity_code="CE",
        description="Non-deductible supervisory board fees",
        amount_eur=Decimal("85000"),
    ))

    # CE: Participation exemption dividends (exempt)
    perms.append(EntityPermanentDifference(
        entity_code="CE",
        description="Participation exemption dividends",
        amount_eur=Decimal("-120000"),
    ))

    # CP: Non-deductible entertainment (50% Bewirtungskosten)
    perms.append(EntityPermanentDifference(
        entity_code="CP",
        description="Non-deductible entertainment (50% Bewirtungskosten)",
        amount_eur=Decimal("45000"),
    ))

    # CM: Non-deductible fines (DGCCRF penalty)
    perms.append(EntityPermanentDifference(
        entity_code="CM",
        description="Non-deductible fines (DGCCRF penalty)",
        amount_eur=Decimal("30000"),
    ))

    # CD: Non-deductible UK business entertainment (100% per HMRC)
    # £22k at 1.17 = €25,740
    perms.append(EntityPermanentDifference(
        entity_code="CD",
        description="Non-deductible UK business entertainment (100% HMRC)",
        amount_eur=_rd(Decimal("22000") * GBP_EUR_AVG),
    ))

    return perms


def compute_eu_temporary_differences(
    year: int,
) -> list[EntityTemporaryDifference]:
    """Compute temporary differences per the design bead."""
    temps: list[EntityTemporaryDifference] = []

    # CP: Accelerated tax depreciation (degressive vs straight-line)
    temps.append(EntityTemporaryDifference(
        entity_code="CP",
        description="Accelerated tax depreciation (degressive Abschreibung)",
        book_amount_eur=Decimal("2800000"),
        tax_amount_eur=Decimal("3950000"),
    ))

    # CP: Warranty provision (Rückstellung, tax timing)
    temps.append(EntityTemporaryDifference(
        entity_code="CP",
        description="Warranty provision (Rückstellung)",
        book_amount_eur=Decimal("680000"),
        tax_amount_eur=Decimal("510000"),
    ))

    # CP: Inventory write-down reserve
    temps.append(EntityTemporaryDifference(
        entity_code="CP",
        description="Inventory write-down reserve",
        book_amount_eur=Decimal("420000"),
        tax_amount_eur=Decimal("0"),
    ))

    # CM: Tax depreciation (amortissement dégressif vs straight-line)
    temps.append(EntityTemporaryDifference(
        entity_code="CM",
        description="Tax depreciation (amortissement dégressif)",
        book_amount_eur=Decimal("1600000"),
        tax_amount_eur=Decimal("2200000"),
    ))

    # CM: Accrued bonus provision (timing)
    temps.append(EntityTemporaryDifference(
        entity_code="CM",
        description="Accrued bonus provision",
        book_amount_eur=Decimal("850000"),
        tax_amount_eur=Decimal("0"),
    ))

    # CM: Bad debt provision
    temps.append(EntityTemporaryDifference(
        entity_code="CM",
        description="Bad debt provision",
        book_amount_eur=Decimal("320000"),
        tax_amount_eur=Decimal("180000"),
    ))

    # CD: Capital allowances vs book depreciation
    temps.append(EntityTemporaryDifference(
        entity_code="CD",
        description="Capital allowances vs book depreciation",
        book_amount_eur=_rd(Decimal("450000") * GBP_EUR_AVG),
        tax_amount_eur=_rd(Decimal("620000") * GBP_EUR_AVG),
    ))

    # CD: Holiday pay accrual (deductible when paid)
    temps.append(EntityTemporaryDifference(
        entity_code="CD",
        description="Holiday pay accrual",
        book_amount_eur=_rd(Decimal("185000") * GBP_EUR_AVG),
        tax_amount_eur=Decimal("0"),
    ))

    # CD: Stock provisions
    temps.append(EntityTemporaryDifference(
        entity_code="CD",
        description="Stock provisions",
        book_amount_eur=_rd(Decimal("95000") * GBP_EUR_AVG),
        tax_amount_eur=Decimal("0"),
    ))

    # CE: Intercompany loan impairment provision
    temps.append(EntityTemporaryDifference(
        entity_code="CE",
        description="Intercompany loan impairment provision",
        book_amount_eur=Decimal("250000"),
        tax_amount_eur=Decimal("0"),
    ))

    return temps


def _entity_rate(entity_code: str) -> Decimal:
    """Statutory rate for an entity's jurisdiction."""
    return {"CE": NL_RATE, "CP": DE_RATE, "CM": FR_RATE, "CD": UK_RATE}[entity_code]


def _entity_jurisdiction(entity_code: str) -> str:
    return {"CE": "NL", "CP": "DE", "CM": "FR", "CD": "UK"}[entity_code]


def compute_eu_tax_provision(year: int = 2025) -> EuropeanTaxProvision:
    """Compute the full IAS 12 provision for the Cascade Europe group."""

    # ── Pre-tax book income ─────────────────────────────────────────
    entity_income = compute_eu_pre_tax_income(year)
    consolidated_pre_tax = sum(entity_income.values())

    # ── Permanent differences ───────────────────────────────────────
    perms = compute_eu_permanent_differences(year, entity_income)
    perms_by_entity: dict[str, Decimal] = {}
    for p in perms:
        perms_by_entity[p.entity_code] = (
            perms_by_entity.get(p.entity_code, Decimal(0)) + p.amount_eur
        )

    # ── Temporary differences ───────────────────────────────────────
    temps = compute_eu_temporary_differences(year)
    temps_by_entity: dict[str, Decimal] = {}
    for t in temps:
        temps_by_entity[t.entity_code] = (
            temps_by_entity.get(t.entity_code, Decimal(0)) + t.difference
        )

    # ── R&D tax credits ─────────────────────────────────────────────
    # France CIR: 30% of eligible R&D spend. CM revenue €32M × 1.08 × 12% R&D
    cm_revenue = _rd(_REVENUE["CM"] * Decimal("1.08")) if year == 2025 else _REVENUE["CM"]
    cm_rd_spend = _rd(cm_revenue * Decimal("0.12"))
    cir_credit = _rd(cm_rd_spend * Decimal("0.30"))

    # Germany Forschungszulage: 25% of personnel costs, max €500K benefit
    # CP R&D = 4% of revenue; personnel ~60% of R&D spend
    cp_revenue = _rd(_REVENUE["CP"] * Decimal("1.08")) if year == 2025 else _REVENUE["CP"]
    cp_rd_spend = _rd(cp_revenue * Decimal("0.04"))
    cp_rd_personnel = _rd(cp_rd_spend * Decimal("0.60"))
    forschungszulage = min(_rd(cp_rd_personnel * Decimal("0.25")), Decimal("500000"))

    # UK RDEC: 20% of eligible R&D (we model as above-the-line credit,
    # so it's already in book income — no separate tax credit deduction needed here).
    # For simplicity, UK RDEC is treated as already reflected in pre-tax income.
    uk_rdec = Decimal("0")

    credits_by_entity = {
        "CE": Decimal("0"),
        "CP": forschungszulage,
        "CM": cir_credit,
        "CD": uk_rdec,
    }

    # ── Current tax by entity ───────────────────────────────────────
    current_taxes: list[EntityCurrentTax] = []
    total_current = Decimal(0)

    for entity in ("CE", "CP", "CM", "CD"):
        pti = entity_income[entity]
        perm_adj = perms_by_entity.get(entity, Decimal(0))
        temp_adj = temps_by_entity.get(entity, Decimal(0))
        taxable = pti + perm_adj + temp_adj
        rate = _entity_rate(entity)
        gross_tax = _rd(taxable * rate)
        credits = credits_by_entity[entity]
        net_tax = gross_tax - credits

        local_tax = None
        local_ccy = "EUR"
        if entity == "CD":
            # Compute in GBP, then translate
            local_ccy = "GBP"
            local_tax = _rd(net_tax / GBP_EUR_AVG)
            # Re-translate to EUR at average rate
            net_tax = _rd(local_tax * GBP_EUR_AVG)

        current_taxes.append(EntityCurrentTax(
            entity_code=entity,
            jurisdiction=_entity_jurisdiction(entity),
            pre_tax_book_income_eur=pti,
            permanent_adjustments_eur=perm_adj,
            temporary_adjustments_eur=temp_adj,
            taxable_income_eur=taxable,
            statutory_rate=rate,
            gross_tax_eur=gross_tax,
            tax_credits_eur=credits,
            current_tax_eur=net_tax,
            current_tax_local=local_tax,
            local_currency=local_ccy,
        ))
        total_current += net_tax

    # ── Deferred tax rollforward ────────────────────────────────────
    # FY2024 opening balances (deterministic, per design)
    # We compute FY2024 movements and derive opening/closing
    _fy24_opening = {
        ("CP", "Accelerated tax depreciation (degressive Abschreibung)"): Decimal("285000"),
        ("CP", "Warranty provision (Rückstellung)"): Decimal("-42000"),
        ("CP", "Inventory write-down reserve"): Decimal("-95000"),
        ("CM", "Tax depreciation (amortissement dégressif)"): Decimal("148000"),
        ("CM", "Accrued bonus provision"): Decimal("-195000"),
        ("CM", "Bad debt provision"): Decimal("-32000"),
        ("CD", "Capital allowances vs book depreciation"): Decimal("38000"),
        ("CD", "Holiday pay accrual"): Decimal("-48000"),
        ("CD", "Stock provisions"): Decimal("-22000"),
        ("CE", "Intercompany loan impairment provision"): Decimal("-58000"),
    }

    deferred_items: list[EntityDeferredTaxItem] = []
    total_opening = Decimal(0)
    total_closing = Decimal(0)

    for temp in temps:
        key = (temp.entity_code, temp.description)
        opening = _fy24_opening.get(key, Decimal(0))
        rate = _entity_rate(temp.entity_code)
        movement = _rd(temp.difference * rate)
        closing = opening + movement
        item_type = "DTL" if closing > 0 else "DTA"

        deferred_items.append(EntityDeferredTaxItem(
            entity_code=temp.entity_code,
            description=temp.description,
            opening_eur=opening,
            movement_eur=movement,
            closing_eur=closing,
            item_type=item_type,
        ))
        total_opening += opening
        total_closing += closing

    total_movement = total_closing - total_opening

    # ── Total provision ─────────────────────────────────────────────
    total_provision = total_current + total_movement

    # ── Weighted statutory rate ──────────────────────────────────────
    weighted_rate = Decimal(0)
    for entity in ("CE", "CP", "CM", "CD"):
        weight = entity_income[entity] / consolidated_pre_tax if consolidated_pre_tax else Decimal(0)
        weighted_rate += weight * _entity_rate(entity)
    weighted_rate = weighted_rate.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)

    # ── ETR ─────────────────────────────────────────────────────────
    effective_rate = Decimal(0)
    if consolidated_pre_tax:
        effective_rate = (total_provision / consolidated_pre_tax).quantize(
            Decimal("0.0001"), rounding=ROUND_HALF_UP
        )

    # ── Rate reconciliation ─────────────────────────────────────────
    recon: list[RateReconItem] = []

    # Weighted statutory rate
    weighted_tax = _rd(consolidated_pre_tax * weighted_rate)
    recon.append(RateReconItem(
        description="Weighted group statutory rate",
        amount_eur=weighted_tax,
        rate_impact=weighted_rate,
    ))

    # Rate differentials (difference from weighted rate per entity)
    for entity in ("CE", "CP", "CM", "CD"):
        diff = _entity_rate(entity) - weighted_rate
        if diff != 0:
            impact = _rd(entity_income[entity] * diff)
            rate_impact = (impact / consolidated_pre_tax).quantize(
                Decimal("0.0001"), rounding=ROUND_HALF_UP
            ) if consolidated_pre_tax else Decimal(0)
            jur = _entity_jurisdiction(entity)
            recon.append(RateReconItem(
                description=f"Rate differential — {jur} ({entity})",
                amount_eur=impact,
                rate_impact=rate_impact,
            ))

    # Permanent differences (combined impact)
    total_perm_tax = Decimal(0)
    for p in perms:
        rate = _entity_rate(p.entity_code)
        total_perm_tax += _rd(p.amount_eur * rate)
    if total_perm_tax:
        recon.append(RateReconItem(
            description="Permanent differences",
            amount_eur=total_perm_tax,
            rate_impact=(total_perm_tax / consolidated_pre_tax).quantize(
                Decimal("0.0001"), rounding=ROUND_HALF_UP
            ) if consolidated_pre_tax else Decimal(0),
        ))

    # R&D credits
    total_credits = cir_credit + forschungszulage
    if total_credits:
        recon.append(RateReconItem(
            description="R&D tax credits (CIR + Forschungszulage)",
            amount_eur=-total_credits,
            rate_impact=(-total_credits / consolidated_pre_tax).quantize(
                Decimal("0.0001"), rounding=ROUND_HALF_UP
            ) if consolidated_pre_tax else Decimal(0),
        ))

    # Deferred tax movement
    if total_movement:
        recon.append(RateReconItem(
            description="Deferred tax movement",
            amount_eur=total_movement,
            rate_impact=(total_movement / consolidated_pre_tax).quantize(
                Decimal("0.0001"), rounding=ROUND_HALF_UP
            ) if consolidated_pre_tax else Decimal(0),
        ))

    # FX effects (UK translation rounding)
    recon_subtotal = sum(r.amount_eur for r in recon)
    fx_diff = total_provision - recon_subtotal
    if fx_diff:
        recon.append(RateReconItem(
            description="FX translation and rounding",
            amount_eur=fx_diff,
            rate_impact=(fx_diff / consolidated_pre_tax).quantize(
                Decimal("0.0001"), rounding=ROUND_HALF_UP
            ) if consolidated_pre_tax else Decimal(0),
        ))

    return EuropeanTaxProvision(
        year=year,
        entity_pre_tax=entity_income,
        consolidated_pre_tax=consolidated_pre_tax,
        permanent_differences=perms,
        temporary_differences=temps,
        current_tax_by_entity=current_taxes,
        total_current_tax=total_current,
        deferred_items=deferred_items,
        total_deferred_opening=total_opening,
        total_deferred_movement=total_movement,
        total_deferred_closing=total_closing,
        weighted_statutory_rate=weighted_rate,
        effective_tax_rate=effective_rate,
        rate_reconciliation=recon,
        total_provision=total_provision,
        pillar_two_applicable=False,
    )


def compute_eu_tax_provision_fy2024() -> EuropeanTaxProvision:
    """FY2024 provision (prior year) for the workpaper.

    Uses the same structure but with FY2024 numbers (no growth).
    """
    return compute_eu_tax_provision(year=2024)
