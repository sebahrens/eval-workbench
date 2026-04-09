"""Cascade Industries customers, contracts, key personnel, and litigation.

Provides queryable data for TC-11 (Quality of Earnings) and TC-12 (Data Room
Triage). Builds on top of the AR customer list (ar.py) for concentration
analysis and adds contract/legal/personnel data not captured elsewhere.

Key constraints from the spec:
- Top customer (Acme Manufacturing) = 18% of consolidated revenue.
- 8 customer contracts with terms, volumes, pricing, renewal dates.
- Acme contract has a change-of-control termination clause (TC-12 red flag).
- CEO employment agreement: golden parachute = 3× salary (TC-12 red flag).
- CFO and CTO employment agreements also present.
- Pending litigation: product liability suit, $2.5M potential exposure.

Feeds: TC-11 (QofE customer concentration, contract renewal risk),
       TC-12 (data room red flags, key employee agreements, litigation).
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal

from generator.model.ar import CUSTOMERS as AR_CUSTOMERS
from generator.model.entities import SUBSIDIARIES

# ── Customer contracts ──────────────────────────────────────────────────────
# 8 contracts covering the top customers.  Contract values are approximate
# annual volumes consistent with AR revenue shares.  Renewal dates are
# deterministic and spread across 2025–2028.

@dataclass(frozen=True)
class CustomerContract:
    """A customer contract with terms and renewal information."""

    contract_id: str
    customer_id: str
    customer_name: str
    entity_code: str
    effective_date: datetime.date
    expiration_date: datetime.date
    annual_volume: Decimal  # approximate annual contract value
    pricing_terms: str
    payment_terms: str
    auto_renew: bool
    change_of_control_clause: bool  # termination on ownership change
    notes: str = ""


CONTRACTS: tuple[CustomerContract, ...] = (
    # Acme Manufacturing — largest customer, has change-of-control clause
    CustomerContract(
        contract_id="CTR-001",
        customer_id="CUST-001",
        customer_name="Acme Manufacturing Corp",
        entity_code="PC",
        effective_date=datetime.date(2022, 1, 15),
        expiration_date=datetime.date(2025, 12, 31),
        annual_volume=Decimal("36_000_000"),
        pricing_terms="Fixed unit pricing with annual CPI adjustment cap 3%",
        payment_terms="Net 45",
        auto_renew=False,
        change_of_control_clause=True,
        notes=(
            "Master supply agreement. Change-of-control provision allows "
            "Acme to terminate within 90 days of ownership change without "
            "penalty. Represents ~18% of consolidated revenue."
        ),
    ),
    # Northwest Precision — second largest PC customer
    CustomerContract(
        contract_id="CTR-002",
        customer_id="CUST-002",
        customer_name="Northwest Precision Industries",
        entity_code="PC",
        effective_date=datetime.date(2023, 4, 1),
        expiration_date=datetime.date(2026, 3, 31),
        annual_volume=Decimal("14_250_000"),
        pricing_terms="Tiered volume pricing, quarterly true-up",
        payment_terms="Net 35",
        auto_renew=True,
        change_of_control_clause=False,
    ),
    # Columbia River Works
    CustomerContract(
        contract_id="CTR-003",
        customer_id="CUST-003",
        customer_name="Columbia River Works",
        entity_code="PC",
        effective_date=datetime.date(2021, 7, 1),
        expiration_date=datetime.date(2025, 6, 30),
        annual_volume=Decimal("11_400_000"),
        pricing_terms="Cost-plus-12% with quarterly raw material index adjustment",
        payment_terms="Net 40",
        auto_renew=True,
        change_of_control_clause=False,
        notes="Expiring within 12 months of FY2025 — renewal risk.",
    ),
    # TechAlloy Systems — largest AM customer
    CustomerContract(
        contract_id="CTR-004",
        customer_id="CUST-011",
        customer_name="TechAlloy Systems",
        entity_code="AM",
        effective_date=datetime.date(2023, 1, 1),
        expiration_date=datetime.date(2027, 12, 31),
        annual_volume=Decimal("14_300_000"),
        pricing_terms="Fixed pricing per kg, renegotiated annually",
        payment_terms="Net 40",
        auto_renew=True,
        change_of_control_clause=False,
    ),
    # Quantum Materials
    CustomerContract(
        contract_id="CTR-005",
        customer_id="CUST-012",
        customer_name="Quantum Materials Inc",
        entity_code="AM",
        effective_date=datetime.date(2022, 6, 1),
        expiration_date=datetime.date(2025, 5, 31),
        annual_volume=Decimal("11_700_000"),
        pricing_terms="Spot pricing with volume discount tiers",
        payment_terms="Net 35",
        auto_renew=False,
        change_of_control_clause=False,
        notes="Expiring within 12 months of FY2025 — renewal risk.",
    ),
    # NextGen Composites — government subcontract
    CustomerContract(
        contract_id="CTR-006",
        customer_id="CUST-013",
        customer_name="NextGen Composites LLC",
        entity_code="AM",
        effective_date=datetime.date(2024, 3, 1),
        expiration_date=datetime.date(2028, 2, 28),
        annual_volume=Decimal("9_750_000"),
        pricing_terms="Fixed-price per deliverable, milestone-based",
        payment_terms="Net 50",
        auto_renew=False,
        change_of_control_clause=False,
        notes="Government subcontract; subject to DFARS flow-down provisions.",
    ),
    # GlobalTrade Logistics — largest DS customer
    CustomerContract(
        contract_id="CTR-007",
        customer_id="CUST-019",
        customer_name="GlobalTrade Logistics",
        entity_code="DS",
        effective_date=datetime.date(2023, 9, 1),
        expiration_date=datetime.date(2026, 8, 31),
        annual_volume=Decimal("8_000_000"),
        pricing_terms="Per-pallet warehousing fee + per-mile freight rate",
        payment_terms="Net 35",
        auto_renew=True,
        change_of_control_clause=False,
    ),
    # Continental Supply Chain
    CustomerContract(
        contract_id="CTR-008",
        customer_id="CUST-020",
        customer_name="Continental Supply Chain",
        entity_code="DS",
        effective_date=datetime.date(2024, 1, 1),
        expiration_date=datetime.date(2025, 12, 31),
        annual_volume=Decimal("6_400_000"),
        pricing_terms="Flat monthly retainer + overage charges",
        payment_terms="Net 40",
        auto_renew=True,
        change_of_control_clause=False,
        notes="Expiring within 12 months of FY2025 — renewal risk.",
    ),
)


# ── Customer concentration analysis ────────────────────────────────────────


@dataclass
class CustomerConcentration:
    """Revenue concentration data for one customer."""

    customer_id: str
    customer_name: str
    entity_code: str
    annual_revenue: Decimal
    pct_of_entity: Decimal  # as decimal (e.g. 0.379)
    pct_of_consolidated: Decimal  # as decimal (e.g. 0.18)


def compute_customer_concentration(
    year: int = 2025,
) -> list[CustomerConcentration]:
    """Compute top-customer concentration for a given fiscal year.

    Uses AR customer revenue shares × entity FY targets. Returns all
    customers sorted by consolidated share descending.
    """
    consolidated_revenue = sum(
        Decimal(e.revenue_target) for e in SUBSIDIARIES.values()
    )

    results: list[CustomerConcentration] = []
    for cust in AR_CUSTOMERS:
        entity_target = Decimal(SUBSIDIARIES[cust.entity_code].revenue_target)
        annual_rev = (entity_target * cust.revenue_share).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP,
        )
        pct_entity = cust.revenue_share
        pct_consolidated = (annual_rev / consolidated_revenue).quantize(
            Decimal("0.0001"), rounding=ROUND_HALF_UP,
        )

        results.append(CustomerConcentration(
            customer_id=cust.id,
            customer_name=cust.name,
            entity_code=cust.entity_code,
            annual_revenue=annual_rev,
            pct_of_entity=pct_entity,
            pct_of_consolidated=pct_consolidated,
        ))

    results.sort(key=lambda c: (-c.pct_of_consolidated, c.customer_id))
    return results


def top_n_concentration(n: int = 10, year: int = 2025) -> Decimal:
    """Return the combined consolidated revenue share of the top N customers."""
    conc = compute_customer_concentration(year)
    return sum(c.pct_of_consolidated for c in conc[:n])


# ── Contract renewal risk ──────────────────────────────────────────────────


def contracts_expiring_within(
    months: int = 12,
    reference_date: datetime.date | None = None,
) -> list[CustomerContract]:
    """Return contracts expiring within N months of the reference date.

    Default reference: 2025-12-31 (FY2025 year-end).
    """
    if reference_date is None:
        reference_date = datetime.date(2025, 12, 31)

    cutoff = datetime.date(
        reference_date.year + (reference_date.month + months - 1) // 12,
        (reference_date.month + months - 1) % 12 + 1,
        reference_date.day,
    )

    return [
        c for c in CONTRACTS
        if c.expiration_date <= cutoff and c.expiration_date >= reference_date
    ]


def contracts_with_change_of_control() -> list[CustomerContract]:
    """Return contracts that have change-of-control termination clauses."""
    return [c for c in CONTRACTS if c.change_of_control_clause]


# ── Key personnel & employment agreements ──────────────────────────────────
# Deterministic key people for data room (TC-12).  These reference
# employees from the roster but carry additional contract metadata not
# stored in the Employee dataclass.

@dataclass(frozen=True)
class KeyPersonAgreement:
    """An employment agreement for a key executive."""

    person_name: str
    title: str
    entity_code: str
    base_salary: int
    bonus_target_pct: Decimal  # as decimal, e.g. 0.50 = 50%
    severance_months: int
    change_of_control_multiplier: Decimal  # 0 = none, 3 = 3× salary
    non_compete_months: int
    effective_date: datetime.date
    notes: str = ""


KEY_PERSONNEL: tuple[KeyPersonAgreement, ...] = (
    KeyPersonAgreement(
        person_name="Robert J. Cascade",
        title="Chief Executive Officer",
        entity_code="CI",
        base_salary=325_000,
        bonus_target_pct=Decimal("0.75"),
        severance_months=24,
        change_of_control_multiplier=Decimal("3"),
        non_compete_months=24,
        effective_date=datetime.date(2019, 3, 1),
        notes=(
            "Golden parachute: 3× base salary ($975,000) upon "
            "change of control. Board-approved retention provision."
        ),
    ),
    KeyPersonAgreement(
        person_name="Margaret L. Chen",
        title="Chief Financial Officer",
        entity_code="CI",
        base_salary=260_000,
        bonus_target_pct=Decimal("0.50"),
        severance_months=18,
        change_of_control_multiplier=Decimal("2"),
        non_compete_months=18,
        effective_date=datetime.date(2020, 8, 15),
        notes="Standard executive employment agreement with 2× CoC provision.",
    ),
    KeyPersonAgreement(
        person_name="David R. Nakamura",
        title="Chief Technology Officer",
        entity_code="CI",
        base_salary=280_000,
        bonus_target_pct=Decimal("0.50"),
        severance_months=18,
        change_of_control_multiplier=Decimal("2"),
        non_compete_months=12,
        effective_date=datetime.date(2021, 1, 10),
        notes=(
            "Also serves as head of Advanced Materials R&D. "
            "IP assignment clause references founding-era work product."
        ),
    ),
)


def key_personnel_with_coc() -> list[KeyPersonAgreement]:
    """Return key personnel with non-zero change-of-control provisions."""
    return [
        kp for kp in KEY_PERSONNEL
        if kp.change_of_control_multiplier > 0
    ]


def total_coc_exposure() -> Decimal:
    """Total change-of-control payment exposure across all key personnel."""
    return sum(
        Decimal(kp.base_salary) * kp.change_of_control_multiplier
        for kp in KEY_PERSONNEL
    )


# ── Pending litigation ─────────────────────────────────────────────────────

@dataclass(frozen=True)
class LitigationMatter:
    """A pending litigation matter for data room disclosure."""

    matter_id: str
    title: str
    case_type: str
    filing_date: datetime.date
    court: str
    plaintiff: str
    defendant: str
    description: str
    potential_exposure: Decimal
    accrued_liability: Decimal  # amount accrued on balance sheet
    status: str
    outside_counsel: str


LITIGATION: tuple[LitigationMatter, ...] = (
    LitigationMatter(
        matter_id="LIT-001",
        title="Henderson v. Cascade Precision Components LLC",
        case_type="Product Liability",
        filing_date=datetime.date(2024, 6, 15),
        court="Multnomah County Circuit Court, Oregon",
        plaintiff="Henderson Industrial Services, Inc.",
        defendant="Cascade Precision Components LLC",
        description=(
            "Plaintiff alleges defective industrial bearing assembly (Part "
            "#PC-4872) caused equipment failure and production downtime. "
            "Claims include property damage ($1.8M), lost profits ($0.5M), "
            "and consequential damages ($0.2M). Cascade disputes defect "
            "allegations and contends plaintiff failed to follow "
            "maintenance specifications."
        ),
        potential_exposure=Decimal("2_500_000"),
        accrued_liability=Decimal("750_000"),
        status="Discovery phase — depositions scheduled Q1 2026",
        outside_counsel="Mitchell, Hartwell & Associates LLP",
    ),
)


def total_litigation_exposure() -> Decimal:
    """Sum of potential exposure across all pending matters."""
    return sum(m.potential_exposure for m in LITIGATION)


# ── Combined red flags for TC-12 ───────────────────────────────────────────


@dataclass
class RedFlag:
    """A red flag item for data room triage."""

    category: str  # "contract", "personnel", "litigation", "ip", "missing"
    severity: str  # "high", "medium", "low"
    title: str
    detail: str


def data_room_red_flags() -> list[RedFlag]:
    """Compile all known red flags for TC-12 data room triage.

    Returns items sorted by severity (high first), then category.
    """
    flags: list[RedFlag] = []

    # Litigation
    for matter in LITIGATION:
        flags.append(RedFlag(
            category="litigation",
            severity="high",
            title=f"Pending litigation: {matter.title}",
            detail=(
                f"{matter.case_type} — potential exposure "
                f"${matter.potential_exposure:,.0f}. "
                f"Status: {matter.status}"
            ),
        ))

    # Change-of-control contracts
    for contract in contracts_with_change_of_control():
        flags.append(RedFlag(
            category="contract",
            severity="high",
            title=(
                f"Change-of-control clause: {contract.customer_name} "
                f"({contract.contract_id})"
            ),
            detail=(
                f"Customer may terminate within 90 days of ownership change. "
                f"Annual volume ~${contract.annual_volume:,.0f} "
                f"(~18% of consolidated revenue)."
            ),
        ))

    # CEO golden parachute
    for kp in KEY_PERSONNEL:
        if kp.change_of_control_multiplier >= 3:
            payout = Decimal(kp.base_salary) * kp.change_of_control_multiplier
            flags.append(RedFlag(
                category="personnel",
                severity="high",
                title=f"{kp.title} golden parachute: {kp.change_of_control_multiplier}× salary",
                detail=(
                    f"{kp.person_name} — payout ${payout:,.0f} upon "
                    f"change of control. {kp.notes}"
                ),
            ))

    # Contracts expiring within 12 months
    expiring = contracts_expiring_within(months=12)
    for contract in expiring:
        flags.append(RedFlag(
            category="contract",
            severity="medium",
            title=(
                f"Contract expiring: {contract.customer_name} "
                f"({contract.expiration_date.isoformat()})"
            ),
            detail=(
                f"Annual volume ~${contract.annual_volume:,.0f}. "
                f"Auto-renew: {'yes' if contract.auto_renew else 'no'}."
            ),
        ))

    # Sort: high → medium → low, then by category
    severity_order = {"high": 0, "medium": 1, "low": 2}
    flags.sort(key=lambda f: (severity_order.get(f.severity, 9), f.category, f.title))
    return flags
