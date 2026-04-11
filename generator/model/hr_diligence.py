"""Cascade Industries HR diligence canonical model.

Provides deterministic employment agreements, retention awards, severance
exposures, contractor classification signals, and diligence requests for
M&A due-diligence test cases (TC-20, TC-21).

Cross-references existing model objects:
- Entity codes (CI, PC, AM, DS) from entities.py
- Key personnel names from customers.py (KeyPersonAgreement)
- Employee roster from employees.py

All data is hardcoded tuples for determinism.  ID conventions follow the
design decision in bead synth-data-ups.6:
  EA-NNN, RET-NNN, SEV-NNN, CCS-NNN, DR-NNN
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass
from decimal import Decimal

# ── Dataclasses ──────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class EmploymentAgreement:
    """An employment agreement for diligence purposes.

    Superset of KeyPersonAgreement (customers.py): adds agreement_id,
    ip_assignment, and executed flag.  Both coexist; this serves TC-19/20/21.
    """

    agreement_id: str  # EA-NNN
    employee_name: str
    employee_title: str
    entity_code: str  # CI, PC, AM, DS
    base_salary: int
    bonus_target_pct: Decimal
    severance_multiplier: Decimal  # e.g. 1.5 = 1.5x base salary
    severance_months: int  # notice period in months
    change_of_control_multiplier: Decimal
    non_compete_months: int
    ip_assignment: bool
    effective_date: datetime.date
    executed: bool  # True = signed copy exists; False = draft/missing
    notes: str = ""


@dataclass(frozen=True)
class RetentionAward:
    """A retention award tied to a specific employee."""

    award_id: str  # RET-NNN
    employee_name: str
    employee_title: str
    entity_code: str
    award_amount: Decimal
    vesting_date: datetime.date
    retention_period_months: int
    forfeiture_conditions: str
    notes: str = ""


@dataclass(frozen=True)
class SeveranceExposure:
    """Computed severance exposure for an employee."""

    exposure_id: str  # SEV-NNN
    employee_name: str
    employee_title: str
    entity_code: str
    base_salary: int
    severance_multiplier: Decimal
    estimated_payout: Decimal  # base_salary x multiplier
    trigger: str  # termination_without_cause, change_of_control, mutual_separation
    accrued: bool  # whether already accrued on balance sheet
    notes: str = ""


@dataclass(frozen=True)
class ContractorClassificationSignal:
    """A contractor classification risk signal.

    Notes field is a follow-up recommendation, NOT a definitive legal
    conclusion.  This is a judgment trap: agents must present these as
    signals requiring further investigation, not determinations.
    """

    signal_id: str  # CCS-NNN
    contractor_name: str
    entity_code: str
    role_description: str
    tenure_months: int
    exclusive_engagement: bool
    uses_company_equipment: bool
    has_set_hours: bool
    risk_level: str  # high, medium, low
    notes: str = ""


@dataclass(frozen=True)
class DiligenceRequest:
    """A diligence information request."""

    request_id: str  # DR-NNN
    category: str  # legal, hr, financial, operational
    description: str
    requested_from: str  # management, legal, hr
    status: str  # open, received, partial, not_received
    due_date: datetime.date
    received_date: datetime.date | None
    source_refs: tuple[str, ...]  # document IDs received in response
    notes: str = ""


# ── Canonical Data ───────────────────────────────────────────────────────────
# All values are deterministic and cross-reference existing Cascade entities.
# Employee names for CEO/CFO/CTO match KeyPersonAgreement in customers.py.


EMPLOYMENT_AGREEMENTS: tuple[EmploymentAgreement, ...] = (
    # CEO — matches Robert J. Cascade from customers.KEY_PERSONNEL
    EmploymentAgreement(
        agreement_id="EA-001",
        employee_name="Robert J. Cascade",
        employee_title="Chief Executive Officer",
        entity_code="CI",
        base_salary=325_000,
        bonus_target_pct=Decimal("0.75"),
        severance_multiplier=Decimal("3"),
        severance_months=24,
        change_of_control_multiplier=Decimal("3"),
        non_compete_months=24,
        ip_assignment=True,
        effective_date=datetime.date(2019, 3, 1),
        executed=True,
        notes=(
            "Golden parachute: 3x base salary ($975,000) upon change of "
            "control. Board-approved retention provision. Signed copy on file."
        ),
    ),
    # CFO — matches Margaret L. Chen from customers.KEY_PERSONNEL
    EmploymentAgreement(
        agreement_id="EA-002",
        employee_name="Margaret L. Chen",
        employee_title="Chief Financial Officer",
        entity_code="CI",
        base_salary=260_000,
        bonus_target_pct=Decimal("0.50"),
        severance_multiplier=Decimal("2"),
        severance_months=18,
        change_of_control_multiplier=Decimal("2"),
        non_compete_months=18,
        ip_assignment=True,
        effective_date=datetime.date(2020, 8, 15),
        executed=True,
        notes="Standard executive employment agreement with 2x CoC provision.",
    ),
    # CTO — matches David R. Nakamura from customers.KEY_PERSONNEL
    EmploymentAgreement(
        agreement_id="EA-003",
        employee_name="David R. Nakamura",
        employee_title="Chief Technology Officer",
        entity_code="CI",
        base_salary=280_000,
        bonus_target_pct=Decimal("0.50"),
        severance_multiplier=Decimal("2"),
        severance_months=18,
        change_of_control_multiplier=Decimal("2"),
        non_compete_months=12,
        ip_assignment=True,
        effective_date=datetime.date(2021, 1, 10),
        executed=True,
        notes=(
            "Also serves as head of Advanced Materials R&D. "
            "IP assignment clause references founding-era work product."
        ),
    ),
    # VP Operations — executed agreement
    EmploymentAgreement(
        agreement_id="EA-004",
        employee_name="Susan M. Torres",
        employee_title="VP Operations",
        entity_code="PC",
        base_salary=155_000,
        bonus_target_pct=Decimal("0.30"),
        severance_multiplier=Decimal("1.5"),
        severance_months=12,
        change_of_control_multiplier=Decimal("1.5"),
        non_compete_months=12,
        ip_assignment=False,
        effective_date=datetime.date(2022, 5, 1),
        executed=True,
        notes="Operations leadership agreement. Signed copy on file.",
    ),
    # VP Sales — executed agreement
    EmploymentAgreement(
        agreement_id="EA-005",
        employee_name="James K. Whitfield",
        employee_title="VP Sales",
        entity_code="PC",
        base_salary=160_000,
        bonus_target_pct=Decimal("0.40"),
        severance_multiplier=Decimal("1"),
        severance_months=12,
        change_of_control_multiplier=Decimal("1"),
        non_compete_months=12,
        ip_assignment=False,
        effective_date=datetime.date(2023, 2, 15),
        executed=True,
        notes="Sales leadership agreement with commission override provisions.",
    ),
    # R&D Director — MISSING executed copy (judgment trap)
    EmploymentAgreement(
        agreement_id="EA-006",
        employee_name="Dr. Anika Patel",
        employee_title="R&D Director",
        entity_code="AM",
        base_salary=170_000,
        bonus_target_pct=Decimal("0.30"),
        severance_multiplier=Decimal("1.5"),
        severance_months=12,
        change_of_control_multiplier=Decimal("1.5"),
        non_compete_months=18,
        ip_assignment=True,
        effective_date=datetime.date(2021, 9, 1),
        executed=False,
        notes=(
            "Draft agreement only — no executed copy on file. "
            "Critical for IP assignment coverage on AM R&D output. "
            "Management states agreement was signed but cannot locate original."
        ),
    ),
    # Controller — executed agreement
    EmploymentAgreement(
        agreement_id="EA-007",
        employee_name="Thomas H. Bradley",
        employee_title="Controller",
        entity_code="DS",
        base_salary=140_000,
        bonus_target_pct=Decimal("0.20"),
        severance_multiplier=Decimal("1"),
        severance_months=6,
        change_of_control_multiplier=Decimal("1"),
        non_compete_months=6,
        ip_assignment=False,
        effective_date=datetime.date(2023, 7, 1),
        executed=True,
        notes="Standard controller agreement for Distribution Services.",
    ),
)


RETENTION_AWARDS: tuple[RetentionAward, ...] = (
    # CEO — retention to ensure continuity through transaction
    RetentionAward(
        award_id="RET-001",
        employee_name="Robert J. Cascade",
        employee_title="Chief Executive Officer",
        entity_code="CI",
        award_amount=Decimal("500_000"),
        vesting_date=datetime.date(2026, 6, 30),
        retention_period_months=18,
        forfeiture_conditions="Voluntary resignation prior to vesting date",
        notes=(
            "Board-approved retention to ensure CEO continuity through "
            "potential transaction close. Separate from severance."
        ),
    ),
    # CFO — retention award
    RetentionAward(
        award_id="RET-002",
        employee_name="Margaret L. Chen",
        employee_title="Chief Financial Officer",
        entity_code="CI",
        award_amount=Decimal("300_000"),
        vesting_date=datetime.date(2026, 6, 30),
        retention_period_months=18,
        forfeiture_conditions="Voluntary resignation prior to vesting date",
        notes="CFO retention award tied to transaction timeline.",
    ),
    # CTO — retention award
    RetentionAward(
        award_id="RET-003",
        employee_name="David R. Nakamura",
        employee_title="Chief Technology Officer",
        entity_code="CI",
        award_amount=Decimal("350_000"),
        vesting_date=datetime.date(2026, 6, 30),
        retention_period_months=18,
        forfeiture_conditions="Voluntary resignation or termination for cause",
        notes=(
            "CTO retention award — critical for AM R&D continuity. "
            "IP assignment clause makes retention strategically important."
        ),
    ),
    # R&D Director — retention (excluded from severance double-count)
    RetentionAward(
        award_id="RET-004",
        employee_name="Dr. Anika Patel",
        employee_title="R&D Director",
        entity_code="AM",
        award_amount=Decimal("150_000"),
        vesting_date=datetime.date(2026, 3, 31),
        retention_period_months=12,
        forfeiture_conditions="Voluntary resignation or termination for cause",
        notes=(
            "AM R&D key-person retention. Award is NOT additive with "
            "severance — agreement specifies greater-of, not both."
        ),
    ),
)


SEVERANCE_EXPOSURES: tuple[SeveranceExposure, ...] = (
    # CEO — termination without cause
    SeveranceExposure(
        exposure_id="SEV-001",
        employee_name="Robert J. Cascade",
        employee_title="Chief Executive Officer",
        entity_code="CI",
        base_salary=325_000,
        severance_multiplier=Decimal("3"),
        estimated_payout=Decimal("975_000"),
        trigger="change_of_control",
        accrued=False,
        notes=(
            "Golden parachute triggered by change of control. "
            "Per EA-001: 3x base salary = $975,000."
        ),
    ),
    # CFO — termination without cause
    SeveranceExposure(
        exposure_id="SEV-002",
        employee_name="Margaret L. Chen",
        employee_title="Chief Financial Officer",
        entity_code="CI",
        base_salary=260_000,
        severance_multiplier=Decimal("2"),
        estimated_payout=Decimal("520_000"),
        trigger="change_of_control",
        accrued=False,
        notes="Per EA-002: 2x base salary = $520,000.",
    ),
    # CTO — change of control
    SeveranceExposure(
        exposure_id="SEV-003",
        employee_name="David R. Nakamura",
        employee_title="Chief Technology Officer",
        entity_code="CI",
        base_salary=280_000,
        severance_multiplier=Decimal("2"),
        estimated_payout=Decimal("560_000"),
        trigger="change_of_control",
        accrued=False,
        notes="Per EA-003: 2x base salary = $560,000.",
    ),
    # VP Operations — termination without cause
    SeveranceExposure(
        exposure_id="SEV-004",
        employee_name="Susan M. Torres",
        employee_title="VP Operations",
        entity_code="PC",
        base_salary=155_000,
        severance_multiplier=Decimal("1.5"),
        estimated_payout=Decimal("232_500"),
        trigger="termination_without_cause",
        accrued=False,
        notes="Per EA-004: 1.5x base salary = $232,500.",
    ),
    # VP Sales — termination without cause
    SeveranceExposure(
        exposure_id="SEV-005",
        employee_name="James K. Whitfield",
        employee_title="VP Sales",
        entity_code="PC",
        base_salary=160_000,
        severance_multiplier=Decimal("1"),
        estimated_payout=Decimal("160_000"),
        trigger="termination_without_cause",
        accrued=False,
        notes="Per EA-005: 1x base salary = $160,000.",
    ),
    # R&D Director — change of control (greater-of with retention, not both)
    SeveranceExposure(
        exposure_id="SEV-006",
        employee_name="Dr. Anika Patel",
        employee_title="R&D Director",
        entity_code="AM",
        base_salary=170_000,
        severance_multiplier=Decimal("1.5"),
        estimated_payout=Decimal("255_000"),
        trigger="change_of_control",
        accrued=False,
        notes=(
            "Per EA-006: 1.5x base salary = $255,000. "
            "Greater-of with RET-004 ($150,000), not additive. "
            "Net exposure is $255,000, not $405,000."
        ),
    ),
    # Controller — termination without cause
    SeveranceExposure(
        exposure_id="SEV-007",
        employee_name="Thomas H. Bradley",
        employee_title="Controller",
        entity_code="DS",
        base_salary=140_000,
        severance_multiplier=Decimal("1"),
        estimated_payout=Decimal("140_000"),
        trigger="termination_without_cause",
        accrued=False,
        notes="Per EA-007: 1x base salary = $140,000.",
    ),
)


CONTRACTOR_CLASSIFICATION_SIGNALS: tuple[ContractorClassificationSignal, ...] = (
    # High risk — long-tenure exclusive contractor using company equipment
    ContractorClassificationSignal(
        signal_id="CCS-001",
        contractor_name="Martinez Technical Consulting",
        entity_code="AM",
        role_description=(
            "Senior process engineer embedded in AM manufacturing line. "
            "Performs same duties as FTE process engineers."
        ),
        tenure_months=28,
        exclusive_engagement=True,
        uses_company_equipment=True,
        has_set_hours=True,
        risk_level="high",
        notes=(
            "Multiple classification risk factors present. Recommend "
            "engagement review by employment counsel. This is a signal "
            "for follow-up, not a legal determination of misclassification."
        ),
    ),
    # Medium risk — long tenure but not exclusive
    ContractorClassificationSignal(
        signal_id="CCS-002",
        contractor_name="Pinnacle IT Solutions",
        entity_code="CI",
        role_description=(
            "IT infrastructure support for CI corporate systems. "
            "Also serves other clients in the Portland metro area."
        ),
        tenure_months=36,
        exclusive_engagement=False,
        uses_company_equipment=False,
        has_set_hours=False,
        risk_level="medium",
        notes=(
            "Long tenure but non-exclusive engagement with own equipment. "
            "Lower risk profile. Recommend periodic review of engagement "
            "terms as a matter of good practice."
        ),
    ),
    # High risk — set hours and company equipment
    ContractorClassificationSignal(
        signal_id="CCS-003",
        contractor_name="RDL Engineering Services",
        entity_code="PC",
        role_description=(
            "Quality assurance inspector on PC production floor. "
            "Reports to QA manager, follows shift schedule."
        ),
        tenure_months=19,
        exclusive_engagement=True,
        uses_company_equipment=True,
        has_set_hours=True,
        risk_level="high",
        notes=(
            "Reports to internal manager with set shift schedule. "
            "Recommend engagement review by employment counsel. "
            "This assessment is a preliminary signal, not a legal conclusion."
        ),
    ),
    # Low risk — short tenure, independent
    ContractorClassificationSignal(
        signal_id="CCS-004",
        contractor_name="GreenField Environmental Consulting",
        entity_code="AM",
        role_description=(
            "Environmental compliance advisor for AM facility expansion. "
            "Project-based engagement with defined deliverables."
        ),
        tenure_months=6,
        exclusive_engagement=False,
        uses_company_equipment=False,
        has_set_hours=False,
        risk_level="low",
        notes=(
            "Short-term project engagement with clear deliverable scope. "
            "No classification risk factors identified at this time."
        ),
    ),
)


DILIGENCE_REQUESTS: tuple[DiligenceRequest, ...] = (
    DiligenceRequest(
        request_id="DR-001",
        category="hr",
        description="Complete employee census with hire dates, titles, compensation, and termination status",
        requested_from="hr",
        status="received",
        due_date=datetime.date(2025, 10, 15),
        received_date=datetime.date(2025, 10, 12),
        source_refs=("employee_census.xlsx",),
        notes="Full 850-employee roster provided by HR.",
    ),
    DiligenceRequest(
        request_id="DR-002",
        category="hr",
        description="Executed employment agreements for all officers and key executives",
        requested_from="legal",
        status="partial",
        due_date=datetime.date(2025, 10, 15),
        received_date=datetime.date(2025, 10, 14),
        source_refs=("EA-001", "EA-002", "EA-003", "EA-004", "EA-005", "EA-007"),
        notes=(
            "Six of seven agreements received. EA-006 (Dr. Patel, R&D "
            "Director) — executed copy not located. Draft provided instead."
        ),
    ),
    DiligenceRequest(
        request_id="DR-003",
        category="hr",
        description="Retention plan documentation including board approvals",
        requested_from="hr",
        status="received",
        due_date=datetime.date(2025, 10, 20),
        received_date=datetime.date(2025, 10, 18),
        source_refs=("RET-001", "RET-002", "RET-003", "RET-004"),
        notes="All four retention award letters provided with board resolutions.",
    ),
    DiligenceRequest(
        request_id="DR-004",
        category="hr",
        description="Contractor and consultant roster with engagement terms",
        requested_from="hr",
        status="received",
        due_date=datetime.date(2025, 10, 20),
        received_date=datetime.date(2025, 10, 21),
        source_refs=("contractor_roster.xlsx",),
        notes="Received one day late. Roster covers all four entities.",
    ),
    DiligenceRequest(
        request_id="DR-005",
        category="hr",
        description="Severance schedule and change-of-control provisions summary",
        requested_from="legal",
        status="received",
        due_date=datetime.date(2025, 10, 25),
        received_date=datetime.date(2025, 10, 23),
        source_refs=("severance_schedule.xlsx",),
        notes="Schedule includes all seven executive severance arrangements.",
    ),
    DiligenceRequest(
        request_id="DR-006",
        category="legal",
        description="Executed copy of R&D Director employment agreement (EA-006)",
        requested_from="legal",
        status="not_received",
        due_date=datetime.date(2025, 11, 1),
        received_date=None,
        source_refs=(),
        notes=(
            "Follow-up request for missing executed agreement. "
            "Management unable to locate signed original as of request date."
        ),
    ),
)


# ── Query helpers ────────────────────────────────────────────────────────────


def missing_executed_agreements() -> list[EmploymentAgreement]:
    """Return employment agreements where executed copy is not on file."""
    return [ea for ea in EMPLOYMENT_AGREEMENTS if not ea.executed]


def total_severance_exposure() -> Decimal:
    """Sum of all estimated severance payouts."""
    return sum(s.estimated_payout for s in SEVERANCE_EXPOSURES)


def total_retention_awards() -> Decimal:
    """Sum of all retention award amounts."""
    return sum(r.award_amount for r in RETENTION_AWARDS)


def high_risk_contractors() -> list[ContractorClassificationSignal]:
    """Return contractor signals with high classification risk."""
    return [c for c in CONTRACTOR_CLASSIFICATION_SIGNALS if c.risk_level == "high"]


def open_diligence_requests() -> list[DiligenceRequest]:
    """Return diligence requests that are not yet fully received."""
    return [
        dr for dr in DILIGENCE_REQUESTS
        if dr.status in ("open", "partial", "not_received")
    ]


def severance_exposure_for_employee(employee_name: str) -> SeveranceExposure | None:
    """Look up severance exposure by employee name."""
    for sev in SEVERANCE_EXPOSURES:
        if sev.employee_name == employee_name:
            return sev
    return None


def retention_award_for_employee(employee_name: str) -> RetentionAward | None:
    """Look up retention award by employee name."""
    for ret in RETENTION_AWARDS:
        if ret.employee_name == employee_name:
            return ret
    return None
