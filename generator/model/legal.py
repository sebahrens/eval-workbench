"""Cascade Industries legal contract and clause canonical model.

Provides deterministic contracts, clauses, amendments, and legal diligence
issues for M&A legal diligence test cases (TC-19, TC-21).

Cross-references existing model objects:
- Customer IDs (CUST-NNN) from ar.py
- Vendor IDs (VEND-NNN) from ap.py
- Entity codes (CI, PC, AM, DS) from entities.py
- Customer contracts (CTR-NNN) from customers.py

All data is hardcoded tuples for determinism.  ID conventions follow the
design decision in bead synth-data-ups.6:
  LCTR-NNN, CLS-NNN, AMD-NNN, LDI-NNN
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass
from decimal import Decimal

# ── Dataclasses ──────────────────────────────────────────────────────────────


@dataclass(frozen=True)
class LegalContract:
    """A legal contract for diligence purposes.

    Extends beyond CustomerContract (customers.py) with governing law,
    auto-renew terms, and stable LCTR- IDs.  Both coexist; this serves
    TC-19/TC-21.
    """

    contract_id: str  # LCTR-NNN
    counterparty_id: str  # CUST-NNN or VEND-NNN
    counterparty_name: str
    entity_code: str  # CI, PC, AM, DS
    contract_type: str  # supply, service, license, distribution, lease, employment
    effective_date: datetime.date
    expiration_date: datetime.date
    annual_value: Decimal
    governing_law: str  # state/jurisdiction
    auto_renew: bool
    notes: str = ""


@dataclass(frozen=True)
class ContractClause:
    """A notable clause within a legal contract."""

    clause_id: str  # CLS-NNN
    contract_id: str  # LCTR-NNN reference
    clause_type: str  # change_of_control, non_compete, mfn, exclusivity, etc.
    summary: str
    risk_level: str  # high, medium, low
    business_impact: str  # revenue_at_risk, operational, compliance, retention
    source_section: str  # e.g. "Section 12.3"
    notes: str = ""


@dataclass(frozen=True)
class ContractAmendment:
    """An amendment or side letter modifying a contract."""

    amendment_id: str  # AMD-NNN
    contract_id: str  # LCTR-NNN reference
    effective_date: datetime.date
    description: str
    changes_clause_ids: tuple[str, ...]  # CLS-NNN references affected
    supersedes_original: bool  # does this replace the original clause entirely?
    notes: str = ""


@dataclass(frozen=True)
class LegalDiligenceIssue:
    """A legal diligence issue or finding.

    Issues represent findings that an agent must surface during contract
    review.  They may be clause-specific or contract-level.
    """

    issue_id: str  # LDI-NNN
    contract_id: str  # LCTR-NNN reference
    clause_id: str | None  # CLS-NNN if clause-specific, None if contract-level
    issue_type: str  # missing_consent, stale_document, contradicts_summary, scope_boundary
    severity: str  # high, medium, low
    description: str
    recommended_action: str  # flag, investigate, request_document, escalate
    source_refs: tuple[str, ...]  # document IDs or section references
    notes: str = ""


# ── Canonical Data ───────────────────────────────────────────────────────────
# All values are deterministic and cross-reference existing Cascade entities.
# Counterparty IDs tie to CUST-NNN (ar.py) or VEND-NNN (ap.py).

LEGAL_CONTRACTS: tuple[LegalContract, ...] = (
    # ── Customer contracts ────────────────────────────────────────────────
    # Acme Manufacturing — largest customer, change-of-control risk
    LegalContract(
        contract_id="LCTR-001",
        counterparty_id="CUST-001",
        counterparty_name="Acme Manufacturing Corp",
        entity_code="PC",
        contract_type="supply",
        effective_date=datetime.date(2022, 1, 15),
        expiration_date=datetime.date(2025, 12, 31),
        annual_value=Decimal("36_000_000"),
        governing_law="Oregon",
        auto_renew=False,
        notes=(
            "Master supply agreement. Largest customer (~18% consolidated "
            "revenue). Change-of-control provision in Section 14.2."
        ),
    ),
    # Northwest Precision — auto-renewing
    LegalContract(
        contract_id="LCTR-002",
        counterparty_id="CUST-002",
        counterparty_name="Northwest Precision Industries",
        entity_code="PC",
        contract_type="supply",
        effective_date=datetime.date(2023, 4, 1),
        expiration_date=datetime.date(2026, 3, 31),
        annual_value=Decimal("14_250_000"),
        governing_law="Oregon",
        auto_renew=True,
    ),
    # TechAlloy Systems — MFN clause
    LegalContract(
        contract_id="LCTR-003",
        counterparty_id="CUST-011",
        counterparty_name="TechAlloy Systems",
        entity_code="AM",
        contract_type="supply",
        effective_date=datetime.date(2023, 1, 1),
        expiration_date=datetime.date(2027, 12, 31),
        annual_value=Decimal("14_300_000"),
        governing_law="Oregon",
        auto_renew=True,
        notes="Contains most-favored-nation pricing clause (Section 8.1).",
    ),
    # NextGen Composites — government subcontract, exclusivity
    LegalContract(
        contract_id="LCTR-004",
        counterparty_id="CUST-013",
        counterparty_name="NextGen Composites LLC",
        entity_code="AM",
        contract_type="supply",
        effective_date=datetime.date(2024, 3, 1),
        expiration_date=datetime.date(2028, 2, 28),
        annual_value=Decimal("9_750_000"),
        governing_law="Virginia",
        auto_renew=False,
        notes=(
            "Government subcontract; subject to DFARS flow-down provisions. "
            "Exclusivity clause restricts AM from supplying competing "
            "defense contractors for identical alloy specs."
        ),
    ),
    # GlobalTrade Logistics — DS customer
    LegalContract(
        contract_id="LCTR-005",
        counterparty_id="CUST-019",
        counterparty_name="GlobalTrade Logistics",
        entity_code="DS",
        contract_type="service",
        effective_date=datetime.date(2023, 9, 1),
        expiration_date=datetime.date(2026, 8, 31),
        annual_value=Decimal("8_000_000"),
        governing_law="Oregon",
        auto_renew=True,
    ),
    # Columbia River Works — expiring soon
    LegalContract(
        contract_id="LCTR-006",
        counterparty_id="CUST-003",
        counterparty_name="Columbia River Works",
        entity_code="PC",
        contract_type="supply",
        effective_date=datetime.date(2021, 7, 1),
        expiration_date=datetime.date(2025, 6, 30),
        annual_value=Decimal("11_400_000"),
        governing_law="Oregon",
        auto_renew=True,
        notes="Expiring within 12 months of FY2025 — renewal risk.",
    ),
    # Quantum Materials — expiring, no auto-renew
    LegalContract(
        contract_id="LCTR-007",
        counterparty_id="CUST-012",
        counterparty_name="Quantum Materials Inc",
        entity_code="AM",
        contract_type="supply",
        effective_date=datetime.date(2022, 6, 1),
        expiration_date=datetime.date(2025, 5, 31),
        annual_value=Decimal("11_700_000"),
        governing_law="Oregon",
        auto_renew=False,
        notes="Expiring within 12 months of FY2025 — renewal risk.",
    ),
    # Continental Supply Chain — DS customer
    LegalContract(
        contract_id="LCTR-008",
        counterparty_id="CUST-020",
        counterparty_name="Continental Supply Chain",
        entity_code="DS",
        contract_type="service",
        effective_date=datetime.date(2024, 1, 1),
        expiration_date=datetime.date(2025, 12, 31),
        annual_value=Decimal("6_400_000"),
        governing_law="Oregon",
        auto_renew=True,
    ),
    # ── Vendor contracts ──────────────────────────────────────────────────
    # Portland Steel Supply — largest PC vendor
    LegalContract(
        contract_id="LCTR-009",
        counterparty_id="VEND-001",
        counterparty_name="Portland Steel Supply",
        entity_code="PC",
        contract_type="supply",
        effective_date=datetime.date(2022, 3, 1),
        expiration_date=datetime.date(2026, 2, 28),
        annual_value=Decimal("5_200_000"),
        governing_law="Oregon",
        auto_renew=True,
        notes="Primary raw material supplier for PC operations.",
    ),
    # ChemSource International — largest AM vendor, IP license
    LegalContract(
        contract_id="LCTR-010",
        counterparty_id="VEND-011",
        counterparty_name="ChemSource International",
        entity_code="AM",
        contract_type="license",
        effective_date=datetime.date(2021, 6, 1),
        expiration_date=datetime.date(2026, 5, 31),
        annual_value=Decimal("4_600_000"),
        governing_law="Delaware",
        auto_renew=False,
        notes=(
            "Combined supply and technology license agreement. "
            "IP license covers proprietary coating formulations used "
            "in AM advanced materials line."
        ),
    ),
)


CONTRACT_CLAUSES: tuple[ContractClause, ...] = (
    # ── LCTR-001 (Acme) ──────────────────────────────────────────────────
    # Change-of-control — high risk (TC-19 required trap)
    ContractClause(
        clause_id="CLS-001",
        contract_id="LCTR-001",
        clause_type="change_of_control",
        summary=(
            "Acme may terminate within 90 days of any change in ownership "
            "or control of Cascade Precision Components without penalty."
        ),
        risk_level="high",
        business_impact="revenue_at_risk",
        source_section="Section 14.2",
        notes=(
            "Acme represents ~18% of consolidated revenue. "
            "Termination without penalty creates material transaction risk."
        ),
    ),
    # Indemnification — standard
    ContractClause(
        clause_id="CLS-002",
        contract_id="LCTR-001",
        clause_type="indemnification",
        summary="Mutual indemnification with carve-out for gross negligence.",
        risk_level="low",
        business_impact="compliance",
        source_section="Section 11.1",
    ),
    # ── LCTR-003 (TechAlloy) ─────────────────────────────────────────────
    # MFN pricing — medium risk (TC-19 required trap)
    ContractClause(
        clause_id="CLS-003",
        contract_id="LCTR-003",
        clause_type="mfn",
        summary=(
            "TechAlloy is entitled to pricing no less favorable than any "
            "comparable customer for equivalent volume and alloy grades. "
            "Cascade must notify TechAlloy of better terms within 30 days."
        ),
        risk_level="medium",
        business_impact="revenue_at_risk",
        source_section="Section 8.1",
        notes=(
            "MFN clause triggered if AM offers better pricing to another "
            "customer at comparable volume. Could limit margin improvement "
            "across the AM portfolio."
        ),
    ),
    # IP assignment — AM technology
    ContractClause(
        clause_id="CLS-004",
        contract_id="LCTR-003",
        clause_type="ip_assignment",
        summary=(
            "Joint IP developed under this agreement is co-owned. "
            "Cascade retains rights to background IP."
        ),
        risk_level="medium",
        business_impact="operational",
        source_section="Section 15.2",
    ),
    # ── LCTR-004 (NextGen) ───────────────────────────────────────────────
    # Exclusivity — medium risk (TC-19 required trap)
    ContractClause(
        clause_id="CLS-005",
        contract_id="LCTR-004",
        clause_type="exclusivity",
        summary=(
            "Cascade Advanced Materials may not supply identical alloy "
            "specifications to competing defense contractors for the "
            "duration of the agreement."
        ),
        risk_level="medium",
        business_impact="revenue_at_risk",
        source_section="Section 6.4",
        notes=(
            "Restricts AM's ability to grow defense-sector revenue "
            "with competing primes. Scope limited to identical specs — "
            "derivative formulations are not restricted."
        ),
    ),
    # Assignment — government subcontract
    ContractClause(
        clause_id="CLS-006",
        contract_id="LCTR-004",
        clause_type="assignment",
        summary=(
            "Assignment requires prior written consent of the prime "
            "contractor and, where applicable, the contracting officer."
        ),
        risk_level="high",
        business_impact="compliance",
        source_section="Section 22.1",
        notes=(
            "Government subcontract assignment requires consent — "
            "change of ownership may trigger novation requirement."
        ),
    ),
    # ── LCTR-010 (ChemSource) ────────────────────────────────────────────
    # Termination for convenience — vendor
    ContractClause(
        clause_id="CLS-007",
        contract_id="LCTR-010",
        clause_type="termination",
        summary=(
            "Either party may terminate with 180 days written notice. "
            "Termination of supply does not terminate the IP license, "
            "which survives for 36 months post-termination."
        ),
        risk_level="medium",
        business_impact="operational",
        source_section="Section 19.3",
    ),
    # IP license — ChemSource formulations
    ContractClause(
        clause_id="CLS-008",
        contract_id="LCTR-010",
        clause_type="ip_assignment",
        summary=(
            "ChemSource grants non-exclusive license to proprietary "
            "coating formulations for AM advanced materials line. "
            "License fee included in supply pricing."
        ),
        risk_level="medium",
        business_impact="operational",
        source_section="Section 20.1",
        notes=(
            "IP license survives contract termination for 36 months "
            "(Section 19.3). After that, AM loses rights to formulations. "
            "Critical dependency for advanced materials product line."
        ),
    ),
    # ── LCTR-009 (Portland Steel) ────────────────────────────────────────
    # Non-compete — low risk
    ContractClause(
        clause_id="CLS-009",
        contract_id="LCTR-009",
        clause_type="non_compete",
        summary=(
            "Portland Steel agrees not to supply Cascade competitors "
            "in the Pacific Northwest precision components market "
            "during the contract term."
        ),
        risk_level="low",
        business_impact="operational",
        source_section="Section 9.1",
    ),
)


CONTRACT_AMENDMENTS: tuple[ContractAmendment, ...] = (
    # Amendment to Acme contract — extends pricing lock
    ContractAmendment(
        amendment_id="AMD-001",
        contract_id="LCTR-001",
        effective_date=datetime.date(2024, 7, 1),
        description=(
            "Amendment extends CPI adjustment cap from 3% to 4% for "
            "contract years 2024-2025. All other terms unchanged."
        ),
        changes_clause_ids=(),
        supersedes_original=False,
        notes="Pricing modification only; does not affect Section 14.2 (CoC).",
    ),
    # Side letter to TechAlloy — modifies MFN threshold (TC-19 trap: missing amendment)
    ContractAmendment(
        amendment_id="AMD-002",
        contract_id="LCTR-003",
        effective_date=datetime.date(2025, 1, 15),
        description=(
            "Side letter raises the MFN notification threshold from "
            "any pricing difference to differences exceeding 5% on a "
            "per-kg basis. Reduces MFN trigger frequency."
        ),
        changes_clause_ids=("CLS-003",),
        supersedes_original=False,
        notes=(
            "This side letter materially narrows the MFN clause. "
            "Management summary memo does not reference this amendment — "
            "creating a contradicts_summary diligence issue."
        ),
    ),
    # Amendment to NextGen — scope expansion
    ContractAmendment(
        amendment_id="AMD-003",
        contract_id="LCTR-004",
        effective_date=datetime.date(2025, 3, 1),
        description=(
            "Adds Phase II deliverables expanding scope to include "
            "thermal barrier coatings. Annual value increases by $2.1M. "
            "Exclusivity clause (Section 6.4) extended to cover new specs."
        ),
        changes_clause_ids=("CLS-005",),
        supersedes_original=False,
        notes=(
            "Expands exclusivity scope to thermal barrier coatings. "
            "Original exclusivity was limited to alloy specs only."
        ),
    ),
)


LEGAL_DILIGENCE_ISSUES: tuple[LegalDiligenceIssue, ...] = (
    # Change-of-control consent requirement — Acme (TC-19 required trap)
    LegalDiligenceIssue(
        issue_id="LDI-001",
        contract_id="LCTR-001",
        clause_id="CLS-001",
        issue_type="missing_consent",
        severity="high",
        description=(
            "Acme Manufacturing contract (LCTR-001) contains a "
            "change-of-control termination clause (Section 14.2). "
            "No waiver or consent has been obtained from Acme. "
            "Transaction close requires either Acme consent or "
            "risk assessment of potential $36M annual revenue loss."
        ),
        recommended_action="escalate",
        source_refs=("LCTR-001", "CLS-001", "Section 14.2"),
        notes=(
            "Highest-priority diligence item. Revenue at risk represents "
            "~18% of consolidated revenue."
        ),
    ),
    # Management summary contradicts primary contract — TechAlloy MFN
    # (TC-19 required trap: stale-summary contradiction)
    LegalDiligenceIssue(
        issue_id="LDI-002",
        contract_id="LCTR-003",
        clause_id="CLS-003",
        issue_type="contradicts_summary",
        severity="high",
        description=(
            "Management summary memo describes TechAlloy MFN clause as "
            "triggering on any pricing difference. However, side letter "
            "AMD-002 (effective 2025-01-15) raised the threshold to "
            "differences exceeding 5% per kg. The summary is stale and "
            "does not reflect the current contractual position."
        ),
        recommended_action="flag",
        source_refs=("LCTR-003", "CLS-003", "AMD-002", "management_summary_memo"),
        notes=(
            "Agent must identify that the management summary contradicts "
            "the actual contract terms as amended. Source reliance trap."
        ),
    ),
    # Missing amendment — NextGen exclusivity expansion not in summary
    # (TC-19 required trap: missing amendment or side letter)
    LegalDiligenceIssue(
        issue_id="LDI-003",
        contract_id="LCTR-004",
        clause_id="CLS-005",
        issue_type="stale_document",
        severity="medium",
        description=(
            "The exclusivity clause summary in the diligence binder "
            "references only the original alloy specs exclusivity. "
            "Amendment AMD-003 (effective 2025-03-01) expanded exclusivity "
            "to include thermal barrier coatings, but this expansion is "
            "not reflected in the current diligence summary."
        ),
        recommended_action="investigate",
        source_refs=("LCTR-004", "CLS-005", "AMD-003"),
        notes=(
            "Agent should cross-reference amendments against clause "
            "summaries and flag any out-of-date characterizations."
        ),
    ),
    # Government subcontract assignment consent — NextGen
    LegalDiligenceIssue(
        issue_id="LDI-004",
        contract_id="LCTR-004",
        clause_id="CLS-006",
        issue_type="missing_consent",
        severity="high",
        description=(
            "NextGen government subcontract (LCTR-004) requires prior "
            "written consent for assignment (Section 22.1). Change of "
            "ownership may trigger novation requirement under FAR 42.12. "
            "No consent or novation request has been initiated."
        ),
        recommended_action="escalate",
        source_refs=("LCTR-004", "CLS-006", "Section 22.1"),
    ),
    # IP license dependency — ChemSource
    LegalDiligenceIssue(
        issue_id="LDI-005",
        contract_id="LCTR-010",
        clause_id="CLS-008",
        issue_type="scope_boundary",
        severity="medium",
        description=(
            "ChemSource IP license (LCTR-010, Section 20.1) survives "
            "contract termination for only 36 months. If the supply "
            "relationship ends post-acquisition, AM loses rights to "
            "proprietary coating formulations after the survival period. "
            "This creates a critical product-line dependency."
        ),
        recommended_action="investigate",
        source_refs=("LCTR-010", "CLS-007", "CLS-008", "Section 19.3", "Section 20.1"),
        notes=(
            "Scope boundary issue — not a defect in the contract but a "
            "strategic risk that should be flagged in diligence findings."
        ),
    ),
)


# ── Query helpers ────────────────────────────────────────────────────────────


def contracts_by_entity(entity_code: str) -> list[LegalContract]:
    """Return legal contracts for a given entity."""
    return [c for c in LEGAL_CONTRACTS if c.entity_code == entity_code]


def clauses_for_contract(contract_id: str) -> list[ContractClause]:
    """Return all clauses for a given contract."""
    return [cl for cl in CONTRACT_CLAUSES if cl.contract_id == contract_id]


def amendments_for_contract(contract_id: str) -> list[ContractAmendment]:
    """Return all amendments for a given contract."""
    return [a for a in CONTRACT_AMENDMENTS if a.contract_id == contract_id]


def high_risk_clauses() -> list[ContractClause]:
    """Return all clauses with high risk level."""
    return [cl for cl in CONTRACT_CLAUSES if cl.risk_level == "high"]


def change_of_control_clauses() -> list[ContractClause]:
    """Return clauses that are change-of-control provisions."""
    return [cl for cl in CONTRACT_CLAUSES if cl.clause_type == "change_of_control"]


def issues_for_contract(contract_id: str) -> list[LegalDiligenceIssue]:
    """Return all diligence issues for a given contract."""
    return [i for i in LEGAL_DILIGENCE_ISSUES if i.contract_id == contract_id]


def issues_by_severity(severity: str) -> list[LegalDiligenceIssue]:
    """Return diligence issues filtered by severity."""
    return [i for i in LEGAL_DILIGENCE_ISSUES if i.severity == severity]


def total_revenue_at_risk() -> Decimal:
    """Sum annual values of contracts with high-severity diligence issues."""
    contract_ids = {
        i.contract_id for i in LEGAL_DILIGENCE_ISSUES if i.severity == "high"
    }
    return sum(
        c.annual_value for c in LEGAL_CONTRACTS if c.contract_id in contract_ids
    )


def stale_summary_issues() -> list[LegalDiligenceIssue]:
    """Return issues where management summary contradicts primary documents."""
    return [
        i for i in LEGAL_DILIGENCE_ISSUES
        if i.issue_type in ("contradicts_summary", "stale_document")
    ]


def missing_consent_issues() -> list[LegalDiligenceIssue]:
    """Return issues where consent has not been obtained."""
    return [
        i for i in LEGAL_DILIGENCE_ISSUES if i.issue_type == "missing_consent"
    ]
