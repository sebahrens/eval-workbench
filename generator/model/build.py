"""Canonical model builder — Phase 1 of the two-phase architecture.

Populates a single GL Ledger with all entity activity (revenue, COGS,
opex, PP&E, AR, AP, IC, leases, tax) for FY2023–FY2025.  Phase 2
formatters read from this ledger via the views module.

Returns a CascadeModel dataclass that bundles the ledger with all
auxiliary model objects (employees, leases, assets, etc.) so
formatters have everything they need.
"""

from __future__ import annotations

import random
from dataclasses import dataclass
from decimal import Decimal

from generator.model.bank import BankModel, generate_bank_model, post_bank_to_gl
from generator.model.employees import Employee, generate_employees
from generator.model.gl import Ledger
from generator.model.hr_diligence import (
    CONTRACTOR_CLASSIFICATION_SIGNALS,
    DILIGENCE_REQUESTS,
    EMPLOYMENT_AGREEMENTS,
    RETENTION_AWARDS,
    SEVERANCE_EXPOSURES,
    ContractorClassificationSignal,
    DiligenceRequest,
    EmploymentAgreement,
    RetentionAward,
    SeveranceExposure,
)
from generator.model.intercompany import (
    generate_ic_transactions,
    post_ic_loan_principal,
    post_ic_transactions_to_gl,
)
from generator.model.leases import (
    Lease,
    LeaseScheduleRow,
    compute_lease_schedules,
    generate_leases,
    post_leases_to_gl,
)
from generator.model.legal import (
    CONTRACT_AMENDMENTS,
    CONTRACT_CLAUSES,
    LEGAL_CONTRACTS,
    LEGAL_DILIGENCE_ISSUES,
    ContractAmendment,
    ContractClause,
    LegalContract,
    LegalDiligenceIssue,
)
from generator.model.opex import MonthlyOpex, generate_opex, post_opex_to_gl
from generator.model.ppe import (
    FixedAsset,
    generate_fixed_assets,
    post_asset_acquisitions_to_gl,
    post_depreciation_to_gl,
)
from generator.model.revenue import (
    MonthlyRevenue,
    generate_monthly_revenue,
    post_revenue_to_gl,
)
from generator.model.tax import TaxProvision, compute_provisions_multi_year, post_tax_provision_to_gl
from generator.model.views import build_income_statement


@dataclass
class CascadeModel:
    """Bundle of all canonical model objects.

    Passed to Phase 2 formatters so they can generate files from a
    single source of truth.
    """

    ledger: Ledger
    revenue_records: list[MonthlyRevenue]
    employees: list[Employee]
    opex_records: list[MonthlyOpex]
    assets: list[FixedAsset]
    leases: list[Lease]
    lease_schedules: list[LeaseScheduleRow]
    tax_provisions: dict[int, TaxProvision]
    bank: BankModel | None = None
    # Legal diligence (TC-19, TC-21)
    legal_contracts: tuple[LegalContract, ...] = ()
    contract_clauses: tuple[ContractClause, ...] = ()
    contract_amendments: tuple[ContractAmendment, ...] = ()
    legal_diligence_issues: tuple[LegalDiligenceIssue, ...] = ()
    # HR diligence (TC-20, TC-21)
    employment_agreements: tuple[EmploymentAgreement, ...] = ()
    retention_awards: tuple[RetentionAward, ...] = ()
    severance_exposures: tuple[SeveranceExposure, ...] = ()
    contractor_signals: tuple[ContractorClassificationSignal, ...] = ()
    diligence_requests: tuple[DiligenceRequest, ...] = ()


def build_model(seed: int = 42) -> CascadeModel:
    """Build the complete Cascade Industries canonical model.

    Wires: revenue, COGS, intercompany, employees, opex, PP&E,
    leases, and tax provisions.
    """
    rng = random.Random(seed)
    ledger = Ledger()

    # ── Revenue & COGS ──────────────────────────────────────────────
    revenue_records = generate_monthly_revenue(rng)
    post_revenue_to_gl(ledger, revenue_records)

    # ── Intercompany ────────────────────────────────────────────────
    totals: dict[tuple[str, int, int], Decimal] = {}
    for r in revenue_records:
        key = (r.entity_code, r.year, r.month)
        totals[key] = totals.get(key, Decimal(0)) + r.revenue

    ic_txns = generate_ic_transactions(totals)
    post_ic_loan_principal(ledger)
    post_ic_transactions_to_gl(ledger, ic_txns)

    # ── Employees ───────────────────────────────────────────────────
    employees = generate_employees(rng)

    # ── Operating expenses ──────────────────────────────────────────
    rev_by_entity_year: dict[tuple[str, int], Decimal] = {}
    for r in revenue_records:
        key = (r.entity_code, r.year)
        rev_by_entity_year[key] = rev_by_entity_year.get(key, Decimal(0)) + r.revenue

    opex_records = generate_opex(rng, employees, rev_by_entity_year)
    post_opex_to_gl(ledger, opex_records)

    # ── PP&E ────────────────────────────────────────────────────────
    assets = generate_fixed_assets(rng)
    post_asset_acquisitions_to_gl(ledger, assets)
    post_depreciation_to_gl(ledger, assets)

    # ── Leases ──────────────────────────────────────────────────────
    leases = generate_leases(rng)
    lease_schedules = compute_lease_schedules(leases)
    post_leases_to_gl(ledger, leases, lease_schedules)

    # ── Bank transactions (TC-02) ─────────────────────────────────
    bank_rng = random.Random(seed + 2)  # Isolated RNG for bank model
    bank_model = generate_bank_model(bank_rng)
    post_bank_to_gl(ledger, bank_model)

    # ── Tax provisions ──────────────────────────────────────────────
    # Pre-tax income comes from the income statement (built from GL)
    # computed BEFORE tax entries so we get the right base.
    pre_tax_by_year: dict[int, Decimal] = {}
    for year in [2023, 2024, 2025]:
        is_stmt = build_income_statement(ledger, year)
        pre_tax_by_year[year] = is_stmt.pre_tax_income

    tax_provisions = compute_provisions_multi_year(
        pre_tax_by_year, opex_records, assets, lease_schedules,
    )
    post_tax_provision_to_gl(ledger, tax_provisions)

    return CascadeModel(
        ledger=ledger,
        revenue_records=revenue_records,
        employees=employees,
        opex_records=opex_records,
        assets=assets,
        leases=leases,
        lease_schedules=lease_schedules,
        tax_provisions=tax_provisions,
        bank=bank_model,
        # Legal diligence — hardcoded canonical tuples (no RNG)
        legal_contracts=LEGAL_CONTRACTS,
        contract_clauses=CONTRACT_CLAUSES,
        contract_amendments=CONTRACT_AMENDMENTS,
        legal_diligence_issues=LEGAL_DILIGENCE_ISSUES,
        # HR diligence — hardcoded canonical tuples (no RNG)
        employment_agreements=EMPLOYMENT_AGREEMENTS,
        retention_awards=RETENTION_AWARDS,
        severance_exposures=SEVERANCE_EXPOSURES,
        contractor_signals=CONTRACTOR_CLASSIFICATION_SIGNALS,
        diligence_requests=DILIGENCE_REQUESTS,
    )
