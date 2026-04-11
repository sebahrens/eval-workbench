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
from dataclasses import dataclass, field
from decimal import Decimal
from typing import TYPE_CHECKING

from generator.model.ap_ledger import APLedgerResult, generate_ap_ledger
from generator.model.ar import (
    AllowanceAnalysis,
    ARAgingEntry,
    Invoice,
    MonthlyCollection,
    Receipt,
    generate_allowance,
    generate_ar_aging,
    generate_collections,
    generate_invoices,
    generate_receipts,
)
from generator.model.bank import BankModel, generate_bank_model, post_bank_to_gl
from generator.model.entities import (
    ENTITIES,
    SUBSIDIARIES,
    Entity,
    entities_from_config,
)

if TYPE_CHECKING:
    from generator.config import Config
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
from generator.model.rd import (
    QREResult,
    RDSupplyExpense,
    TimeRecord,
    compute_qres,
    generate_supply_expenses,
    generate_time_records,
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

    entities: dict[str, Entity]
    subsidiaries: dict[str, Entity]
    ledger: Ledger
    revenue_records: list[MonthlyRevenue]
    employees: list[Employee]
    opex_records: list[MonthlyOpex]
    assets: list[FixedAsset]
    leases: list[Lease]
    lease_schedules: list[LeaseScheduleRow]
    tax_provisions: dict[int, TaxProvision]
    # AR lifecycle (sales-to-cash chain)
    ar_aging: list[ARAgingEntry] = field(default_factory=list)
    ar_collections: list[MonthlyCollection] = field(default_factory=list)
    ar_allowance: list[AllowanceAnalysis] = field(default_factory=list)
    ar_invoices: list[Invoice] = field(default_factory=list)
    ar_receipts: list[Receipt] = field(default_factory=list)
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
    # R&D tax credit data (TC-08)
    rd_time_records: list[TimeRecord] = field(default_factory=list)
    rd_supply_expenses: list[RDSupplyExpense] = field(default_factory=list)
    rd_qre_result: QREResult | None = None
    # AP ledger (TC-13)
    ap_ledger: APLedgerResult | None = None


def build_model(config: Config | None = None, *, seed: int = 42) -> CascadeModel:
    """Build the complete Cascade Industries canonical model.

    Parameters
    ----------
    config : Config, optional
        Full configuration object.  When provided, the seed is taken
        from ``config.seed`` and the *seed* keyword argument is ignored.
    seed : int
        Fallback seed used when *config* is not supplied (default 42).
        Preserved for backward compatibility with tests.

    Wires: revenue, COGS, intercompany, employees, opex, PP&E,
    leases, and tax provisions.
    """
    effective_seed = config.seed if config is not None else seed
    rng = random.Random(effective_seed)

    # Derive entity maps from config (or fall back to hardcoded constants)
    if config is not None:
        all_entities, sub_entities = entities_from_config(config.company)
    else:
        all_entities, sub_entities = ENTITIES, SUBSIDIARIES

    ledger = Ledger()

    # ── Revenue & COGS ──────────────────────────────────────────────
    revenue_records = generate_monthly_revenue(rng, config=config)
    post_revenue_to_gl(ledger, revenue_records)

    # ── AR aging, collections, & allowance ────────────────────────
    # Pre-compute AR data for the model but do NOT post collections
    # or allowance to the GL.  The model's GL is accrual-based with
    # subsidiary cash flows handled centrally through the parent bank
    # model (TC-02).  Posting collections would credit 1100 (a WC asset)
    # without a matching WC offset, causing a large NWC distortion.
    # The lifecycle data (aging, collections, invoices, receipts) is
    # stored on the model for formatters and validation to consume.
    ar_aging = generate_ar_aging(revenue_records, year=2025)
    ar_collections = generate_collections(revenue_records)
    ar_allowance = generate_allowance(revenue_records)

    # ── AR lifecycle records (invoices → receipts) ─────────────────
    ar_invoices = generate_invoices(revenue_records)
    ar_receipts = generate_receipts(ar_invoices, ar_collections)

    # ── Intercompany ────────────────────────────────────────────────
    totals: dict[tuple[str, int, int], Decimal] = {}
    for r in revenue_records:
        key = (r.entity_code, r.year, r.month)
        totals[key] = totals.get(key, Decimal(0)) + r.revenue

    ic_txns = generate_ic_transactions(totals)
    post_ic_loan_principal(ledger)
    post_ic_transactions_to_gl(ledger, ic_txns)

    # ── Employees ───────────────────────────────────────────────────
    employees = generate_employees(rng, config=config)

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
    bank_rng = random.Random(effective_seed + 2)  # Isolated RNG for bank model
    bank_model = generate_bank_model(bank_rng)
    post_bank_to_gl(ledger, bank_model)

    # ── R&D time records & supply expenses (TC-08) ──────────────────
    rd_time_rng = random.Random(effective_seed)
    rd_time_records = generate_time_records(employees, rd_time_rng)
    rd_supply_rng = random.Random(effective_seed)
    rd_supply_expenses = generate_supply_expenses(rd_supply_rng)
    rd_qre_result = compute_qres(rd_time_records, rd_supply_expenses, employees)

    # ── AP ledger (TC-13) ─────────────────────────────────────────
    ap_rng = random.Random(effective_seed)
    ap_ledger_result = generate_ap_ledger(ap_rng, employees)

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
        entities=all_entities,
        subsidiaries=sub_entities,
        ledger=ledger,
        revenue_records=revenue_records,
        employees=employees,
        opex_records=opex_records,
        assets=assets,
        leases=leases,
        lease_schedules=lease_schedules,
        tax_provisions=tax_provisions,
        # AR lifecycle
        ar_aging=ar_aging,
        ar_collections=ar_collections,
        ar_allowance=ar_allowance,
        ar_invoices=ar_invoices,
        ar_receipts=ar_receipts,
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
        # R&D (TC-08)
        rd_time_records=rd_time_records,
        rd_supply_expenses=rd_supply_expenses,
        rd_qre_result=rd_qre_result,
        # AP ledger (TC-13)
        ap_ledger=ap_ledger_result,
    )
