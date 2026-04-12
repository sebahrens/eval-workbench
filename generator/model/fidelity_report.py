"""Fidelity regression report for the generated test suite (synth-data-8ok.6).

Produces a deterministic summary of cross-file tie-outs and data-shape
metrics.  The report captures:

1. Revenue / AR / cash lifecycle tie-outs
2. AP ledger and anomaly counts
3. R&D time records and QRE computation
4. Lease portfolio metrics
5. Tax provision summary
6. Data-shape counts (employees, assets, GL entries, file counts)

The report is a flat dict of string keys → string values, suitable for
JSON serialization and snapshot comparison.  Running the generator twice
with the same seed must produce identical reports.
"""

from __future__ import annotations

import datetime
from decimal import Decimal
from typing import TYPE_CHECKING

from generator.model.consolidation import (
    build_income_statement,
    verify_ic_elimination,
)

if TYPE_CHECKING:
    from generator.model.build import CascadeModel


def _fmt(d: Decimal) -> str:
    """Format a Decimal to 2 decimal places for stable output."""
    return f"{d:.2f}"


def build_fidelity_report(model: CascadeModel) -> dict[str, str]:
    """Build the fidelity regression report from a CascadeModel.

    Returns a dict with deterministic string keys and values.
    """
    report: dict[str, str] = {}

    # ── 1. Data shape counts ───────────────────────────────────────
    report["employees.total"] = str(len(model.employees))
    report["employees.active"] = str(
        sum(1 for e in model.employees if e.termination_date is None)
    )
    report["entities.total"] = str(len(model.entities))
    report["subsidiaries.total"] = str(len(model.subsidiaries))
    report["assets.total"] = str(len(model.assets))
    report["leases.total"] = str(len(model.leases))
    report["lease_schedule_rows.total"] = str(len(model.lease_schedules))
    report["gl_entries.total"] = str(len(model.ledger.entries))
    report["revenue_records.total"] = str(len(model.revenue_records))
    report["opex_records.total"] = str(len(model.opex_records))

    # ── 2. Revenue / AR / cash tie-outs ────────────────────────────
    for year in [2023, 2024, 2025]:
        is_stmt = build_income_statement(model.ledger, year)
        report[f"revenue.fy{year}.gl_total"] = _fmt(is_stmt.total_revenue)
        report[f"revenue.fy{year}.cogs"] = _fmt(is_stmt.total_cogs)
        report[f"revenue.fy{year}.gross_profit"] = _fmt(is_stmt.gross_profit)
        report[f"revenue.fy{year}.pre_tax_income"] = _fmt(is_stmt.pre_tax_income)

    # Revenue records vs GL tie-out (FY2025)
    rev_from_records_2025 = sum(
        r.revenue for r in model.revenue_records if r.year == 2025
    )
    report["revenue.fy2025.records_total"] = _fmt(rev_from_records_2025)

    # AR aging
    report["ar_aging.count"] = str(len(model.ar_aging))
    if model.ar_aging:
        aging_total = sum(e.total for e in model.ar_aging)
        report["ar_aging.total"] = _fmt(aging_total)

    # AR invoices & receipts
    report["ar_invoices.count"] = str(len(model.ar_invoices))
    report["ar_receipts.count"] = str(len(model.ar_receipts))
    if model.ar_invoices:
        inv_total = sum(i.amount for i in model.ar_invoices)
        report["ar_invoices.total"] = _fmt(inv_total)
    if model.ar_receipts:
        rct_total = sum(r.amount for r in model.ar_receipts)
        report["ar_receipts.total"] = _fmt(rct_total)

    # AR collections
    report["ar_collections.count"] = str(len(model.ar_collections))
    if model.ar_collections:
        coll_total = sum(c.amount for c in model.ar_collections)
        report["ar_collections.total"] = _fmt(coll_total)

    # AR allowance
    report["ar_allowance.count"] = str(len(model.ar_allowance))

    # Bank model
    if model.bank is not None:
        report["bank.transactions.count"] = str(len(model.bank.bank_transactions))
        bank_total = sum(t.amount for t in model.bank.bank_transactions)
        report["bank.transactions.net"] = _fmt(bank_total)
        report["bank.ending_balance"] = _fmt(model.bank.bank_ending_balance)
        report["bank.gl_ending_balance"] = _fmt(model.bank.gl_ending_balance)

    # ── 3. Intercompany netting ────────────────────────────────────
    for year in [2023, 2024, 2025]:
        eoy = datetime.date(year, 12, 31)
        is_zero, total_imbalance, _ = verify_ic_elimination(
            model.ledger, eoy,
        )
        report[f"ic_netting.fy{year}.balanced"] = str(is_zero)
        report[f"ic_netting.fy{year}.imbalance"] = _fmt(total_imbalance)

    # ── 4. AP ledger and anomalies ─────────────────────────────────
    if model.ap_ledger is not None:
        report["ap.transactions.count"] = str(len(model.ap_ledger.transactions))
        ap_total = sum(t.amount for t in model.ap_ledger.transactions)
        report["ap.transactions.total"] = _fmt(ap_total)
        report["ap.vendors.count"] = str(len(model.ap_ledger.vendor_summaries))
        # Anomaly breakdown
        anomaly_count = sum(
            len(ids) for ids in model.ap_ledger.anomaly_index.values()
        )
        report["ap.anomalies.total"] = str(anomaly_count)
        for atype in sorted(model.ap_ledger.anomaly_index):
            report[f"ap.anomalies.{atype}"] = str(
                len(model.ap_ledger.anomaly_index[atype])
            )

    # ── 5. R&D / QRE ──────────────────────────────────────────────
    report["rd.time_records.count"] = str(len(model.rd_time_records))
    report["rd.supply_expenses.count"] = str(len(model.rd_supply_expenses))
    if model.rd_qre_result is not None:
        qre = model.rd_qre_result
        report["rd.qre.year"] = str(qre.year)
        report["rd.qre.wage_qres"] = _fmt(qre.wage_qres)
        report["rd.qre.supply_qres"] = _fmt(qre.supply_qres)
        report["rd.qre.total_qres"] = _fmt(qre.wage_qres + qre.supply_qres)

    # ── 6. Leases ──────────────────────────────────────────────────
    if model.leases:
        # Sum beginning liability for the first schedule year per lease
        seen_leases: set[str] = set()
        lease_liability_total = Decimal(0)
        for ls in model.lease_schedules:
            if ls.lease_id not in seen_leases:
                seen_leases.add(ls.lease_id)
                lease_liability_total += ls.lease_liability_beg
        report["leases.initial_liability_total"] = _fmt(lease_liability_total)
        # Entity breakdown
        for ec in sorted({le.entity_code for le in model.leases}):
            count = sum(1 for le in model.leases if le.entity_code == ec)
            report[f"leases.by_entity.{ec}"] = str(count)

    # ── 7. Tax provisions ──────────────────────────────────────────
    for year in sorted(model.tax_provisions):
        tp = model.tax_provisions[year]
        report[f"tax.fy{year}.pre_tax_book_income"] = _fmt(tp.pre_tax_book_income)
        report[f"tax.fy{year}.federal_current"] = _fmt(tp.federal_current)
        report[f"tax.fy{year}.state_current"] = _fmt(tp.state_current)
        report[f"tax.fy{year}.total_provision"] = _fmt(tp.total_provision)

    # ── 8. Legal / HR diligence counts ─────────────────────────────
    report["legal.contracts.count"] = str(len(model.legal_contracts))
    report["legal.clauses.count"] = str(len(model.contract_clauses))
    report["legal.amendments.count"] = str(len(model.contract_amendments))
    report["legal.diligence_issues.count"] = str(len(model.legal_diligence_issues))
    report["hr.employment_agreements.count"] = str(len(model.employment_agreements))
    report["hr.retention_awards.count"] = str(len(model.retention_awards))
    report["hr.severance_exposures.count"] = str(len(model.severance_exposures))
    report["hr.contractor_signals.count"] = str(len(model.contractor_signals))

    return report
