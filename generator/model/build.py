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

from generator.model.gl import Ledger
from generator.model.intercompany import (
    generate_ic_transactions,
    post_ic_loan_principal,
    post_ic_transactions_to_gl,
)
from generator.model.revenue import (
    MonthlyRevenue,
    generate_monthly_revenue,
    post_revenue_to_gl,
)


@dataclass
class CascadeModel:
    """Bundle of all canonical model objects.

    Passed to Phase 2 formatters so they can generate files from a
    single source of truth.
    """

    ledger: Ledger
    revenue_records: list[MonthlyRevenue]


def build_model(seed: int = 42) -> CascadeModel:
    """Build the complete Cascade Industries canonical model.

    Currently wires: revenue, COGS, and intercompany transactions.
    Additional generators (opex, PP&E, leases, tax, etc.) will be
    wired in by later beads — the model is extensible.
    """
    rng = random.Random(seed)
    ledger = Ledger()

    # ── Revenue & COGS ──────────────────────────────────────────────
    revenue_records = generate_monthly_revenue(rng)
    post_revenue_to_gl(ledger, revenue_records)

    # ── Intercompany ────────────────────────────────────────────────
    # Build monthly revenue totals for IC fee calculation
    totals: dict[tuple[str, int, int], Decimal] = {}
    for r in revenue_records:
        key = (r.entity_code, r.year, r.month)
        totals[key] = totals.get(key, Decimal(0)) + r.revenue

    ic_txns = generate_ic_transactions(totals)
    post_ic_loan_principal(ledger)
    post_ic_transactions_to_gl(ledger, ic_txns)

    return CascadeModel(
        ledger=ledger,
        revenue_records=revenue_records,
    )
