"""Model-level scenario invariant validation (synth-data-ln9.6).

Run after build_model() to verify cross-cutting invariants that
individual modules cannot check in isolation:

1. Entity/headcount consistency — employees match expected entities,
   total headcount aligns with config.
2. GL accounting checks — every posted journal entry is balanced,
   trial balance debits equal credits.
3. Consolidated revenue tolerance — FY2025 revenue is within
   $198M–$202M (±1% of $200M target).
4. Intercompany netting — all 9xxx accounts sum to zero on
   consolidation for every fiscal year.
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass, field
from decimal import Decimal
from typing import TYPE_CHECKING

from generator.model.consolidation import (
    build_income_statement,
    verify_ic_elimination,
)

if TYPE_CHECKING:
    from generator.config import Config
    from generator.model.build import CascadeModel


@dataclass
class ValidationError:
    """A single invariant violation."""

    category: str  # e.g. "headcount", "gl_balance", "revenue", "ic_netting"
    message: str
    detail: dict[str, object] = field(default_factory=dict)


def validate_model(
    model: CascadeModel,
    config: Config | None = None,
) -> list[ValidationError]:
    """Run all model-level invariant checks.

    Returns an empty list when all invariants hold.  Each violation
    is returned as a :class:`ValidationError` with a category, message,
    and optional detail dict.
    """
    errors: list[ValidationError] = []
    errors.extend(_check_entity_headcount(model, config))
    errors.extend(_check_gl_balance(model))
    errors.extend(_check_consolidated_revenue(model))
    errors.extend(_check_ic_netting(model))
    return errors


# ── 1. Entity / headcount consistency ────────────────────────────────────────


def _check_entity_headcount(
    model: CascadeModel,
    config: Config | None = None,
) -> list[ValidationError]:
    errors: list[ValidationError] = []

    valid_entity_codes = set(model.entities)

    # Every employee must belong to a known entity
    for emp in model.employees:
        if emp.entity_code not in valid_entity_codes:
            errors.append(ValidationError(
                category="headcount",
                message=f"Employee {emp.employee_id} has unknown entity '{emp.entity_code}'",
                detail={"employee_id": emp.employee_id, "entity_code": emp.entity_code},
            ))

    # Total headcount must match config if provided
    if config is not None:
        expected_total = config.company.employees.total_count
        actual_total = len(model.employees)
        if actual_total != expected_total:
            errors.append(ValidationError(
                category="headcount",
                message=(
                    f"Employee count mismatch: config expects {expected_total}, "
                    f"model has {actual_total}"
                ),
                detail={"expected": expected_total, "actual": actual_total},
            ))

    # Revenue-generating entities must have at least one employee
    for code, entity in model.subsidiaries.items():
        entity_emps = [e for e in model.employees if e.entity_code == code]
        if not entity_emps:
            errors.append(ValidationError(
                category="headcount",
                message=f"Subsidiary '{code}' ({entity.name}) has zero employees",
                detail={"entity_code": code},
            ))

    return errors


# ── 2. GL accounting checks ─────────────────────────────────────────────────


def _check_gl_balance(model: CascadeModel) -> list[ValidationError]:
    errors: list[ValidationError] = []

    # Every journal entry must be balanced (debits == credits)
    for i, entry in enumerate(model.ledger.entries):
        if not entry.is_balanced():
            errors.append(ValidationError(
                category="gl_balance",
                message=(
                    f"Unbalanced JE #{i}: debits={entry.total_debits()}, "
                    f"credits={entry.total_credits()}, desc='{entry.description}'"
                ),
                detail={
                    "index": i,
                    "debits": str(entry.total_debits()),
                    "credits": str(entry.total_credits()),
                },
            ))

    # Consolidated trial balance: sum of all debits must equal sum of all credits
    # across all accounts (i.e. net of all account balances should be zero).
    for year in [2023, 2024, 2025]:
        eoy = datetime.date(year, 12, 31)
        all_balances = {}
        for entity_code in model.entities:
            entity_bals = model.ledger.balance_by_account(entity_code, as_of_date=eoy)
            for acct, bal in entity_bals.items():
                all_balances[acct] = all_balances.get(acct, Decimal(0)) + bal
        net = sum(all_balances.values(), Decimal(0))
        if net != Decimal(0):
            errors.append(ValidationError(
                category="gl_balance",
                message=f"Trial balance does not net to zero for FY{year}: net={net}",
                detail={"year": year, "net_imbalance": str(net)},
            ))

    return errors


# ── 3. Consolidated revenue tolerance ────────────────────────────────────────


def _check_consolidated_revenue(model: CascadeModel) -> list[ValidationError]:
    errors: list[ValidationError] = []

    # FY2025 consolidated revenue must be in $198M–$202M range
    is_2025 = build_income_statement(model.ledger, 2025)
    rev = is_2025.total_revenue
    lower = Decimal("198_000_000")
    upper = Decimal("202_000_000")

    if not (lower <= rev <= upper):
        errors.append(ValidationError(
            category="revenue",
            message=(
                f"FY2025 consolidated revenue ${rev:,.0f} outside "
                f"tolerance range ${lower:,.0f}–${upper:,.0f}"
            ),
            detail={
                "revenue": str(rev),
                "lower_bound": str(lower),
                "upper_bound": str(upper),
            },
        ))

    # Revenue records must sum consistently with GL-derived revenue
    revenue_from_records = sum(
        r.revenue for r in model.revenue_records if r.year == 2025
    )
    # Allow small rounding difference (records are cents-precision, GL is whole dollars)
    diff = abs(rev - revenue_from_records)
    # Each product-line-month gets rounded to whole dollars, so max rounding
    # error is $1 per entry.  6 product lines × 12 months = 72 entries max.
    if diff > Decimal("100"):
        errors.append(ValidationError(
            category="revenue",
            message=(
                f"FY2025 revenue mismatch: GL shows ${rev:,.0f}, "
                f"records sum to ${revenue_from_records:,.2f} (diff=${diff:,.2f})"
            ),
            detail={
                "gl_revenue": str(rev),
                "records_revenue": str(revenue_from_records),
                "difference": str(diff),
            },
        ))

    return errors


# ── 4. Intercompany netting ──────────────────────────────────────────────────


def _check_ic_netting(model: CascadeModel) -> list[ValidationError]:
    errors: list[ValidationError] = []

    for year in [2023, 2024, 2025]:
        eoy = datetime.date(year, 12, 31)
        is_zero, total_imbalance, ic_detail = verify_ic_elimination(
            model.ledger, eoy,
        )
        if not is_zero:
            # Show the individual non-zero IC accounts for diagnostics
            nonzero = {
                acct: str(bal) for acct, bal in ic_detail.items() if bal != 0
            }
            errors.append(ValidationError(
                category="ic_netting",
                message=(
                    f"IC accounts do not net to zero for FY{year}: "
                    f"total imbalance={total_imbalance}"
                ),
                detail={
                    "year": year,
                    "total_imbalance": str(total_imbalance),
                    "nonzero_accounts": nonzero,
                },
            ))

    return errors
