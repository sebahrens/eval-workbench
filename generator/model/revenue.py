"""Cascade Industries revenue & COGS generators (§2.4 of prompt.md).

Generates monthly revenue by product line × entity × month for FY2023–FY2025
with seasonal patterns (Q4 heavy, Q1 dip). Applies growth rates from config
(FY23→24 +6%, FY24→25 +9% with Advanced Materials driving growth). Posts
revenue and COGS journal entries to the GL via ``ledger.post()``.

Key constraints from the spec:
- Consolidated FY25 YoY growth must equal 9.2% (TC-03 gold value).
- One product line in Advanced Materials declines 4% (TC-03 detection).
- COGS correlates at entity gross margins: PC=35%, AM=52%, DS=18%.
- Monthly revenue follows seasonal weights from config (Q4 heavy).

Revenue computation strategy:
1. Fix entity FY2025 targets from entities.py ($95M, $65M, $40M).
2. Fix entity-level FY24→25 growth rates so consolidated = 9.2%.
   PC +5.0%, AM +18.0%, DS +6.5% → consolidated FY24 = $183.12M → 9.20%.
3. Fix FY23→24 growth at +6% uniform (spec §2.4 rule 6).
4. Within each entity, product-line shares define FY25 revenue.
   FY24 product-line revenue is derived by back-computing from the entity
   FY24 target, using per-product growth rates that satisfy the entity total.
"""

from __future__ import annotations

import datetime
import random
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal

from generator.model.entities import SUBSIDIARIES
from generator.model.gl import JournalEntry, JournalEntryLine, Ledger

# ── Entity-level growth rates ───────────────────────────────────────────────
# Calibrated so consolidated FY24→25 growth = 9.2%.
#
# PC FY25=$95M / 1.05  = $90,476,190 (FY24)
# AM FY25=$65M / 1.18  = $55,084,746 (FY24)
# DS FY25=$40M / 1.065 = $37,558,685 (FY24)
# Consolidated FY24    = $183,119,621
# Growth = $200M / $183,119,621 - 1 = 9.204% → rounds to 9.2% ✓

_ENTITY_GROWTH_FY24_TO_FY25: dict[str, Decimal] = {
    "PC": Decimal("0.05"),
    "AM": Decimal("0.18"),
    "DS": Decimal("0.065"),
}

_ENTITY_GROWTH_FY23_TO_FY24 = Decimal("0.06")  # Uniform per §2.4 rule 6


# ── Product line definitions ────────────────────────────────────────────────
# 6 product lines across 3 entities, per TC-03 requirement.

@dataclass(frozen=True)
class ProductLine:
    """A revenue-generating product line within a subsidiary."""

    name: str
    entity_code: str  # PC, AM, DS
    revenue_account: str  # GL account for revenue
    cogs_account: str  # GL account for COGS
    fy25_share: float  # Share of entity FY2025 revenue
    fy24_to_fy25_growth: float  # Product-level YoY growth FY24→FY25


# Product-line FY24 revenue is derived from FY25 / (1 + growth).
# The sum of product-line FY24 revenues within an entity must equal
# the entity FY24 target. We enforce this by setting one product line's
# FY24 revenue as the entity residual (see _build_annual_table).
#
# TC-03 constraints:
# - Two product lines drive growth: Advanced Composites (+45%), Warehousing (+10%)
# - One product line declines 4%: Specialty Coatings (-4%)

PRODUCT_LINES: tuple[ProductLine, ...] = (
    # Precision Components — 2 product lines (stable, ~5% growth each)
    ProductLine(
        name="Industrial Parts",
        entity_code="PC",
        revenue_account="4010",
        cogs_account="5010",
        fy25_share=0.70,
        fy24_to_fy25_growth=0.05,  # Matches entity growth
    ),
    ProductLine(
        name="Custom Machining",
        entity_code="PC",
        revenue_account="4020",
        cogs_account="5020",
        fy25_share=0.30,
        fy24_to_fy25_growth=0.05,  # Matches entity growth
    ),
    # Advanced Materials — 2 product lines (one growing, one declining)
    # AM entity growth = 18%. Specialty Coatings declines 4%.
    # Advanced Composites absorbs the remainder → very high growth.
    # SC FY24 = 29.25M / 0.96 = 30,468,750
    # AC FY24 = entity FY24 (55,084,746) - 30,468,750 = 24,615,996
    # AC growth = 35.75M / 24.616M - 1 ≈ 45.2% → growth driver
    ProductLine(
        name="Advanced Composites",
        entity_code="AM",
        revenue_account="4010",
        cogs_account="5010",
        fy25_share=0.55,
        fy24_to_fy25_growth=0.0,  # Placeholder — derived as entity residual
    ),
    ProductLine(
        name="Specialty Coatings",
        entity_code="AM",
        revenue_account="4020",
        cogs_account="5020",
        fy25_share=0.45,
        fy24_to_fy25_growth=-0.04,  # Declining product line (TC-03)
    ),
    # Distribution Services — 2 product lines
    # DS entity growth = 6.5%. Warehousing grows faster, Freight slower.
    # WH FY24 = 24M / 1.10 = 21,818,182
    # FR FY24 = entity FY24 (37,558,685) - 21,818,182 = 15,740,503
    # FR growth = 16M / 15.741M - 1 ≈ 1.65%
    ProductLine(
        name="Warehousing Services",
        entity_code="DS",
        revenue_account="4030",
        cogs_account="5030",
        fy25_share=0.60,
        fy24_to_fy25_growth=0.10,  # Second growth driver (TC-03)
    ),
    ProductLine(
        name="Freight & Logistics",
        entity_code="DS",
        revenue_account="4030",
        cogs_account="5040",
        fy25_share=0.40,
        fy24_to_fy25_growth=0.0,  # Placeholder — derived as entity residual
    ),
)

# Within each entity, one product line is the "residual" whose FY24 revenue
# is computed as: entity_FY24_total - sum(other_product_lines_FY24).
# This guarantees entity-level FY24 totals are exact.
_RESIDUAL_PRODUCT_LINES = {"Advanced Composites", "Freight & Logistics"}


# ── Monthly seasonal weights ────────────────────────────────────────────────

_MONTHLY_WEIGHTS: tuple[float, ...] = (
    # Q1: 0.20 total — Q1 dip
    0.060, 0.065, 0.075,
    # Q2: 0.25 total
    0.080, 0.085, 0.085,
    # Q3: 0.25 total
    0.085, 0.085, 0.080,
    # Q4: 0.30 total — Q4 heavy
    0.095, 0.100, 0.105,
)

assert abs(sum(_MONTHLY_WEIGHTS) - 1.0) < 1e-9, "Monthly weights must sum to 1.0"


# ── Annual revenue table ───────────────────────────────────────────────────

@dataclass
class MonthlyRevenue:
    """A single month's revenue for one product line."""

    year: int
    month: int  # 1-12
    entity_code: str
    product_line: str
    revenue: Decimal  # Positive, in cents-precision
    cogs: Decimal  # Positive, in cents-precision
    revenue_account: str
    cogs_account: str


def _build_annual_table() -> dict[tuple[str, int], Decimal]:
    """Build product-line annual revenue for all years.

    Returns {(product_line_name, year): annual_revenue}.
    """
    table: dict[tuple[str, int], Decimal] = {}

    # Group product lines by entity
    by_entity: dict[str, list[ProductLine]] = {}
    for pl in PRODUCT_LINES:
        by_entity.setdefault(pl.entity_code, []).append(pl)

    for entity_code, pls in sorted(by_entity.items()):
        entity_fy25 = Decimal(SUBSIDIARIES[entity_code].revenue_target)
        entity_fy24 = entity_fy25 / (1 + _ENTITY_GROWTH_FY24_TO_FY25[entity_code])

        # Identify the residual product line for this entity
        residual_pl = None
        non_residual_pls = []
        for pl in pls:
            if pl.name in _RESIDUAL_PRODUCT_LINES:
                residual_pl = pl
            else:
                non_residual_pls.append(pl)

        # FY2025: straightforward share allocation
        for pl in pls:
            table[(pl.name, 2025)] = (
                entity_fy25 * Decimal(str(pl.fy25_share))
            ).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

        # FY2024: non-residual PLs computed from their FY25 and growth rate.
        # Residual PL gets whatever is left to match entity FY24 total.
        non_residual_fy24_sum = Decimal(0)
        for pl in non_residual_pls:
            fy25 = table[(pl.name, 2025)]
            fy24 = (fy25 / (1 + Decimal(str(pl.fy24_to_fy25_growth)))).quantize(
                Decimal("0.01"), rounding=ROUND_HALF_UP
            )
            table[(pl.name, 2024)] = fy24
            non_residual_fy24_sum += fy24

        if residual_pl is not None:
            residual_fy24 = (entity_fy24 - non_residual_fy24_sum).quantize(
                Decimal("0.01"), rounding=ROUND_HALF_UP
            )
            table[(residual_pl.name, 2024)] = residual_fy24
        else:
            # No residual — all PLs have explicit growth rates
            # (PC has uniform growth, so both match entity total naturally)
            pass

        # FY2023: uniform 6% growth for all product lines within entity
        for pl in pls:
            fy24 = table[(pl.name, 2024)]
            fy23 = (fy24 / (1 + _ENTITY_GROWTH_FY23_TO_FY24)).quantize(
                Decimal("0.01"), rounding=ROUND_HALF_UP
            )
            table[(pl.name, 2023)] = fy23

    return table


# Pre-compute the table at module load (deterministic, no RNG needed)
_ANNUAL_TABLE = _build_annual_table()


def generate_monthly_revenue(
    rng: random.Random,
    years: list[int] | None = None,
) -> list[MonthlyRevenue]:
    """Generate monthly revenue and COGS for all product lines across all years.

    Uses seasonal monthly weights with small random perturbation for realism.
    The perturbation is seeded via ``rng`` for determinism.

    Returns a sorted list of MonthlyRevenue records (by year, month, entity, product).
    """
    if years is None:
        years = [2023, 2024, 2025]

    gross_margins: dict[str, Decimal] = {
        code: Decimal(str(entity.gross_margin))
        for code, entity in SUBSIDIARIES.items()
    }

    records: list[MonthlyRevenue] = []

    for pl in PRODUCT_LINES:
        margin = gross_margins[pl.entity_code]

        for year in sorted(years):
            annual = _ANNUAL_TABLE[(pl.name, year)]

            # Apply monthly weights with small random perturbation (±2%)
            raw_weights = []
            for w in _MONTHLY_WEIGHTS:
                perturbed = w * (1 + rng.uniform(-0.02, 0.02))
                raw_weights.append(perturbed)

            # Normalize so they sum to 1.0
            weight_sum = sum(raw_weights)
            monthly_weights = [w / weight_sum for w in raw_weights]

            # Allocate annual revenue to months
            allocated = Decimal(0)
            monthly_amounts: list[Decimal] = []
            for i, mw in enumerate(monthly_weights):
                if i == 11:
                    # Last month gets the remainder to ensure exact annual total
                    amount = annual - allocated
                else:
                    amount = (annual * Decimal(str(mw))).quantize(
                        Decimal("0.01"), rounding=ROUND_HALF_UP
                    )
                    allocated += amount
                monthly_amounts.append(amount)

            for month_idx, rev_amount in enumerate(monthly_amounts):
                month = month_idx + 1
                cogs_amount = (rev_amount * (1 - margin)).quantize(
                    Decimal("0.01"), rounding=ROUND_HALF_UP
                )
                records.append(
                    MonthlyRevenue(
                        year=year,
                        month=month,
                        entity_code=pl.entity_code,
                        product_line=pl.name,
                        revenue=rev_amount,
                        cogs=cogs_amount,
                        revenue_account=pl.revenue_account,
                        cogs_account=pl.cogs_account,
                    )
                )

    # Sort deterministically
    records.sort(key=lambda r: (r.year, r.month, r.entity_code, r.product_line))
    return records


def post_revenue_to_gl(
    ledger: Ledger,
    records: list[MonthlyRevenue],
) -> None:
    """Post all revenue and COGS records as journal entries to the ledger.

    Each month × product line generates two JEs:
    1. Revenue: DR Accounts Receivable (1100), CR Revenue account
    2. COGS: DR COGS account, CR Inventory (1220 for mfg, 2020 for services)
    """
    for rec in records:
        date = datetime.date(rec.year, rec.month, 15)  # Mid-month posting

        # Round to whole dollars for JE posting (§2.4 rule 7)
        rev_dollars = rec.revenue.quantize(Decimal("1"), rounding=ROUND_HALF_UP)
        cogs_dollars = rec.cogs.quantize(Decimal("1"), rounding=ROUND_HALF_UP)

        if rev_dollars <= 0:
            continue

        # Revenue JE: DR A/R, CR Revenue
        rev_entry = JournalEntry(
            date=date,
            entity_code=rec.entity_code,
            description=f"Revenue — {rec.product_line} {rec.year}-{rec.month:02d}",
            lines=(
                JournalEntryLine(
                    account="1100",
                    debit=rev_dollars,
                    credit=Decimal(0),
                    memo=f"{rec.product_line} sales",
                ),
                JournalEntryLine(
                    account=rec.revenue_account,
                    debit=Decimal(0),
                    credit=rev_dollars,
                    memo=f"{rec.product_line} sales",
                ),
            ),
        )
        ledger.post(rev_entry)

        if cogs_dollars <= 0:
            continue

        # COGS JE: DR COGS, CR Inventory (finished goods for mfg, accrued for services)
        cr_account = "1220" if rec.entity_code in ("PC", "AM") else "2020"

        cogs_entry = JournalEntry(
            date=date,
            entity_code=rec.entity_code,
            description=f"COGS — {rec.product_line} {rec.year}-{rec.month:02d}",
            lines=(
                JournalEntryLine(
                    account=rec.cogs_account,
                    debit=cogs_dollars,
                    credit=Decimal(0),
                    memo=f"{rec.product_line} cost of sales",
                ),
                JournalEntryLine(
                    account=cr_account,
                    debit=Decimal(0),
                    credit=cogs_dollars,
                    memo=f"{rec.product_line} cost of sales",
                ),
            ),
        )
        ledger.post(cogs_entry)


# ── Validation helpers ──────────────────────────────────────────────────────

def validate_consolidated_growth(records: list[MonthlyRevenue]) -> dict[str, Decimal]:
    """Compute and return consolidated YoY growth rates.

    Returns a dict like {"FY2024_growth": Decimal("0.060"), "FY2025_growth": Decimal("0.092")}.
    """
    by_year: dict[int, Decimal] = {}
    for rec in records:
        by_year[rec.year] = by_year.get(rec.year, Decimal(0)) + rec.revenue

    result = {}
    for year in sorted(by_year):
        prior = year - 1
        if prior in by_year:
            growth = (by_year[year] - by_year[prior]) / by_year[prior]
            result[f"FY{year}_growth"] = growth

    return result


def validate_product_line_growth(
    records: list[MonthlyRevenue],
) -> dict[str, Decimal]:
    """Compute FY24→FY25 growth by product line.

    Returns a dict of product_line_name → growth rate.
    """
    by_pl_year: dict[tuple[str, int], Decimal] = {}
    for rec in records:
        key = (rec.product_line, rec.year)
        by_pl_year[key] = by_pl_year.get(key, Decimal(0)) + rec.revenue

    result = {}
    for pl in PRODUCT_LINES:
        fy24 = by_pl_year.get((pl.name, 2024), Decimal(0))
        fy25 = by_pl_year.get((pl.name, 2025), Decimal(0))
        if fy24 > 0:
            result[pl.name] = (fy25 - fy24) / fy24

    return result
