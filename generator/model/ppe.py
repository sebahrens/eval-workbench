"""Cascade Industries fixed asset register and depreciation engine.

Generates a fixed asset register with book depreciation (straight-line) and
tax depreciation (MACRS). The book-tax depreciation difference is a key
temporary difference for TC-06 (ASC 740 tax provision).

Key constraints:
- Precision Components: high fixed assets (mature manufacturing).
- Advanced Materials: R&D equipment with shorter lives.
- Distribution Services: asset-light (mostly vehicles and warehouse equipment).
- Book depreciation uses straight-line; tax uses MACRS (5, 7, 15, 39-year).
- The difference creates a deferred tax liability for TC-06.
"""

from __future__ import annotations

import datetime
import random
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal

from generator.model.gl import JournalEntry, JournalEntryLine, Ledger

# ── MACRS tables (half-year convention) ──────────────────────────────────────
# Source: IRS Publication 946, Table A-1 through A-6.
# Each tuple gives annual depreciation percentages for the recovery period.

MACRS_TABLES: dict[int, tuple[Decimal, ...]] = {
    5: tuple(Decimal(x) for x in (
        "0.2000", "0.3200", "0.1920", "0.1152", "0.1152", "0.0576",
    )),
    7: tuple(Decimal(x) for x in (
        "0.1429", "0.2449", "0.1749", "0.1249", "0.0893",
        "0.0892", "0.0893", "0.0446",
    )),
    15: tuple(Decimal(x) for x in (
        "0.0500", "0.0950", "0.0855", "0.0770", "0.0693",
        "0.0623", "0.0590", "0.0590", "0.0591", "0.0590",
        "0.0591", "0.0590", "0.0591", "0.0590", "0.0591",
        "0.0295",
    )),
    39: tuple(
        # 39-year nonresidential real property (mid-month convention simplified
        # to half-year for our purposes). Year 1 and year 40 get half.
        [Decimal("0.01282")] +
        [Decimal("0.02564")] * 38 +
        [Decimal("0.01282")]
    ),
}


# ── Asset class definitions ──────────────────────────────────────────────────

@dataclass(frozen=True)
class AssetClass:
    """Defines a class of fixed assets with GL mapping and depreciation terms."""

    name: str
    asset_account: str       # 14xx GL account
    accum_depr_account: str  # 15xx GL account
    depr_expense_account: str  # 6210 (book) or 6330 (R&D equipment)
    book_life_years: int     # Straight-line useful life (book)
    macrs_life_years: int    # MACRS recovery period (tax)
    salvage_pct: float       # Salvage value as % of cost (book only; MACRS has no salvage)


ASSET_CLASSES: dict[str, AssetClass] = {
    "Buildings": AssetClass(
        "Buildings", "1410", "1500", "6210",
        book_life_years=30, macrs_life_years=39, salvage_pct=0.10,
    ),
    "Machinery": AssetClass(
        "Machinery", "1420", "1510", "6210",
        book_life_years=10, macrs_life_years=7, salvage_pct=0.05,
    ),
    "Vehicles": AssetClass(
        "Vehicles", "1430", "1520", "6210",
        book_life_years=5, macrs_life_years=5, salvage_pct=0.10,
    ),
    "Furniture": AssetClass(
        "Furniture", "1440", "1530", "6210",
        book_life_years=7, macrs_life_years=7, salvage_pct=0.00,
    ),
    "Computers": AssetClass(
        "Computers", "1450", "1540", "6210",
        book_life_years=3, macrs_life_years=5, salvage_pct=0.00,
    ),
    "R&D Equipment": AssetClass(
        "R&D Equipment", "1420", "1510", "6330",
        book_life_years=5, macrs_life_years=5, salvage_pct=0.05,
    ),
}


# ── Fixed asset record ───────────────────────────────────────────────────────

@dataclass
class FixedAsset:
    """A single fixed asset in the register."""

    asset_id: str
    entity_code: str
    description: str
    asset_class: str
    acquisition_date: datetime.date
    cost: Decimal
    salvage_value: Decimal
    book_life_years: int
    macrs_life_years: int
    asset_account: str
    accum_depr_account: str
    depr_expense_account: str
    disposed: bool = False
    disposal_date: datetime.date | None = None

    def book_depr_annual(self) -> Decimal:
        """Annual straight-line book depreciation."""
        depreciable = self.cost - self.salvage_value
        if self.book_life_years <= 0 or depreciable <= 0:
            return Decimal(0)
        return (depreciable / Decimal(self.book_life_years)).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP
        )

    def book_depr_for_year(self, year: int) -> Decimal:
        """Book depreciation for a specific fiscal year (prorated first/last year)."""
        if self.disposed and self.disposal_date and self.disposal_date.year < year:
            return Decimal(0)

        annual = self.book_depr_annual()
        if annual <= 0:
            return Decimal(0)

        acq_year = self.acquisition_date.year
        end_year = acq_year + self.book_life_years

        if year < acq_year or year > end_year:
            return Decimal(0)

        # First year: prorate from acquisition month
        if year == acq_year:
            months = 12 - self.acquisition_date.month + 1
            return (annual * Decimal(months) / Decimal(12)).quantize(
                Decimal("1"), rounding=ROUND_HALF_UP
            )

        # Last year: remaining depreciation
        if year == end_year:
            total_prior = sum(
                self.book_depr_for_year(y) for y in range(acq_year, year)
            )
            depreciable = self.cost - self.salvage_value
            remaining = depreciable - total_prior
            return max(Decimal(0), remaining)

        # If disposed mid-year, prorate
        if self.disposed and self.disposal_date and self.disposal_date.year == year:
            months = self.disposal_date.month
            return (annual * Decimal(months) / Decimal(12)).quantize(
                Decimal("1"), rounding=ROUND_HALF_UP
            )

        return annual

    def tax_depr_for_year(self, year: int) -> Decimal:
        """MACRS tax depreciation for a specific fiscal year."""
        if self.disposed and self.disposal_date and self.disposal_date.year < year:
            return Decimal(0)

        acq_year = self.acquisition_date.year
        recovery_year = year - acq_year  # 0-based index into MACRS table

        table = MACRS_TABLES.get(self.macrs_life_years)
        if table is None or recovery_year < 0 or recovery_year >= len(table):
            return Decimal(0)

        # MACRS depreciates full cost (no salvage value)
        return (self.cost * table[recovery_year]).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP
        )

    def cumulative_book_depr(self, through_year: int) -> Decimal:
        """Total accumulated book depreciation through end of given year."""
        acq_year = self.acquisition_date.year
        return sum(
            self.book_depr_for_year(y) for y in range(acq_year, through_year + 1)
        )

    def cumulative_tax_depr(self, through_year: int) -> Decimal:
        """Total accumulated MACRS depreciation through end of given year."""
        acq_year = self.acquisition_date.year
        return sum(
            self.tax_depr_for_year(y) for y in range(acq_year, through_year + 1)
        )


# ── Asset generation templates by entity ─────────────────────────────────────
# (description_prefix, asset_class, cost_low, cost_high, count)

_ASSET_TEMPLATES: dict[str, list[tuple[str, str, int, int, int]]] = {
    "PC": [
        ("CNC Machine", "Machinery", 150_000, 400_000, 12),
        ("Press Brake", "Machinery", 80_000, 200_000, 6),
        ("Delivery Van", "Vehicles", 35_000, 55_000, 4),
        ("Office Furniture", "Furniture", 5_000, 15_000, 8),
        ("Workstation", "Computers", 2_000, 5_000, 15),
        ("Factory Building Improvements", "Buildings", 500_000, 1_500_000, 2),
    ],
    "AM": [
        ("Lab Instrument", "R&D Equipment", 50_000, 250_000, 8),
        ("Materials Testing Rig", "R&D Equipment", 100_000, 350_000, 4),
        ("Production Line", "Machinery", 200_000, 600_000, 5),
        ("Delivery Vehicle", "Vehicles", 40_000, 60_000, 3),
        ("Server Cluster", "Computers", 20_000, 80_000, 4),
        ("Office Furniture", "Furniture", 5_000, 15_000, 6),
        ("Lab Renovation", "Buildings", 300_000, 800_000, 2),
    ],
    "DS": [
        ("Forklift", "Machinery", 25_000, 45_000, 6),
        ("Delivery Truck", "Vehicles", 50_000, 90_000, 8),
        ("Pallet Racking System", "Furniture", 15_000, 35_000, 5),
        ("Barcode Scanner System", "Computers", 3_000, 8_000, 10),
        ("Warehouse Lighting", "Buildings", 50_000, 150_000, 2),
    ],
}


def generate_fixed_assets(rng: random.Random) -> list[FixedAsset]:
    """Generate the complete fixed asset register deterministically.

    Assets are acquired across 2020–2025 (some predating our 3-year window
    so they have accumulated depreciation in FY2023).

    Returns a sorted list of FixedAsset records.
    """
    assets: list[FixedAsset] = []
    seq = 0

    for entity_code in sorted(_ASSET_TEMPLATES.keys()):
        templates = _ASSET_TEMPLATES[entity_code]
        for desc_prefix, class_name, cost_low, cost_high, count in templates:
            ac = ASSET_CLASSES[class_name]
            for i in range(count):
                seq += 1
                asset_id = f"FA-{seq:04d}"

                cost = Decimal(rng.randint(cost_low, cost_high))
                # Round to nearest $1000
                cost = (cost / 1000).quantize(Decimal("1"), rounding=ROUND_HALF_UP) * 1000

                salvage = (cost * Decimal(str(ac.salvage_pct))).quantize(
                    Decimal("1"), rounding=ROUND_HALF_UP
                )

                # Acquisition date: distributed across 2020-01-01 to 2025-06-30.
                # Older assets for mature entities, newer for growing ones.
                acq_start = datetime.date(2020, 1, 1)
                acq_end = datetime.date(2025, 6, 30)
                days_range = (acq_end - acq_start).days
                day_offset = rng.randint(0, days_range)
                acq_date = acq_start + datetime.timedelta(days=day_offset)

                description = f"{desc_prefix} #{i + 1}"

                assets.append(FixedAsset(
                    asset_id=asset_id,
                    entity_code=entity_code,
                    description=description,
                    asset_class=class_name,
                    acquisition_date=acq_date,
                    cost=cost,
                    salvage_value=salvage,
                    book_life_years=ac.book_life_years,
                    macrs_life_years=ac.macrs_life_years,
                    asset_account=ac.asset_account,
                    accum_depr_account=ac.accum_depr_account,
                    depr_expense_account=ac.depr_expense_account,
                ))

    assets.sort(key=lambda a: a.asset_id)
    return assets


# ── Depreciation summary views ───────────────────────────────────────────────

@dataclass(frozen=True)
class DepreciationSummary:
    """Annual depreciation summary for one entity."""

    entity_code: str
    year: int
    book_depreciation: Decimal
    tax_depreciation: Decimal

    @property
    def book_tax_difference(self) -> Decimal:
        """Temporary difference: tax depr - book depr.

        Positive = tax depr > book depr → deferred tax liability.
        Negative = tax depr < book depr → deferred tax asset.
        """
        return self.tax_depreciation - self.book_depreciation


def compute_depreciation_summary(
    assets: list[FixedAsset],
    years: list[int] | None = None,
) -> list[DepreciationSummary]:
    """Compute book vs. tax depreciation by entity and year.

    This is the queryable view referenced in the bead acceptance criteria
    and consumed by the TC-06 tax provision model.
    """
    if years is None:
        years = [2023, 2024, 2025]

    summaries: list[DepreciationSummary] = []

    # Group assets by entity
    by_entity: dict[str, list[FixedAsset]] = {}
    for asset in assets:
        by_entity.setdefault(asset.entity_code, []).append(asset)

    for entity_code in sorted(by_entity.keys()):
        entity_assets = by_entity[entity_code]
        for year in sorted(years):
            book_total = sum(a.book_depr_for_year(year) for a in entity_assets)
            tax_total = sum(a.tax_depr_for_year(year) for a in entity_assets)

            summaries.append(DepreciationSummary(
                entity_code=entity_code,
                year=year,
                book_depreciation=book_total,
                tax_depreciation=tax_total,
            ))

    return summaries


def cumulative_book_tax_difference(
    assets: list[FixedAsset],
    through_year: int,
    entity_code: str | None = None,
) -> Decimal:
    """Cumulative book-tax depreciation difference through end of year.

    Positive = cumulative tax depr exceeds cumulative book depr → DTL.
    Used by the tax provision model to compute deferred tax balances.
    """
    filtered = assets if entity_code is None else [
        a for a in assets if a.entity_code == entity_code
    ]
    return sum(
        a.cumulative_tax_depr(through_year) - a.cumulative_book_depr(through_year)
        for a in filtered
    )


# ── GL posting ────────────────────────────────────────────────────────────────

def post_depreciation_to_gl(
    ledger: Ledger,
    assets: list[FixedAsset],
    years: list[int] | None = None,
) -> None:
    """Post book depreciation expense to the GL (annual entries per asset).

    DR: Depreciation Expense (6210) or R&D Equipment Depreciation (6330)
    CR: Accumulated Depreciation (15xx)

    Tax depreciation is NOT posted to the GL — it only appears in the
    tax provision workpapers.
    """
    if years is None:
        years = [2023, 2024, 2025]

    for asset in assets:
        for year in sorted(years):
            book_depr = asset.book_depr_for_year(year)
            if book_depr <= 0:
                continue

            date = datetime.date(year, 12, 31)  # Year-end depreciation entry

            entry = JournalEntry(
                date=date,
                entity_code=asset.entity_code,
                description=f"Depreciation — {asset.description} FY{year}",
                lines=(
                    JournalEntryLine(
                        account=asset.depr_expense_account,
                        debit=book_depr,
                        credit=Decimal(0),
                        memo=f"{asset.asset_class} depreciation",
                    ),
                    JournalEntryLine(
                        account=asset.accum_depr_account,
                        debit=Decimal(0),
                        credit=book_depr,
                        memo=f"{asset.asset_class} accumulated depreciation",
                    ),
                ),
            )
            ledger.post(entry)


def post_asset_acquisitions_to_gl(
    ledger: Ledger,
    assets: list[FixedAsset],
) -> None:
    """Post asset acquisition entries to the GL.

    DR: PP&E asset account (14xx)
    CR: Cash (1010)
    """
    for asset in assets:
        if asset.cost <= 0:
            continue

        entry = JournalEntry(
            date=asset.acquisition_date,
            entity_code=asset.entity_code,
            description=f"Asset acquisition — {asset.description}",
            lines=(
                JournalEntryLine(
                    account=asset.asset_account,
                    debit=asset.cost,
                    credit=Decimal(0),
                    memo=f"Acquire {asset.description}",
                ),
                JournalEntryLine(
                    account="1010",
                    debit=Decimal(0),
                    credit=asset.cost,
                    memo=f"Acquire {asset.description}",
                ),
            ),
        )
        ledger.post(entry)
