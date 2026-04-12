"""European OECD Transfer Pricing model for TC-09-EU.

Generates deterministic intercompany transaction data, comparable companies,
and interest rate benchmarks for the Cascade Europe Holdings B.V. group.

Five intercompany flows:
1. CP→CM raw materials at cost-plus-6%
2. CP→CD finished goods at cost-plus-8%
3. CE management fees at 1.5% of subsidiary revenue
4. CE→CM intercompany loan (€3M at 4.5%)
5. CM→CP R&D royalty at 3% of CP revenue

All data is deterministic (no RNG) per the design bead synth-data-eu.12.
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal

# ── Entity constants ────────────────────────────────────────────────────────

ENTITY_NAMES = {
    "CE": "Cascade Europe Holdings B.V.",
    "CP": "Cascade Präzisionsteile GmbH",
    "CM": "Cascade Matériaux Avancés SAS",
    "CD": "Cascade Distribution Services Ltd",
}

ENTITY_JURISDICTIONS = {
    "CE": "Netherlands",
    "CP": "Germany",
    "CM": "France",
    "CD": "United Kingdom",
}

# Revenue targets (EUR) — from EU company profile
_REVENUE = {
    "CP": Decimal("45000000"),
    "CM": Decimal("32000000"),
    "CD": Decimal("21000000"),  # £18M at GBP/EUR 1.17
    "CE": Decimal("22000000"),  # Holding — mgmt fees + loan interest
}

# ── Intercompany pricing constants ──────────────────────────────────────────

RAW_MATERIALS_MARKUP_EU = Decimal("0.06")   # CP→CM cost-plus-6%
FINISHED_GOODS_MARKUP_EU = Decimal("0.08")  # CP→CD cost-plus-8%
MANAGEMENT_FEE_PCT_EU = Decimal("0.015")    # CE→subs 1.5%
IC_LOAN_PRINCIPAL_EU = Decimal("3000000")   # €3M
IC_LOAN_RATE_EU = Decimal("0.045")          # 4.5% annual
ROYALTY_RATE_EU = Decimal("0.03")           # CM→CP 3% of CP revenue

# ── Intercompany transaction data class ─────────────────────────────────────


@dataclass(frozen=True)
class ICTransactionEU:
    """A single European intercompany transaction."""

    date: datetime.date
    from_entity: str
    to_entity: str
    transaction_type: str  # goods, services, royalty, interest, management_fee
    description: str
    volume_or_principal: Decimal
    price_or_rate: Decimal
    total_amount_eur: Decimal
    invoicing_currency: str  # EUR or GBP
    arm_length_method: str  # TNMM, Cost Plus, CUP, CUT


# ── Monthly revenue by entity for FY2025 ────────────────────────────────────
# Deterministic monthly allocation: annual / 12, with residual in December.


def _monthly_revenue(entity: str) -> list[tuple[int, Decimal]]:
    """Return (month, revenue) pairs for FY2025."""
    annual = _REVENUE[entity]
    monthly = (annual / 12).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    result = []
    running = Decimal(0)
    for m in range(1, 13):
        if m == 12:
            amt = annual - running
        else:
            amt = monthly
        result.append((m, amt))
        running += amt
    return result


# ── Generate intercompany transactions ──────────────────────────────────────


def generate_ic_transactions_eu() -> list[ICTransactionEU]:
    """Generate all ~120 FY2025 intercompany transactions for TC-09-EU."""
    txns: list[ICTransactionEU] = []

    cp_monthly = _monthly_revenue("CP")
    cm_monthly = _monthly_revenue("CM")
    cd_monthly = _monthly_revenue("CD")

    # 1. CP→CM raw materials: ~25% of CM COGS at cost-plus-6%
    # CM gross margin 52%, so COGS = 48% of revenue
    cm_cogs_pct = Decimal("0.48")
    ic_share = Decimal("0.25")
    for month, cm_rev in cm_monthly:
        base_cost = (cm_rev * cm_cogs_pct * ic_share).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP,
        )
        ic_amount = (base_cost * (1 + RAW_MATERIALS_MARKUP_EU)).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP,
        )
        txns.append(ICTransactionEU(
            date=datetime.date(2025, month, 15),
            from_entity="CP",
            to_entity="CM",
            transaction_type="goods",
            description=f"Raw materials — precision components batch RM-2025-{month:02d}",
            volume_or_principal=base_cost,
            price_or_rate=RAW_MATERIALS_MARKUP_EU,
            total_amount_eur=ic_amount,
            invoicing_currency="EUR",
            arm_length_method="Cost Plus",
        ))

    # 2. CP→CD finished goods: ~30% of CD COGS at cost-plus-8%
    # CD gross margin 18%, so COGS = 82% of revenue
    cd_cogs_pct = Decimal("0.82")
    cd_ic_share = Decimal("0.30")
    for month, cd_rev in cd_monthly:
        base_cost = (cd_rev * cd_cogs_pct * cd_ic_share).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP,
        )
        ic_amount = (base_cost * (1 + FINISHED_GOODS_MARKUP_EU)).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP,
        )
        txns.append(ICTransactionEU(
            date=datetime.date(2025, month, 20),
            from_entity="CP",
            to_entity="CD",
            transaction_type="goods",
            description=f"Finished goods — industrial assemblies FG-2025-{month:02d}",
            volume_or_principal=base_cost,
            price_or_rate=FINISHED_GOODS_MARKUP_EU,
            total_amount_eur=ic_amount,
            invoicing_currency="EUR",
            arm_length_method="Cost Plus",
        ))

    # 3. CE management fees: 1.5% of each subsidiary revenue
    for sub_code, sub_monthly in [("CP", cp_monthly), ("CM", cm_monthly), ("CD", cd_monthly)]:
        for month, sub_rev in sub_monthly:
            fee = (sub_rev * MANAGEMENT_FEE_PCT_EU).quantize(
                Decimal("1"), rounding=ROUND_HALF_UP,
            )
            txns.append(ICTransactionEU(
                date=datetime.date(2025, month, 28),
                from_entity="CE",
                to_entity=sub_code,
                transaction_type="management_fee",
                description=f"Management services — strategic oversight, treasury, legal {month:02d}/2025",
                volume_or_principal=sub_rev,
                price_or_rate=MANAGEMENT_FEE_PCT_EU,
                total_amount_eur=fee,
                invoicing_currency="EUR",
                arm_length_method="Cost Plus",
            ))

    # 4. CE→CM intercompany loan interest: €3M at 4.5%
    monthly_interest = (IC_LOAN_PRINCIPAL_EU * IC_LOAN_RATE_EU / 12).quantize(
        Decimal("1"), rounding=ROUND_HALF_UP,
    )
    for month in range(1, 13):
        txns.append(ICTransactionEU(
            date=datetime.date(2025, month, 28),
            from_entity="CE",
            to_entity="CM",
            transaction_type="interest",
            description=f"Intercompany loan interest — €3M facility {month:02d}/2025",
            volume_or_principal=IC_LOAN_PRINCIPAL_EU,
            price_or_rate=IC_LOAN_RATE_EU,
            total_amount_eur=monthly_interest,
            invoicing_currency="EUR",
            arm_length_method="CUP",
        ))

    # 5. CM→CP R&D royalty: 3% of CP revenue
    # Description deliberately says "R&D royalty — technology license" which is
    # ambiguous (the agent must verify direction via functional analysis).
    for month, cp_rev in cp_monthly:
        royalty = (cp_rev * ROYALTY_RATE_EU).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP,
        )
        txns.append(ICTransactionEU(
            date=datetime.date(2025, month, 25),
            from_entity="CP",
            to_entity="CM",
            transaction_type="royalty",
            description=f"R&D royalty — technology license {month:02d}/2025",
            volume_or_principal=cp_rev,
            price_or_rate=ROYALTY_RATE_EU,
            total_amount_eur=royalty,
            invoicing_currency="EUR",
            arm_length_method="CUT",
        ))

    return sorted(txns, key=lambda t: (t.date, t.transaction_type, t.from_entity))


# ── Comparable companies ────────────────────────────────────────────────────
# Manufacturing comparables: 15 companies, 3 rejected
# Operating margin = oper_income / revenue
#
# Design gold: Q1=3.8%, median=5.6%, Q3=7.9%
# n=12: Q1 at 0.25*(11)=2.75, Q3 at 0.75*(11)=8.25
#
# Sorted margins: [1.5, 2.0, 2.0, 4.4, 5.0, 5.2, 6.0, 6.4, 7.2, 10.0, 11.5, 13.0]
# Q1 = m[2]+0.75*(m[3]-m[2]) = 2.0+0.75*2.4 = 3.8  ✓
# Median = (m[5]+m[6])/2 = (5.2+6.0)/2 = 5.6  ✓
# Q3 = m[8]+0.25*(m[9]-m[8]) = 7.2+0.25*2.8 = 7.9  ✓
#
# All oper_income values are integers so rev-cogs-opex = oper_income exactly.
# (name, country, sic, revenue_M, cogs_M, opex_M, oper_income_M, total_assets_M,
#  roce_pct, is_rejected, rejection_reason)

MFG_COMPARABLES: list[tuple[str, str, str, int, int, int, int, int, float, bool, str]] = [
    # 12 accepted — calibrated for IQR 3.8%–7.9%, median 5.6%
    ("Alpen Maschinenbau AG", "DE", "2899", 200, 135, 62, 3, 220, 1.4, False, ""),        # 1.5%
    ("Borealis Industrietechnik GmbH", "DE", "2899", 200, 133, 63, 4, 230, 1.7, False, ""),  # 2.0%
    ("Ceravision Matériaux SA", "FR", "2899", 250, 170, 75, 5, 275, 1.8, False, ""),       # 2.0%
    ("Duisburg Präzision AG", "DE", "2899", 250, 168, 71, 11, 280, 3.9, False, ""),        # 4.4%
    ("EuroFab Industries NV", "NL", "2899", 200, 130, 60, 10, 220, 4.5, False, ""),        # 5.0%
    ("Fjordstål Komponenter AS", "NO", "2899", 250, 165, 72, 13, 275, 4.7, False, ""),     # 5.2%
    ("Gallium Precision Oy", "FI", "2899", 200, 128, 60, 12, 220, 5.5, False, ""),         # 6.0%
    ("Helvetia Werkstoff AG", "CH", "2899", 250, 162, 72, 16, 275, 5.8, False, ""),        # 6.4%
    ("Industriale Componenti SpA", "IT", "2899", 250, 160, 72, 18, 275, 6.5, False, ""),   # 7.2%
    ("Krakow Manufacturing SA", "PL", "2899", 200, 125, 55, 20, 220, 9.1, False, ""),      # 10.0%
    ("Lynx Teknik A/S", "DK", "2899", 200, 120, 57, 23, 220, 10.5, False, ""),             # 11.5%
    ("Meridian Fertigungstechnik GmbH", "AT", "2899", 200, 115, 59, 26, 220, 11.8, False, ""),  # 13.0%
    # 3 rejected companies
    ("Nordic Logistics Holdings AB", "SE", "4731", 300, 240, 42, 18, 180, 10.0, True,
     "SIC code mismatch (4731 Freight Transportation vs 2899 Manufacturing)"),
    ("Ostrava Heavy Industries a.s.", "CZ", "2899", 85, 72, 28, -15, 95, -15.8, True,
     "Financial distress — negative equity and operating losses for 2 consecutive years"),
    ("Rhineland Grosskonzern AG", "DE", "2899", 4500, 3100, 1150, 250, 5200, 4.8, True,
     "Revenue >10x tested party CP (size outlier per OECD Guidelines ¶3.43-3.46)"),
]

# Distribution comparables: 10 companies, 2 rejected
# Net margin = net_income / revenue
#
# Design gold: Q1=1.2%, median=2.1%, Q3=3.5%
# n=8: Q1 at 0.25*7=1.75, Q3 at 0.75*7=5.25
#
# Sorted margins: [0.8, 1.2, 1.2, 2.0, 2.2, 3.2, 4.4, 5.0]
# Q1 = m[1]+0.75*(m[2]-m[1]) = 1.2+0.75*0 = 1.2  ✓
# Median = (m[3]+m[4])/2 = (2.0+2.2)/2 = 2.1  ✓
# Q3 = m[5]+0.25*(m[6]-m[5]) = 3.2+0.25*1.2 = 3.5  ✓
#
# (name, country, sic, revenue_M, cogs_M, opex_M, net_income_M, total_assets_M,
#  is_rejected, rejection_reason)
# net_income_M uses Decimal for sub-million precision in net margins.

DIST_COMPARABLES: list[tuple[str, str, str, int, int, int, Decimal, int, bool, str]] = [
    # 8 accepted — calibrated for IQR 1.2%–3.5%, median 2.1%
    ("Athena Distribution BV", "NL", "5000", 95, 82, 12, Decimal("0.76"), 55, False, ""),     # 0.8%
    ("Baltic Freight Services UAB", "LT", "5000", 75, 64, 10, Decimal("0.90"), 42, False, ""),  # 1.2%
    ("Channel Logistics Ltd", "UK", "5000", 110, 94, 15, Decimal("1.32"), 65, False, ""),      # 1.2%
    ("Delta Warehousing GmbH", "DE", "5000", 130, 112, 15, Decimal("2.60"), 78, False, ""),    # 2.0%
    ("Europa Supply Chain SA", "FR", "5000", 85, 73, 10, Decimal("1.87"), 50, False, ""),      # 2.2%
    ("Frigate Logistics AB", "SE", "5000", 100, 84, 13, Decimal("3.20"), 58, False, ""),       # 3.2%
    ("Genova Distribuzione Srl", "IT", "5000", 70, 58, 9, Decimal("3.08"), 40, False, ""),     # 4.4%
    ("Harbour Trading BV", "NL", "5000", 60, 50, 7, Decimal("3.00"), 35, False, ""),           # 5.0%
    # 2 rejected
    ("InterGlobal Captive Logistics SA", "LU", "5000", 45, 39, 5, Decimal("1.35"), 28, True,
     "Captive entity of listed parent group — no independent pricing (related-party concentration >80%)"),
    ("Jupiter Restructuring Services GmbH", "DE", "5000", 120, 105, 18, Decimal("-3.60"), 72, True,
     "Restructuring losses — non-recurring charges distort margins (loss-making due to warehouse closure)"),
]

# ── Interest rate benchmarks ────────────────────────────────────────────────
# EURIBOR 12M quarterly for FY2024-FY2025 + BBB credit spreads
# ERR-EU-009: Q3 FY2025 has 0.38% instead of 3.80% (decimal point error)


@dataclass(frozen=True)
class EURIBOREntry:
    """One quarterly EURIBOR observation."""

    period: str       # e.g. "Q1 FY2024"
    year: int
    quarter: int
    tenor: str        # "3M", "6M", "12M"
    rate_pct: Decimal  # Annual rate as percentage (e.g. 3.80 = 3.80%)


@dataclass(frozen=True)
class CreditSpreadEntry:
    """BBB industrial credit spread observation."""

    period: str
    year: int
    quarter: int
    spread_bps: int   # Basis points (e.g. 120 = 1.20%)


def _euribor_data() -> list[EURIBOREntry]:
    """EURIBOR rates for FY2024-FY2025, quarterly, all three tenors."""
    # Rates designed so 12M average FY2025 ≈ 3.8% (with the error, avg drops)
    raw = [
        # FY2024
        ("Q1 FY2024", 2024, 1, "3M", "3.50"),
        ("Q1 FY2024", 2024, 1, "6M", "3.60"),
        ("Q1 FY2024", 2024, 1, "12M", "3.70"),
        ("Q2 FY2024", 2024, 2, "3M", "3.55"),
        ("Q2 FY2024", 2024, 2, "6M", "3.65"),
        ("Q2 FY2024", 2024, 2, "12M", "3.75"),
        ("Q3 FY2024", 2024, 3, "3M", "3.60"),
        ("Q3 FY2024", 2024, 3, "6M", "3.70"),
        ("Q3 FY2024", 2024, 3, "12M", "3.80"),
        ("Q4 FY2024", 2024, 4, "3M", "3.65"),
        ("Q4 FY2024", 2024, 4, "6M", "3.75"),
        ("Q4 FY2024", 2024, 4, "12M", "3.85"),
        # FY2025
        ("Q1 FY2025", 2025, 1, "3M", "3.60"),
        ("Q1 FY2025", 2025, 1, "6M", "3.70"),
        ("Q1 FY2025", 2025, 1, "12M", "3.75"),
        ("Q2 FY2025", 2025, 2, "3M", "3.65"),
        ("Q2 FY2025", 2025, 2, "6M", "3.75"),
        ("Q2 FY2025", 2025, 2, "12M", "3.80"),
        ("Q3 FY2025", 2025, 3, "3M", "3.62"),
        ("Q3 FY2025", 2025, 3, "6M", "3.72"),
        ("Q3 FY2025", 2025, 3, "12M", "3.85"),  # CORRECT value — error version is 0.38%
        ("Q4 FY2025", 2025, 4, "3M", "3.70"),
        ("Q4 FY2025", 2025, 4, "6M", "3.80"),
        ("Q4 FY2025", 2025, 4, "12M", "3.90"),
    ]
    return [
        EURIBOREntry(period=r[0], year=r[1], quarter=r[2], tenor=r[3],
                     rate_pct=Decimal(r[4]))
        for r in raw
    ]


# The errored value for Q3 FY2025 12M EURIBOR
EURIBOR_ERR_EU_009_CORRECT = Decimal("3.85")
EURIBOR_ERR_EU_009_WRONG = Decimal("0.38")  # decimal point error: 3.85 → 0.38


def _credit_spread_data() -> list[CreditSpreadEntry]:
    """BBB-rated European industrial borrower credit spreads."""
    raw = [
        ("Q1 FY2024", 2024, 1, 130),
        ("Q2 FY2024", 2024, 2, 125),
        ("Q3 FY2024", 2024, 3, 120),
        ("Q4 FY2024", 2024, 4, 115),
        ("Q1 FY2025", 2025, 1, 110),
        ("Q2 FY2025", 2025, 2, 105),
        ("Q3 FY2025", 2025, 3, 100),
        ("Q4 FY2025", 2025, 4, 110),
    ]
    return [
        CreditSpreadEntry(period=r[0], year=r[1], quarter=r[2], spread_bps=r[3])
        for r in raw
    ]


def generate_euribor_data() -> list[EURIBOREntry]:
    """Return EURIBOR data (correct values — error applied in formatter)."""
    return _euribor_data()


def generate_credit_spread_data() -> list[CreditSpreadEntry]:
    """Return credit spread data."""
    return _credit_spread_data()


# ── Canary keys ─────────────────────────────────────────────────────────────

ALL_CANARY_KEYS_TC09EU: list[str] = sorted([
    "tc09eu_intercompany_transactions",
    "tc09eu_comparable_companies",
    "tc09eu_master_file_fy2024",
    "tc09eu_local_file_cp_fy2024",
    "tc09eu_interest_rate_benchmarks",
])


# ── Helper: compute financial metrics ───────────────────────────────────────

def compute_mfg_iqr() -> dict[str, object]:
    """Compute manufacturing comparable IQR (operating margin)."""
    accepted = [(p[6] / p[3]) for p in MFG_COMPARABLES if not p[9] and p[3] > 0]
    margins = sorted(float(m) * 100 for m in accepted)
    n = len(margins)

    q1_idx = 0.25 * (n - 1)
    q3_idx = 0.75 * (n - 1)

    q1_low = int(q1_idx)
    q1_frac = q1_idx - q1_low
    q1 = margins[q1_low] + q1_frac * (margins[q1_low + 1] - margins[q1_low])

    q3_low = int(q3_idx)
    q3_frac = q3_idx - q3_low
    q3 = margins[q3_low] + q3_frac * (margins[q3_low + 1] - margins[q3_low])

    mid = n // 2
    if n % 2 == 0:
        median = (margins[mid - 1] + margins[mid]) / 2
    else:
        median = margins[mid]

    return {
        "accepted_count": n,
        "rejected_count": 3,
        "margins_sorted_pct": [round(m, 1) for m in margins],
        "q1_pct": round(q1, 1),
        "median_pct": round(median, 1),
        "q3_pct": round(q3, 1),
        "min_pct": round(margins[0], 1),
        "max_pct": round(margins[-1], 1),
    }


def compute_dist_iqr() -> dict[str, object]:
    """Compute distribution comparable IQR (net margin)."""
    accepted = [(p[6] / p[3]) for p in DIST_COMPARABLES if not p[8] and p[3] > 0]
    margins = sorted(float(m) * 100 for m in accepted)
    n = len(margins)

    q1_idx = 0.25 * (n - 1)
    q3_idx = 0.75 * (n - 1)

    q1_low = int(q1_idx)
    q1_frac = q1_idx - q1_low
    q1 = margins[q1_low] + q1_frac * (margins[q1_low + 1] - margins[q1_low])

    q3_low = int(q3_idx)
    q3_frac = q3_idx - q3_low
    q3 = margins[q3_low] + q3_frac * (margins[q3_low + 1] - margins[q3_low])

    mid = n // 2
    if n % 2 == 0:
        median = (margins[mid - 1] + margins[mid]) / 2
    else:
        median = margins[mid]

    return {
        "accepted_count": n,
        "rejected_count": 2,
        "margins_sorted_pct": [round(m, 1) for m in margins],
        "q1_pct": round(q1, 1),
        "median_pct": round(median, 1),
        "q3_pct": round(q3, 1),
        "min_pct": round(margins[0], 1),
        "max_pct": round(margins[-1], 1),
    }
