"""European partnership investment allocation data (TC-07-EU).

Generates 8 European transparent-entity partnership allocation statements for
Cascade Europe Holdings B.V. (CE).  Replaces TC-07's US Schedule K-1 concept
with jurisdiction-specific allocation statements from DE (KG, GmbH & Co. KG),
FR (SCI), LU (SCSp), NL (CV), and UK (LLP).

Key differences from TC-07:
- Heterogeneous formats per jurisdiction (not uniform K-1 boxes)
- Multilingual (DE/FR/NL/EN)
- EUR amounts (comma-decimal format) except Thames Valley (GBP)
- No Section 199A; replaced by Dutch participation exemption traps
- Amended statement uses German "KORRIGIERT" labelling
- WHT rates are jurisdiction/treaty-specific

Planted errors:
  ERR-EU-004: withholding_tax_rate_mismatch — Thames Valley 20% vs 15% treaty
  ERR-EU-005: partner_share_mismatch — Capital Croissance 4.8% vs 5.0%

Determinism: all values hardcoded, no RNG, no unordered sets.
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from enum import Enum


class AllocPdfStyle(Enum):
    """PDF rendering style for each allocation statement."""
    CLEAN_TEXT = "clean_text"      # reportlab-style clean text PDF
    SCANNED = "scanned"           # fpdf2 low-res scanned-style PDF


class AllocLanguage(Enum):
    """Language of the allocation statement."""
    GERMAN = "de"
    FRENCH = "fr"
    DUTCH = "nl"
    ENGLISH = "en"


@dataclass(frozen=True)
class AllocationCategory:
    """A single income/loss category within an allocation statement."""
    local_label: str          # Label in original language
    english_label: str        # Standardised English label
    amount_eur: Decimal       # Amount in EUR (already converted if GBP)
    amount_local: Decimal | None = None  # Amount in local currency (GBP only)


@dataclass(frozen=True)
class AmendmentDetail:
    """Tracks what changed in a corrected/amended statement."""
    field_changed: str
    original_label: str
    original_value: Decimal
    amended_label: str
    amended_value: Decimal
    description: str


@dataclass(frozen=True)
class EUPartnershipInvestment:
    """A single European partnership allocation statement."""

    alloc_id: str               # 1-indexed identifier for the investment
    fund_name: str
    jurisdiction: str           # DE, FR, LU, NL, UK
    legal_form: str             # KG, GmbH & Co. KG, SCI, SCSp, CV, LLP
    partner_share_pct: Decimal  # CE's share percentage
    fiscal_year: int
    language: AllocLanguage
    pdf_style: AllocPdfStyle
    is_amended: bool
    currency: str               # EUR or GBP

    # Allocation categories (income breakdown)
    categories: tuple[AllocationCategory, ...]

    # Withholding tax
    wht_amount_eur: Decimal
    wht_amount_local: Decimal | None  # GBP amount if currency is GBP
    wht_rate_pct: Decimal             # Effective WHT rate applied

    # Investment register data
    cost_basis_eur: Decimal
    carrying_value_eur: Decimal
    acquisition_date: str
    classification: str        # "IAS 28 associate" or "IFRS 9 financial asset"
    accounting_method: str     # "equity method" or "fair value through P&L"
    classification_note: str   # Note about classification rationale

    # Participation exemption eligibility
    participation_exempt: bool          # True if >=5% and potentially qualifies
    participation_exempt_note: str      # Qualification details/caveats

    # Amendment details
    amendments: tuple[AmendmentDetail, ...] = ()

    @property
    def total_income_eur(self) -> Decimal:
        """Sum of all allocation categories in EUR."""
        return sum(c.amount_eur for c in self.categories)

    @property
    def alloc_filename(self) -> str:
        """Return the PDF filename for this allocation statement."""
        # Map fund names to filesystem-safe names
        return _FILENAMES[self.alloc_id]


# ── GBP→EUR conversion rate ────────────────────────────────────────────────
GBP_EUR_RATE = Decimal("1.17")  # Average FY2025 rate per prompt


# ── Filename mapping ───────────────────────────────────────────────────────
_FILENAMES: dict[str, str] = {
    "ALLOC-001": "tc07eu_alloc_rheinland_kg.pdf",
    "ALLOC-002": "tc07eu_alloc_suedbayern_kgcokg.pdf",
    "ALLOC-003": "tc07eu_alloc_fonds_rhonealpes_sci.pdf",
    "ALLOC-004": "tc07eu_alloc_capital_croissance_slp.pdf",
    "ALLOC-005": "tc07eu_alloc_benelux_ventures_cv.pdf",
    "ALLOC-006": "tc07eu_alloc_thames_valley_llp.pdf",
    "ALLOC-007": "tc07eu_alloc_nordic_infra_scsp.pdf",
    "ALLOC-008": "tc07eu_alloc_beteiligungen_muenchen_kg_amended.pdf",
}


# ── Canary key helpers ─────────────────────────────────────────────────────

def alloc_canary_key(alloc_id: str) -> str:
    """Return the canary file key for an allocation statement PDF."""
    num = alloc_id.split("-")[1]
    return f"tc07eu_alloc_{num}"


INVESTMENT_REGISTER_KEY = "tc07eu_investment_register"
WHT_SUMMARY_KEY = "tc07eu_wht_summary"

ALL_CANARY_KEYS: list[str] = sorted(
    [alloc_canary_key(f"ALLOC-{i:03d}") for i in range(1, 9)]
    + [INVESTMENT_REGISTER_KEY, WHT_SUMMARY_KEY]
)


# ── WHT treaty rates ──────────────────────────────────────────────────────

@dataclass(frozen=True)
class WHTRate:
    """Withholding tax rate entry for the WHT summary schedule."""
    jurisdiction: str
    income_type: str
    domestic_rate_pct: Decimal
    treaty_rate_pct: Decimal
    treaty_reference: str


WHT_RATES: list[WHTRate] = [
    WHTRate("DE", "Interest", Decimal("26.375"), Decimal("0"), "NL-DE Treaty Art. 11"),
    WHTRate("DE", "Dividends", Decimal("26.375"), Decimal("15"), "NL-DE Treaty Art. 10"),
    WHTRate("DE", "Partnership profit", Decimal("0"), Decimal("0"), "Transparent entity — no WHT"),
    WHTRate("FR", "Property income", Decimal("25"), Decimal("15"), "NL-FR Treaty Art. 6"),
    WHTRate("FR", "Capital gains (immovable)", Decimal("25"), Decimal("25"), "NL-FR Treaty Art. 13"),
    WHTRate("FR", "Interest", Decimal("25"), Decimal("0"), "NL-FR Treaty Art. 11"),
    WHTRate("LU", "Partnership allocations", Decimal("0"), Decimal("0"), "Tax transparent — no WHT"),
    WHTRate("LU", "Dividends", Decimal("15"), Decimal("0"), "NL-LU Treaty Art. 10"),
    WHTRate("NL", "All domestic", Decimal("0"), Decimal("0"), "Domestic — no WHT"),
    WHTRate("UK", "Partnership profit shares", Decimal("0"), Decimal("0"), "No WHT on profit shares"),
    WHTRate("UK", "LLP property income (NRL)", Decimal("20"), Decimal("15"),
            "NL-UK Treaty Art. 6; NRL scheme overrides for property income"),
    WHTRate("UK", "Interest", Decimal("20"), Decimal("0"), "NL-UK Treaty Art. 11"),
]


# ── Investment templates ───────────────────────────────────────────────────

def _d(val: str) -> Decimal:
    return Decimal(val)


def generate_eu_investments() -> list[EUPartnershipInvestment]:
    """Generate the 8 European partnership investments deterministically.

    All values are fixed per design bead synth-data-eu.8.
    Returns a list sorted by alloc_id.
    """
    investments = [
        # ALLOC-001: Rheinland Industriepark KG (Germany, KG)
        EUPartnershipInvestment(
            alloc_id="ALLOC-001",
            fund_name="Rheinland Industriepark KG",
            jurisdiction="DE",
            legal_form="KG (Kommanditgesellschaft)",
            partner_share_pct=_d("8.0"),
            fiscal_year=2025,
            language=AllocLanguage.GERMAN,
            pdf_style=AllocPdfStyle.CLEAN_TEXT,
            is_amended=False,
            currency="EUR",
            categories=(
                AllocationCategory("Mieteinnahmen", "Rental income", _d("140000")),
                AllocationCategory("Zinserträge", "Interest income", _d("25000")),
                AllocationCategory("Veräußerungsgewinne", "Capital gains", _d("20000")),
            ),
            wht_amount_eur=_d("6875"),
            wht_amount_local=None,
            wht_rate_pct=_d("26.375"),
            cost_basis_eur=_d("1200000"),
            carrying_value_eur=_d("1450000"),
            acquisition_date="2019-07-01",
            classification="IAS 28 associate",
            accounting_method="equity method",
            classification_note="≥5% share with board representation; active business qualifies",
            participation_exempt=True,
            participation_exempt_note=(
                "8.0% share in active business KG qualifies for deelnemingsvrijstelling. "
                "Income exempt from Dutch CIT; associated WHT NOT creditable."
            ),
        ),
        # ALLOC-002: Südbayern Gewerbe GmbH & Co. KG (Germany)
        EUPartnershipInvestment(
            alloc_id="ALLOC-002",
            fund_name="Südbayern Gewerbe GmbH & Co. KG",
            jurisdiction="DE",
            legal_form="GmbH & Co. KG",
            partner_share_pct=_d("6.0"),
            fiscal_year=2025,
            language=AllocLanguage.GERMAN,
            pdf_style=AllocPdfStyle.CLEAN_TEXT,
            is_amended=False,
            currency="EUR",
            categories=(
                AllocationCategory("Mieteinnahmen", "Rental income", _d("78000")),
                AllocationCategory("Zinserträge", "Interest income", _d("14000")),
            ),
            wht_amount_eur=_d("3690"),
            wht_amount_local=None,
            wht_rate_pct=_d("26.375"),
            cost_basis_eur=_d("800000"),
            carrying_value_eur=_d("920000"),
            acquisition_date="2020-01-15",
            classification="IAS 28 associate",
            accounting_method="equity method",
            classification_note=(
                "≥5% share; however, holds primarily passive real estate — "
                "may fail asset test for participation exemption"
            ),
            participation_exempt=True,
            participation_exempt_note=(
                "6.0% share meets threshold, but GmbH & Co. KG holds primarily "
                "passive real estate — may fail the asset test, making exemption "
                "inapplicable and WHT creditable instead. Borderline case."
            ),
        ),
        # ALLOC-003: Fonds Immobilier Rhône-Alpes SCI (France)
        EUPartnershipInvestment(
            alloc_id="ALLOC-003",
            fund_name="Fonds Immobilier Rhône-Alpes SCI",
            jurisdiction="FR",
            legal_form="SCI (Société Civile Immobilière)",
            partner_share_pct=_d("3.5"),
            fiscal_year=2025,
            language=AllocLanguage.FRENCH,
            pdf_style=AllocPdfStyle.SCANNED,
            is_amended=False,
            currency="EUR",
            categories=(
                AllocationCategory("Revenus fonciers", "Property income", _d("58000")),
                AllocationCategory("Plus-values immobilières", "Capital gains", _d("10500")),
            ),
            wht_amount_eur=_d("8700"),
            wht_amount_local=None,
            wht_rate_pct=_d("15"),
            cost_basis_eur=_d("500000"),
            carrying_value_eur=_d("580000"),
            acquisition_date="2021-06-01",
            classification="IFRS 9 financial asset",
            accounting_method="fair value through P&L",
            classification_note="<5% share; no significant influence",
            participation_exempt=False,
            participation_exempt_note="3.5% share below 5% threshold — does not qualify.",
        ),
        # ALLOC-004: Capital Croissance SLP (Luxembourg SCSp)
        # ERR-EU-005: allocation statement says 4.8%, register says 5.0%
        EUPartnershipInvestment(
            alloc_id="ALLOC-004",
            fund_name="Capital Croissance SLP",
            jurisdiction="LU",
            legal_form="SCSp (Société en Commandite Spéciale)",
            partner_share_pct=_d("4.8"),  # Statement shows 4.8% (ERR-EU-005)
            fiscal_year=2025,
            language=AllocLanguage.ENGLISH,
            pdf_style=AllocPdfStyle.CLEAN_TEXT,
            is_amended=False,
            currency="EUR",
            categories=(
                AllocationCategory("Trading income", "Trading income", _d("1200000")),
                AllocationCategory("Carried interest", "Carried interest", _d("450000")),
                AllocationCategory("Dividend income", "Dividend income", _d("200000")),
            ),
            wht_amount_eur=_d("0"),
            wht_amount_local=None,
            wht_rate_pct=_d("0"),
            cost_basis_eur=_d("5000000"),
            carrying_value_eur=_d("6800000"),
            acquisition_date="2018-03-15",
            classification="IFRS 9 financial asset",
            accounting_method="fair value through P&L",
            classification_note=(
                "Statement shows 4.8% but register shows 5.0%. "
                "At 5.0% may qualify for participation exemption; at 4.8% does not."
            ),
            participation_exempt=False,
            participation_exempt_note=(
                "Boundary case: register says 5.0% (qualifies), "
                "statement says 4.8% (does not qualify). "
                "Agent must flag discrepancy and recommend verification."
            ),
        ),
        # ALLOC-005: Benelux Ventures CV (Netherlands)
        EUPartnershipInvestment(
            alloc_id="ALLOC-005",
            fund_name="Benelux Ventures CV",
            jurisdiction="NL",
            legal_form="CV (Commanditaire Vennootschap)",
            partner_share_pct=_d("12.0"),
            fiscal_year=2025,
            language=AllocLanguage.DUTCH,
            pdf_style=AllocPdfStyle.CLEAN_TEXT,
            is_amended=False,
            currency="EUR",
            categories=(
                AllocationCategory("Winstaandeel", "Trading income", _d("280000")),
                AllocationCategory("Rentebaten", "Interest income", _d("40000")),
            ),
            wht_amount_eur=_d("0"),
            wht_amount_local=None,
            wht_rate_pct=_d("0"),
            cost_basis_eur=_d("2000000"),
            carrying_value_eur=_d("2500000"),
            acquisition_date="2017-09-01",
            classification="IAS 28 associate",
            accounting_method="equity method",
            classification_note=(
                "12% share below 20% but classified as IAS 28 associate — "
                "significant influence via board seat"
            ),
            participation_exempt=True,
            participation_exempt_note=(
                "12.0% share qualifies for deelnemingsvrijstelling if IAS 28 "
                "classification stands. Agent should note board-seat justification "
                "for IAS 28 at <20% rather than silently reclassifying."
            ),
        ),
        # ALLOC-006: Thames Valley Property LLP (UK) — in GBP
        # ERR-EU-004: WHT at 20% NRL vs 15% treaty rate
        EUPartnershipInvestment(
            alloc_id="ALLOC-006",
            fund_name="Thames Valley Property LLP",
            jurisdiction="UK",
            legal_form="LLP (Limited Liability Partnership)",
            partner_share_pct=_d("4.0"),
            fiscal_year=2025,
            language=AllocLanguage.ENGLISH,
            pdf_style=AllocPdfStyle.SCANNED,
            is_amended=False,
            currency="GBP",
            categories=(
                AllocationCategory("Property income", "Property income",
                                   _d("165000") * GBP_EUR_RATE,
                                   amount_local=_d("165000")),
                AllocationCategory("Interest received", "Interest income",
                                   _d("27500") * GBP_EUR_RATE,
                                   amount_local=_d("27500")),
            ),
            wht_amount_eur=_d("38500") * GBP_EUR_RATE,
            wht_amount_local=_d("38500"),
            wht_rate_pct=_d("20"),  # NRL scheme rate — ERR-EU-004
            cost_basis_eur=_d("900000"),
            carrying_value_eur=_d("1050000"),
            acquisition_date="2022-04-01",
            classification="IFRS 9 financial asset",
            accounting_method="fair value through P&L",
            classification_note="<5% share; no significant influence",
            participation_exempt=False,
            participation_exempt_note="4.0% share below 5% threshold — does not qualify.",
        ),
        # ALLOC-007: Nordic Infrastructure SCSp II (Luxembourg)
        EUPartnershipInvestment(
            alloc_id="ALLOC-007",
            fund_name="Nordic Infrastructure SCSp II",
            jurisdiction="LU",
            legal_form="SCSp (Société en Commandite Spéciale)",
            partner_share_pct=_d("2.5"),
            fiscal_year=2025,
            language=AllocLanguage.ENGLISH,
            pdf_style=AllocPdfStyle.CLEAN_TEXT,
            is_amended=False,
            currency="EUR",
            categories=(
                AllocationCategory("Infrastructure income", "Infrastructure income", _d("120000")),
                AllocationCategory("Interest income", "Interest income", _d("25000")),
            ),
            wht_amount_eur=_d("0"),
            wht_amount_local=None,
            wht_rate_pct=_d("0"),
            cost_basis_eur=_d("600000"),
            carrying_value_eur=_d("720000"),
            acquisition_date="2023-01-15",
            classification="IFRS 9 financial asset",
            accounting_method="fair value through P&L",
            classification_note="<5% share; no significant influence",
            participation_exempt=False,
            participation_exempt_note="2.5% share below 5% threshold — does not qualify.",
        ),
        # ALLOC-008: Beteiligungen München KG (AMENDED/KORRIGIERT)
        EUPartnershipInvestment(
            alloc_id="ALLOC-008",
            fund_name="Beteiligungen München KG",
            jurisdiction="DE",
            legal_form="KG (Kommanditgesellschaft)",
            partner_share_pct=_d("5.5"),
            fiscal_year=2025,
            language=AllocLanguage.GERMAN,
            pdf_style=AllocPdfStyle.CLEAN_TEXT,
            is_amended=True,
            currency="EUR",
            categories=(
                # Amended: Gewinnanteil reduced from €340k to €285k
                AllocationCategory("Gewinnanteil", "Profit share", _d("285000")),
                # Added in amendment: Zinserträge €55k
                AllocationCategory("Zinserträge", "Interest income", _d("55000")),
            ),
            wht_amount_eur=_d("14506"),
            wht_amount_local=None,
            wht_rate_pct=_d("26.375"),
            cost_basis_eur=_d("1500000"),
            carrying_value_eur=_d("1800000"),
            acquisition_date="2016-11-01",
            classification="IAS 28 associate",
            accounting_method="equity method",
            classification_note="≥5% share with active business holding",
            participation_exempt=True,
            participation_exempt_note=(
                "5.5% share in active industrial holding KG qualifies for "
                "deelnemingsvrijstelling. Income exempt; WHT not creditable."
            ),
            amendments=(
                AmendmentDetail(
                    field_changed="income_category",
                    original_label="Gewinnanteil",
                    original_value=_d("340000"),
                    amended_label="Gewinnanteil",
                    amended_value=_d("285000"),
                    description=(
                        "Gewinnanteil (profit share) reduced from EUR 340,000 to EUR 285,000. "
                        "EUR 55,000 reclassified as Zinsertraege (interest income)."
                    ),
                ),
                AmendmentDetail(
                    field_changed="new_category",
                    original_label="Zinserträge",
                    original_value=_d("0"),
                    amended_label="Zinserträge",
                    amended_value=_d("55000"),
                    description=(
                        "New line: Zinsertraege (interest income) EUR 55,000 - "
                        "reclassified from original Gewinnanteil."
                    ),
                ),
            ),
        ),
    ]

    investments.sort(key=lambda x: x.alloc_id)
    return investments


# ── Consolidation helpers ──────────────────────────────────────────────────

# Standardised categories for consolidated schedule
STANDARD_CATEGORIES = [
    "Rental/property income",
    "Interest income",
    "Dividend income",
    "Capital gains",
    "Trading/business income",
    "Infrastructure income",
    "Carried interest",
    "Profit share",
]

# Map from English labels to standard categories
_CATEGORY_MAP: dict[str, str] = {
    "Rental income": "Rental/property income",
    "Property income": "Rental/property income",
    "Interest income": "Interest income",
    "Interest received": "Interest income",
    "Dividend income": "Dividend income",
    "Capital gains": "Capital gains",
    "Trading income": "Trading/business income",
    "Infrastructure income": "Infrastructure income",
    "Carried interest": "Carried interest",
    "Profit share": "Profit share",
}


def consolidated_totals_eu(
    investments: list[EUPartnershipInvestment],
) -> dict[str, Decimal]:
    """Consolidate all allocation data by standard category.

    Returns a dict of {standard_category: total_eur} for non-zero totals.
    """
    totals: dict[str, Decimal] = {}
    for inv in investments:
        for cat in inv.categories:
            std = _CATEGORY_MAP.get(cat.english_label, cat.english_label)
            totals[std] = totals.get(std, Decimal(0)) + cat.amount_eur
    # Sort by category name for determinism
    return dict(sorted(totals.items()))


def consolidated_by_jurisdiction(
    investments: list[EUPartnershipInvestment],
) -> dict[str, Decimal]:
    """Total allocated income by jurisdiction (EUR)."""
    by_jur: dict[str, Decimal] = {}
    for inv in investments:
        by_jur[inv.jurisdiction] = (
            by_jur.get(inv.jurisdiction, Decimal(0)) + inv.total_income_eur
        )
    return dict(sorted(by_jur.items()))


def total_wht(investments: list[EUPartnershipInvestment]) -> Decimal:
    """Grand total WHT across all investments (EUR)."""
    return sum(inv.wht_amount_eur for inv in investments)
