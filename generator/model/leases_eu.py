"""European (IFRS 16) lease portfolio for TC-04-EU.

Generates 15 leases across the Cascade Europe group with:
- IFRS 16 single lessee model (no operating/finance classification)
- European entities (CE, CP, CM), EUR currency, European lessors
- DD.MM.YYYY / DD/MM/YYYY date formats, comma-decimal amounts
- Short-term exemption (2 leases), low-value asset exemption (1 lease)
- 3 amendments, 3 scanned-style PDFs
- EURIBOR-based IBR schedule

Determinism: uses only the passed ``rng``; no unordered sets or wall-clock reads.
"""

from __future__ import annotations

import datetime
import random
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal
from enum import Enum

# ── Enums ───────────────────────────────────────────────────────────────────

class EscalationTypeEU(Enum):
    FIXED_PCT = "fixed_pct"
    HICP = "hicp"          # Harmonised Index of Consumer Prices (EU CPI)
    STEPPED = "stepped"
    NONE = "none"


# ── Dataclasses ─────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class LeaseClauseEU:
    """Source-document reference for a single lease clause (EU version)."""

    clause_type: str           # e.g. "premises", "commencement", "rent", etc.
    section_label: str         # e.g. "Artikel II, Abschnitt 2.1"
    page: int
    summary: str


@dataclass(frozen=True)
class AmendmentEU:
    """A material amendment to a European lease."""

    effective_date: datetime.date
    description: str
    new_monthly_rent: Decimal | None = None
    new_term_months: int | None = None
    new_escalation_pct: Decimal | None = None
    clauses: tuple[LeaseClauseEU, ...] = ()


@dataclass(frozen=True)
class LeaseEU:
    """A single lease in the Cascade Europe IFRS 16 portfolio."""

    lease_id: str              # LS-EU-001 through LS-EU-015
    entity_code: str           # CE, CP, CM
    lessee: str
    lessor: str
    description: str
    commencement_date: datetime.date
    term_months: int
    monthly_base_rent: Decimal  # EUR
    escalation_type: EscalationTypeEU
    escalation_pct: Decimal
    escalation_steps: tuple[Decimal, ...] | None
    renewal_option_months: int
    renewal_rent_increase_pct: Decimal
    purchase_option: bool
    purchase_option_price: Decimal | None
    termination_provision: str
    short_term_exempt: bool
    low_value_exempt: bool     # IFRS 16-specific: asset value ≤ ~EUR 4,500
    asset_value_when_new: Decimal | None  # For low-value test (None if N/A)
    amendments: tuple[AmendmentEU, ...]

    # IFRS 16 computations
    discount_rate: Decimal     # EURIBOR-based IBR
    rou_asset_initial: Decimal
    lease_liability_initial: Decimal

    clauses: tuple[LeaseClauseEU, ...] = ()

    @property
    def effective_monthly_rent(self) -> Decimal:
        for amend in reversed(self.amendments):
            if amend.new_monthly_rent is not None:
                return amend.new_monthly_rent
        return self.monthly_base_rent

    @property
    def effective_term_months(self) -> int:
        for amend in reversed(self.amendments):
            if amend.new_term_months is not None:
                return amend.new_term_months
        return self.term_months

    @property
    def end_date(self) -> datetime.date:
        months = self.effective_term_months
        year = self.commencement_date.year + (self.commencement_date.month - 1 + months) // 12
        month = (self.commencement_date.month - 1 + months) % 12 + 1
        if month == 12:
            return datetime.date(year, 12, 31)
        return datetime.date(year, month, 1) - datetime.timedelta(days=1)

    @property
    def is_on_balance_sheet(self) -> bool:
        """IFRS 16: all leases go on balance sheet except short-term and low-value exempt."""
        return not self.short_term_exempt and not self.low_value_exempt


# ── EURIBOR-based IBR schedule ──────────────────────────────────────────────

_IBR_EU_BY_TERM: list[tuple[int, Decimal]] = [
    (24, Decimal("0.0375")),    # ≤2yr
    (60, Decimal("0.0425")),    # ≤5yr
    (120, Decimal("0.0475")),   # ≤10yr
    (999, Decimal("0.0525")),   # >10yr
]


def _ibr_eu_for_term(months: int) -> Decimal:
    for threshold, rate in _IBR_EU_BY_TERM:
        if months <= threshold:
            return rate
    return Decimal("0.0525")


# ── PV of lease payments ────────────────────────────────────────────────────

def _pv_lease_payments(
    monthly_rent: Decimal,
    term_months: int,
    annual_rate: Decimal,
) -> Decimal:
    if term_months <= 0:
        return Decimal(0)
    r = annual_rate / Decimal(12)
    if r == 0:
        return monthly_rent * term_months
    factor = (Decimal(1) - (Decimal(1) + r) ** (-term_months)) / r
    return (monthly_rent * factor).quantize(Decimal("1"), rounding=ROUND_HALF_UP)


# ── European entity names ──────────────────────────────────────────────────

_ENTITY_NAMES_EU: dict[str, str] = {
    "CE": "Cascade Europe Holdings B.V.",
    "CP": "Cascade Praezisionsteile GmbH",
    "CM": "Cascade Materiaux Avances SAS",
}


def _entity_legal_name_eu(code: str) -> str:
    return _ENTITY_NAMES_EU[code]


# ── Lease templates ────────────────────────────────────────────────────────

_LEASE_TEMPLATES_EU: list[dict] = [
    # ── CE (Amsterdam) — 2 leases ────────────────────────────────────
    dict(
        entity_code="CE",
        lessor="Vastgoed Beheer Amsterdam B.V.",
        description="Head office — Zuidas business district, Amsterdam",
        term_months=120, monthly_rent=42_000,
        escalation=EscalationTypeEU.FIXED_PCT, esc_pct="0.025",
        renewal_months=60, purchase=False, lease_type_note="office",
        short_term=False, low_value=False, asset_value_new=None,
        # Amendment: early termination clause exercised, then reversed
        amendments=[(
            48, 0, 0,
            "Early termination clause exercised effective 01.01.2024; subsequently "
            "reversed by mutual agreement — original lease term reinstated with "
            "rent escalation applied from the reinstatement date",
        )],
    ),
    dict(
        entity_code="CE",
        lessor="Parkeergarage Centraal Amsterdam N.V.",
        description="Parking facility — 40 spaces, underground garage, Zuidas",
        term_months=60, monthly_rent=6_800,
        escalation=EscalationTypeEU.NONE, esc_pct="0",
        renewal_months=24, purchase=False, lease_type_note="parking",
        short_term=False, low_value=False, asset_value_new=None,
        amendments=[],
    ),
    # ── CP (Munich) — 7 leases ───────────────────────────────────────
    dict(
        entity_code="CP",
        lessor="Immobiliengruppe Suedbayern GmbH",
        description="Manufacturing facility — Halle A, Industriepark Muenchen-Ost",
        term_months=120, monthly_rent=48_000,
        escalation=EscalationTypeEU.FIXED_PCT, esc_pct="0.030",
        renewal_months=60, purchase=False, lease_type_note="manufacturing",
        short_term=False, low_value=False, asset_value_new=None,
        # Amendment: rent increase + term extension (post-COVID renegotiation)
        amendments=[(
            36, 8_000, 24,
            "Post-COVID renegotiation: expanded production area by 600 m2; "
            "monthly rent increased and lease term extended by 24 months",
        )],
    ),
    dict(
        entity_code="CP",
        lessor="Bayerische Lagerhaus Verwaltung GmbH",
        description="Warehouse — Logistikzentrum Muenchen-Nord",
        term_months=84, monthly_rent=22_000,
        escalation=EscalationTypeEU.HICP, esc_pct="0",
        renewal_months=36, purchase=False, lease_type_note="warehouse",
        short_term=False, low_value=False, asset_value_new=None,
        amendments=[],
    ),
    dict(
        entity_code="CP",
        lessor="Maschinenleasing Bayern AG",
        description="CNC precision milling centre — DMG MORI DMU 80 eVo",
        term_months=60, monthly_rent=9_200,
        escalation=EscalationTypeEU.NONE, esc_pct="0",
        renewal_months=0, purchase=True, purchase_price=92_000,
        lease_type_note="equipment",
        short_term=False, low_value=False, asset_value_new=None,
        amendments=[],
    ),
    dict(
        entity_code="CP",
        lessor="Gewerbepark Schwabing Verwaltungs GmbH",
        description="Office annexe — 2nd floor, Schwanthalerstrasse 45, Munich",
        term_months=36, monthly_rent=8_500,
        escalation=EscalationTypeEU.STEPPED, esc_pct="0",
        stepped_rents=(Decimal("8500"), Decimal("8900"), Decimal("9300")),
        renewal_months=24, purchase=False, lease_type_note="office",
        short_term=False, low_value=False, asset_value_new=None,
        amendments=[],
    ),
    dict(
        entity_code="CP",
        lessor="Bueromaschinen Direkt GmbH",
        description="Office printer — Konica Minolta bizhub C300i",
        term_months=36, monthly_rent=180,
        escalation=EscalationTypeEU.NONE, esc_pct="0",
        renewal_months=0, purchase=False, lease_type_note="low_value_equipment",
        # LOW-VALUE EXEMPT: asset value when new EUR 3,800 (< EUR 4,500 threshold)
        short_term=False, low_value=True, asset_value_new=Decimal("3800"),
        amendments=[],
    ),
    dict(
        entity_code="CP",
        lessor="Flottenloesung Sueddeutschland GmbH",
        description="Vehicle fleet — 4 delivery vans, Mercedes Sprinter",
        term_months=10, monthly_rent=4_600,
        escalation=EscalationTypeEU.NONE, esc_pct="0",
        renewal_months=0, purchase=False, lease_type_note="vehicles",
        # SHORT-TERM EXEMPT: 10 months remaining, no purchase option
        short_term=True, low_value=False, asset_value_new=None,
        amendments=[],
    ),
    dict(
        entity_code="CP",
        lessor="Maschinenleasing Bayern AG",
        description="Surface grinding machine — JUNKER Jumat 6S",
        term_months=48, monthly_rent=5_800,
        escalation=EscalationTypeEU.NONE, esc_pct="0",
        renewal_months=0, purchase=True, purchase_price=65_000,
        lease_type_note="equipment",
        short_term=False, low_value=False, asset_value_new=None,
        amendments=[],
    ),
    # ── CM (Lyon) — 6 leases ────────────────────────────────────────
    dict(
        entity_code="CM",
        lessor="Societe Fonciere Rhone-Alpes SAS",
        description="R&D laboratory — Batiment 7, Parc Technologique de Lyon",
        term_months=96, monthly_rent=35_000,
        escalation=EscalationTypeEU.HICP, esc_pct="0",
        renewal_months=48, purchase=False, lease_type_note="laboratory",
        short_term=False, low_value=False, asset_value_new=None,
        # Amendment: scope change — additional floor added
        amendments=[(
            36, 12_000, 0,
            "Extension de bail: ajout du 3eme etage (450 m2) pour salle blanche "
            "supplementaire; loyer mensuel augmente pour refleter la surface "
            "additionnelle",
        )],
    ),
    dict(
        entity_code="CM",
        lessor="Immobiliere du Parc de la Part-Dieu SAS",
        description="Office space — 4th floor, Tour Oxygene, Lyon Part-Dieu",
        term_months=60, monthly_rent=18_000,
        escalation=EscalationTypeEU.FIXED_PCT, esc_pct="0.020",
        renewal_months=24, purchase=False, lease_type_note="office",
        short_term=False, low_value=False, asset_value_new=None,
        amendments=[],
    ),
    dict(
        entity_code="CM",
        lessor="Societe Fonciere Rhone-Alpes SAS",
        description="Clean room — ISO Class 6, Batiment 7, zone B",
        term_months=120, monthly_rent=28_000,
        escalation=EscalationTypeEU.FIXED_PCT, esc_pct="0.035",
        renewal_months=60, purchase=False, lease_type_note="cleanroom",
        short_term=False, low_value=False, asset_value_new=None,
        amendments=[],
    ),
    dict(
        entity_code="CM",
        lessor="Location Equipements Industriels SAS",
        description="Spectrometer — Bruker S8 TIGER Series 2",
        term_months=48, monthly_rent=7_500,
        escalation=EscalationTypeEU.NONE, esc_pct="0",
        renewal_months=0, purchase=True, purchase_price=185_000,
        lease_type_note="equipment",
        short_term=False, low_value=False, asset_value_new=None,
        amendments=[],
    ),
    dict(
        entity_code="CM",
        lessor="Entrepots Lyon Confluence SARL",
        description="Storage facility — 800 m2, Zone Industrielle Gerland",
        term_months=60, monthly_rent=9_500,
        escalation=EscalationTypeEU.FIXED_PCT, esc_pct="0.020",
        renewal_months=0, purchase=False, lease_type_note="storage",
        short_term=False, low_value=False, asset_value_new=None,
        amendments=[],
    ),
    dict(
        entity_code="CM",
        lessor="SolutionsIT France SAS",
        description="IT server rack — 2 cabinets, colocation Lyon",
        term_months=8, monthly_rent=2_200,
        escalation=EscalationTypeEU.NONE, esc_pct="0",
        renewal_months=0, purchase=False, lease_type_note="it_equipment",
        # SHORT-TERM EXEMPT: 8 months remaining, return clause
        short_term=True, low_value=False, asset_value_new=None,
        amendments=[],
    ),
]


# ── Lease generation ────────────────────────────────────────────────────────

def _commencement_date_eu(rng: random.Random, term_months: int) -> datetime.date:
    """Deterministic commencement date for European leases.

    Same logic as US model but with slightly different date ranges.
    """
    if term_months <= 12:
        month = rng.randint(3, 6)
        return datetime.date(2025, month, 1)
    elif term_months <= 60:
        year = rng.randint(2021, 2023)
        month = rng.randint(1, 12)
        return datetime.date(year, month, 1)
    else:
        year = rng.randint(2019, 2022)
        month = rng.randint(1, 12)
        return datetime.date(year, month, 1)


def _build_clauses_eu(
    tmpl: dict,
    commence: datetime.date,
    term_months: int,
    monthly_rent: Decimal,
    esc_pct: Decimal,
    renewal_months: int,
    purchase_price: Decimal | None,
    termination: str,
) -> tuple[LeaseClauseEU, ...]:
    """Build source-document clause references for a European lease."""
    clauses: list[LeaseClauseEU] = []
    page = 1

    clauses.append(LeaseClauseEU(
        clause_type="premises",
        section_label="Artikel I, Abschnitt 1.1",
        page=page,
        summary=tmpl["description"],
    ))

    page += 1
    clauses.append(LeaseClauseEU(
        clause_type="commencement",
        section_label="Artikel II, Abschnitt 2.1",
        page=page,
        summary=f"Commencement: {commence.strftime('%d.%m.%Y')}",
    ))

    clauses.append(LeaseClauseEU(
        clause_type="term",
        section_label="Artikel II, Abschnitt 2.2",
        page=page,
        summary=f"Initial term: {term_months} months",
    ))

    page += 1
    rent_str = f"EUR {int(monthly_rent):,}".replace(",", ".")
    clauses.append(LeaseClauseEU(
        clause_type="rent",
        section_label="Artikel III, Abschnitt 3.1",
        page=page,
        summary=f"Base rent: {rent_str}/month",
    ))

    esc_type = tmpl["escalation"]
    if esc_type != EscalationTypeEU.NONE:
        page += 1
        if esc_type == EscalationTypeEU.FIXED_PCT:
            esc_summary = f"Fixed annual escalation of {float(esc_pct) * 100:.1f}%"
        elif esc_type == EscalationTypeEU.HICP:
            esc_summary = "Annual adjustment based on Eurozone HICP"
        elif esc_type == EscalationTypeEU.STEPPED:
            steps = tmpl.get("stepped_rents", ())
            step_strs = [f"EUR {int(s):,}".replace(",", ".") for s in steps]
            esc_summary = f"Stepped rent: {', '.join(step_strs)} per year"
        else:
            esc_summary = str(esc_type.value)
        clauses.append(LeaseClauseEU(
            clause_type="escalation",
            section_label="Artikel IV, Abschnitt 4.1",
            page=page,
            summary=esc_summary,
        ))

    if renewal_months > 0:
        page += 1
        clauses.append(LeaseClauseEU(
            clause_type="renewal",
            section_label="Artikel V, Abschnitt 5.1",
            page=page,
            summary=f"Renewal option: {renewal_months} months",
        ))

    if tmpl.get("purchase", False) and purchase_price is not None:
        page += 1
        clauses.append(LeaseClauseEU(
            clause_type="purchase_option",
            section_label="Artikel VI, Abschnitt 6.1",
            page=page,
            summary=f"Purchase option: EUR {int(purchase_price):,}".replace(",", "."),
        ))

    page += 1
    clauses.append(LeaseClauseEU(
        clause_type="termination",
        section_label="Artikel VII, Abschnitt 7.1",
        page=page,
        summary=termination,
    ))

    return tuple(clauses)


def _build_amendment_clauses_eu(
    amend_idx: int,
    amend_desc: str,
    new_rent: Decimal | None,
    new_term: int | None,
) -> tuple[LeaseClauseEU, ...]:
    """Build clause references for a European lease amendment."""
    clauses: list[LeaseClauseEU] = []
    prefix = f"Nachtrag Nr. {amend_idx}"

    if new_rent is not None:
        clauses.append(LeaseClauseEU(
            clause_type="rent",
            section_label=f"{prefix}, Abschnitt 1 -- Angepasste Miete",
            page=0,
            summary=f"Amended rent: EUR {int(new_rent):,}/month".replace(",", "."),
        ))

    if new_term is not None:
        clauses.append(LeaseClauseEU(
            clause_type="term",
            section_label=f"{prefix}, Abschnitt 2 -- Angepasste Laufzeit",
            page=0,
            summary=f"Amended term: {new_term} months",
        ))

    if not clauses:
        clauses.append(LeaseClauseEU(
            clause_type="premises",
            section_label=f"{prefix}, Abschnitt 1",
            page=0,
            summary=amend_desc,
        ))

    return tuple(clauses)


def generate_leases_eu(rng: random.Random) -> list[LeaseEU]:
    """Generate the 15-lease European portfolio deterministically.

    Returns a list sorted by lease_id.
    """
    leases: list[LeaseEU] = []

    for idx, tmpl in enumerate(_LEASE_TEMPLATES_EU, start=1):
        lease_id = f"LS-EU-{idx:03d}"
        entity_code = tmpl["entity_code"]
        term_months = tmpl["term_months"]

        commence = _commencement_date_eu(rng, term_months)
        monthly_rent = Decimal(str(tmpl["monthly_rent"]))
        esc_pct = Decimal(tmpl["esc_pct"])

        # Build amendments
        amendments: list[AmendmentEU] = []
        for amend_idx, amend_spec in enumerate(tmpl.get("amendments", []), start=1):
            months_after, rent_delta, term_delta, desc = amend_spec
            eff_date = datetime.date(
                commence.year + (commence.month - 1 + months_after) // 12,
                (commence.month - 1 + months_after) % 12 + 1,
                1,
            )
            new_rent = (monthly_rent + Decimal(str(rent_delta))) if rent_delta else None
            new_term = (term_months + term_delta) if term_delta else None
            amend_clauses = _build_amendment_clauses_eu(amend_idx, desc, new_rent, new_term)
            amendments.append(AmendmentEU(
                effective_date=eff_date,
                description=desc,
                new_monthly_rent=new_rent,
                new_term_months=new_term,
                clauses=amend_clauses,
            ))

        # Effective values
        eff_rent = monthly_rent
        eff_term = term_months
        for amend in amendments:
            if amend.new_monthly_rent is not None:
                eff_rent = amend.new_monthly_rent
            if amend.new_term_months is not None:
                eff_term = amend.new_term_months

        # Discount rate (EURIBOR-based)
        discount_rate = _ibr_eu_for_term(eff_term)

        # IFRS 16: all leases get ROU/liability EXCEPT short-term and low-value exempt
        is_short_term = tmpl.get("short_term", False)
        is_low_value = tmpl.get("low_value", False)
        asset_value_new = tmpl.get("asset_value_new")
        if asset_value_new is not None and not isinstance(asset_value_new, Decimal):
            asset_value_new = Decimal(str(asset_value_new))

        if is_short_term or is_low_value:
            rou_initial = Decimal(0)
            liability_initial = Decimal(0)
        else:
            liability_initial = _pv_lease_payments(eff_rent, eff_term, discount_rate)
            rou_initial = liability_initial

        # Stepped rents
        stepped = tmpl.get("stepped_rents")

        # Purchase option
        has_purchase = tmpl.get("purchase", False)
        purchase_price = Decimal(str(tmpl["purchase_price"])) if has_purchase else None

        # Termination provision
        if term_months <= 12:
            termination = "Lease expires at term end; no early termination."
        elif has_purchase:
            termination = "Lessee may terminate upon exercise of purchase option or at term end."
        else:
            notice_days = rng.choice([90, 120, 180])
            penalty_months = rng.choice([3, 6])
            termination = f"Early termination with {notice_days}-day notice and {penalty_months}-month penalty."

        # Renewal rent increase
        renewal_months = tmpl.get("renewal_months", 0)
        if renewal_months > 0:
            renewal_increase = Decimal(str(rng.choice(["0.03", "0.04", "0.05"])))
        else:
            renewal_increase = Decimal(0)

        # Build clause references
        lease_clauses = _build_clauses_eu(
            tmpl, commence, term_months, monthly_rent, esc_pct,
            renewal_months, purchase_price, termination,
        )

        lease = LeaseEU(
            lease_id=lease_id,
            entity_code=entity_code,
            lessee=_entity_legal_name_eu(entity_code),
            lessor=tmpl["lessor"],
            description=tmpl["description"],
            commencement_date=commence,
            term_months=term_months,
            monthly_base_rent=monthly_rent,
            escalation_type=tmpl["escalation"],
            escalation_pct=esc_pct,
            escalation_steps=tuple(stepped) if stepped else None,
            renewal_option_months=renewal_months,
            renewal_rent_increase_pct=renewal_increase,
            purchase_option=has_purchase,
            purchase_option_price=purchase_price,
            termination_provision=termination,
            short_term_exempt=is_short_term,
            low_value_exempt=is_low_value,
            asset_value_when_new=asset_value_new,
            amendments=tuple(amendments),
            discount_rate=discount_rate,
            rou_asset_initial=rou_initial,
            lease_liability_initial=liability_initial,
            clauses=lease_clauses,
        )
        leases.append(lease)

    leases.sort(key=lambda ls: ls.lease_id)
    return leases
