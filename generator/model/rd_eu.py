"""European R&D incentive model for TC-08-EU.

Generates deterministic R&D data for two European regimes:
- France CIR (Crédit d'Impôt Recherche) for CM (Lyon): 30% credit
- Germany Forschungszulage for CP (Munich): 25% of personnel costs

Key differences from US TC-08 (Section 41 / ASC method):
- No multi-year averaging formula (CIR applies 30% to current year)
- CIR forfaitaire 43% overhead on researcher salaries
- Public research organism invoices get 2× multiplier in CIR
- Forschungszulage: personnel costs ONLY (no supplies, no subcontractors)
- Forschungszulage €2M assessment basis cap (→ max €500k benefit)
- Frascati Manual criteria instead of US 4-part test

Planted error: ERR-EU-008 — maintenance contract miscoded as R&D supply at CP.

Determinism: all values hardcoded, no RNG, no unordered sets.
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal

# ── Constants ─────────────────────────────────────────────────────────────────

FY = 2025
WEEKS_IN_YEAR = 52
FY_START = datetime.date(FY, 1, 1)

# CIR parameters
CIR_RATE = Decimal("0.30")  # 30% for first €100M
CIR_FORFAITAIRE = Decimal("0.43")  # 43% overhead on researcher salaries
CIR_REDUCED_RATE = Decimal("0.05")  # 5% above €100M (not triggered here)
CIR_THRESHOLD = Decimal("100_000_000")

# Forschungszulage parameters
FZ_RATE = Decimal("0.25")  # 25% of eligible personnel costs
FZ_MAX_BASIS = Decimal("2_000_000")  # Assessment basis cap
FZ_MAX_BENEFIT = Decimal("500_000")  # Maximum benefit

# Prior year R&D spend (contextual only — NO averaging formula)
PRIOR_YEAR_RD_SPEND = {
    "CM": {2023: Decimal("3_000_000"), 2024: Decimal("3_300_000")},
    "CP": {2023: Decimal("1_100_000"), 2024: Decimal("1_200_000")},
}


# ── R&D Project definitions ──────────────────────────────────────────────────

@dataclass(frozen=True)
class RDProjectEU:
    """A single European R&D project."""

    code: str           # e.g. "CM-RD-01"
    entity_code: str    # "CM" or "CP"
    name: str
    objective: str
    technical_challenge: str
    methodology: str
    status: str         # "Active", "Completed"
    qualifies: str      # "yes", "borderline", "no"
    disqualification_reason: str
    # Frascati criteria assessment (for qualifying/borderline)
    frascati_novelty: bool
    frascati_creativity: bool
    frascati_uncertainty: bool
    frascati_systematicity: bool
    frascati_transferability: bool


# CM projects (10): 7 qualifying, 2 borderline, 1 non-qualifying
# CP projects (4): 3 qualifying, 1 non-qualifying
RD_PROJECTS_EU: tuple[RDProjectEU, ...] = (
    # ── CM (Lyon) — CIR eligible ─────────────────────────────────────
    RDProjectEU(
        code="CM-RD-01",
        entity_code="CM",
        name="Advanced Composite Reinforcement Fibres",
        objective=(
            "Develop carbon-fibre reinforced polymer composites with 30% "
            "improved impact resistance for aerospace fuselage panels"
        ),
        technical_challenge=(
            "Whether resin infusion at reduced pressure (< 0.2 bar) can "
            "achieve uniform fibre wet-out in thick-section (> 25 mm) laminates "
            "without void formation exceeding 1%"
        ),
        methodology=(
            "Systematic variation of infusion pressure, resin viscosity, and "
            "fabric architecture using DOE with CT scan and mechanical testing"
        ),
        status="Active",
        qualifies="yes",
        disqualification_reason="",
        frascati_novelty=True,
        frascati_creativity=True,
        frascati_uncertainty=True,
        frascati_systematicity=True,
        frascati_transferability=True,
    ),
    RDProjectEU(
        code="CM-RD-02",
        entity_code="CM",
        name="Polymer Synthesis for High-Temperature Seals",
        objective=(
            "Synthesise fluoroelastomer compounds maintaining elasticity "
            "at 280°C for turbine shaft seal applications"
        ),
        technical_challenge=(
            "Whether cross-link density can be increased without "
            "embrittlement above 250°C — current formulations fail "
            "compression set tests after 500 h at 280°C"
        ),
        methodology=(
            "Iterative monomer ratio adjustment, accelerated ageing tests, "
            "DMA and DSC characterisation of candidate compounds"
        ),
        status="Active",
        qualifies="yes",
        disqualification_reason="",
        frascati_novelty=True,
        frascati_creativity=True,
        frascati_uncertainty=True,
        frascati_systematicity=True,
        frascati_transferability=True,
    ),
    RDProjectEU(
        code="CM-RD-03",
        entity_code="CM",
        name="Nano-Coating for Corrosion Protection",
        objective=(
            "Develop sol-gel derived nano-ceramic coatings providing "
            "5,000-hour salt-spray resistance on aluminium alloys"
        ),
        technical_challenge=(
            "Whether nano-particle agglomeration during spray application "
            "can be prevented while maintaining coating thickness uniformity "
            "(± 5 µm) across complex geometries"
        ),
        methodology=(
            "Nanoparticle surface functionalisation trials, spray parameter "
            "optimisation, EIS and ASTM B117 corrosion testing"
        ),
        status="Active",
        qualifies="yes",
        disqualification_reason="",
        frascati_novelty=True,
        frascati_creativity=True,
        frascati_uncertainty=True,
        frascati_systematicity=True,
        frascati_transferability=True,
    ),
    RDProjectEU(
        code="CM-RD-04",
        entity_code="CM",
        name="Catalyst Development for Green Hydrogen",
        objective=(
            "Design non-precious-metal catalysts for PEM electrolysis "
            "achieving > 1.5 A/cm² at 1.8 V cell voltage"
        ),
        technical_challenge=(
            "Whether iron-nitrogen-carbon (Fe-N-C) catalysts can maintain "
            "activity and stability over 10,000 hours in acidic PEM "
            "environment — current degradation rate is 15 µV/h"
        ),
        methodology=(
            "Combinatorial synthesis of Fe-N-C compositions, rotating disk "
            "electrode screening, single-cell MEA durability testing"
        ),
        status="Active",
        qualifies="yes",
        disqualification_reason="",
        frascati_novelty=True,
        frascati_creativity=True,
        frascati_uncertainty=True,
        frascati_systematicity=True,
        frascati_transferability=True,
    ),
    RDProjectEU(
        code="CM-RD-05",
        entity_code="CM",
        name="Biocompatible Surface Treatment for Implants",
        objective=(
            "Develop plasma-sprayed hydroxyapatite coatings on titanium "
            "implants with improved osseointegration (> 80% bone-implant "
            "contact at 12 weeks in vivo)"
        ),
        technical_challenge=(
            "Whether crystallinity of hydroxyapatite coating (target > 70%) "
            "can be maintained during atmospheric plasma spray without "
            "decomposition to tricalcium phosphate"
        ),
        methodology=(
            "Plasma spray parameter optimisation (power, standoff, feed "
            "rate), XRD crystallinity measurement, simulated body fluid "
            "immersion tests, in vitro cell adhesion assays"
        ),
        status="Active",
        qualifies="yes",
        disqualification_reason="",
        frascati_novelty=True,
        frascati_creativity=True,
        frascati_uncertainty=True,
        frascati_systematicity=True,
        frascati_transferability=True,
    ),
    RDProjectEU(
        code="CM-RD-06",
        entity_code="CM",
        name="Thermal Barrier Coatings for Turbine Blades",
        objective=(
            "Develop suspension plasma-sprayed TBC systems with thermal "
            "conductivity < 0.8 W/m·K and cyclic life > 2,000 cycles"
        ),
        technical_challenge=(
            "Whether columnar microstructure achievable via SPS can "
            "match EB-PVD strain tolerance while reducing cost by 60%"
        ),
        methodology=(
            "SPS deposition trials with varying suspension concentration "
            "and substrate temperature, thermal cycling to failure, "
            "cross-section SEM analysis of columnar structure"
        ),
        status="Completed",
        qualifies="yes",
        disqualification_reason="",
        frascati_novelty=True,
        frascati_creativity=True,
        frascati_uncertainty=True,
        frascati_systematicity=True,
        frascati_transferability=True,
    ),
    RDProjectEU(
        code="CM-RD-07",
        entity_code="CM",
        name="Recyclable Polymer Formulation for Automotive",
        objective=(
            "Formulate thermoplastic composites with mechanical properties "
            "matching thermoset equivalents while enabling end-of-life "
            "recycling through controlled depolymerisation"
        ),
        technical_challenge=(
            "Whether dynamic covalent bonds (vitrimers) can be incorporated "
            "into the resin system while maintaining Tg > 150°C and lap "
            "shear strength > 25 MPa"
        ),
        methodology=(
            "Vitrimer chemistry synthesis, rheological characterisation, "
            "mechanical testing of recycled material vs. virgin, "
            "life-cycle assessment of recyclability"
        ),
        status="Active",
        qualifies="yes",
        disqualification_reason="",
        frascati_novelty=True,
        frascati_creativity=True,
        frascati_uncertainty=True,
        frascati_systematicity=True,
        frascati_transferability=True,
    ),
    # ── CM borderline projects (2) ───────────────────────────────────
    RDProjectEU(
        code="CM-RD-08",
        entity_code="CM",
        name="Product Line Extension Study — Specialty Adhesives",
        objective=(
            "Adapt existing epoxy adhesive formulations for new automotive "
            "OEM specifications (PSA Group bonding requirements)"
        ),
        technical_challenge=(
            "Adjusting cure profile and filler loading to meet revised "
            "lap-shear and peel specifications — within known chemistry"
        ),
        methodology=(
            "Systematic variation of filler type and loading within "
            "established formulation framework, standard mechanical testing"
        ),
        status="Active",
        qualifies="borderline",
        disqualification_reason=(
            "Routine adaptation of existing materials for new client "
            "specifications — lacks genuine technical uncertainty. "
            "Adjustments are within well-understood formulation space."
        ),
        frascati_novelty=False,
        frascati_creativity=False,
        frascati_uncertainty=False,
        frascati_systematicity=True,
        frascati_transferability=True,
    ),
    RDProjectEU(
        code="CM-RD-09",
        entity_code="CM",
        name="EU REACH Regulatory Compliance Testing",
        objective=(
            "Conduct testing programme to meet new EU REACH Annex XVII "
            "restrictions on PFAS in surface treatment chemicals"
        ),
        technical_challenge=(
            "Screening alternative non-PFAS surface treatments that "
            "meet existing performance specifications — arguably "
            "systematic but primarily compliance-driven"
        ),
        methodology=(
            "Testing candidate replacement chemistries against existing "
            "ISO and ASTM performance standards, toxicological assessment"
        ),
        status="Active",
        qualifies="borderline",
        disqualification_reason=(
            "Testing to meet regulatory requirements — arguably systematic "
            "but primarily compliance-driven rather than advancing "
            "scientific or technical knowledge."
        ),
        frascati_novelty=False,
        frascati_creativity=True,
        frascati_uncertainty=True,
        frascati_systematicity=True,
        frascati_transferability=False,
    ),
    # ── CM non-qualifying (1) ────────────────────────────────────────
    RDProjectEU(
        code="CM-RD-10",
        entity_code="CM",
        name="R&D Market Analysis — European Specialty Chemicals Demand",
        objective=(
            "Assess European specialty chemicals market demand for "
            "advanced polymer products and establish competitive pricing"
        ),
        technical_challenge=(
            "None — this is market research, not scientific or "
            "technical research"
        ),
        methodology=(
            "Competitive landscape analysis, customer interviews, "
            "trade show attendance, pricing benchmarks"
        ),
        status="Active",
        qualifies="no",
        disqualification_reason=(
            "Pure market research relabelled as 'R&D market analysis'. "
            "Does not aim to advance scientific or technical knowledge. "
            "No technical uncertainty or systematic experimentation."
        ),
        frascati_novelty=False,
        frascati_creativity=False,
        frascati_uncertainty=False,
        frascati_systematicity=False,
        frascati_transferability=False,
    ),
    # ── CP (Munich) — Forschungszulage eligible ──────────────────────
    RDProjectEU(
        code="CP-RD-01",
        entity_code="CP",
        name="Precision Laser Machining Process Optimisation",
        objective=(
            "Develop femtosecond laser ablation parameters for micro-"
            "machining hardened steel (62 HRC) with surface roughness "
            "Ra < 0.1 µm and no heat-affected zone"
        ),
        technical_challenge=(
            "Whether pulse duration and repetition rate can be optimised "
            "to eliminate recast layer formation in deep (> 500 µm) "
            "micro-channels — current process produces 15 µm recast"
        ),
        methodology=(
            "Systematic laser parameter variation (pulse energy, "
            "repetition rate, scan speed), SEM/EDX surface analysis, "
            "profilometry for roughness measurement"
        ),
        status="Active",
        qualifies="yes",
        disqualification_reason="",
        frascati_novelty=True,
        frascati_creativity=True,
        frascati_uncertainty=True,
        frascati_systematicity=True,
        frascati_transferability=True,
    ),
    RDProjectEU(
        code="CP-RD-02",
        entity_code="CP",
        name="Additive Manufacturing Tolerances — Metal Powder Sintering",
        objective=(
            "Achieve dimensional tolerance of ± 20 µm on SLM-produced "
            "316L stainless steel components without post-machining"
        ),
        technical_challenge=(
            "Whether in-situ melt pool monitoring and adaptive scan "
            "strategy can compensate for thermal distortion in "
            "overhanging features (> 45° from vertical)"
        ),
        methodology=(
            "Closed-loop control system development with pyrometer "
            "feedback, build-plate thermal management, CMM measurement "
            "of test artefacts with known overhang features"
        ),
        status="Active",
        qualifies="yes",
        disqualification_reason="",
        frascati_novelty=True,
        frascati_creativity=True,
        frascati_uncertainty=True,
        frascati_systematicity=True,
        frascati_transferability=True,
    ),
    RDProjectEU(
        code="CP-RD-03",
        entity_code="CP",
        name="Automated Quality Inspection — Computer Vision",
        objective=(
            "Develop deep-learning computer vision system for automated "
            "detection of surface defects (cracks, porosity, inclusions) "
            "on precision-machined components at production line speed"
        ),
        technical_challenge=(
            "Whether a CNN trained on synthetic defect images can "
            "generalise to real production defects with > 95% detection "
            "rate and < 2% false positive rate across varying surface "
            "finishes (ground, polished, as-machined)"
        ),
        methodology=(
            "Synthetic training data generation via GAN, transfer "
            "learning from pre-trained models, validation on real "
            "production samples with known defect maps"
        ),
        status="Active",
        qualifies="yes",
        disqualification_reason="",
        frascati_novelty=True,
        frascati_creativity=True,
        frascati_uncertainty=True,
        frascati_systematicity=True,
        frascati_transferability=True,
    ),
    # ── CP non-qualifying (1) ────────────────────────────────────────
    RDProjectEU(
        code="CP-RD-04",
        entity_code="CP",
        name="Quality R&D Programme — Statistical Process Control",
        objective=(
            "Implement enhanced SPC monitoring across all CNC machining "
            "centres to reduce out-of-tolerance parts from 2.1% to 1.5%"
        ),
        technical_challenge=(
            "None — uses established SPC methods (X-bar/R charts, "
            "Cpk analysis) applied to existing processes"
        ),
        methodology=(
            "Deploy SPC software on existing CNC machines, establish "
            "control limits, train operators, track Cpk trends"
        ),
        status="Active",
        qualifies="no",
        disqualification_reason=(
            "Routine statistical process control and defect tracking "
            "using established methods. No technical uncertainty, "
            "novelty, or creative advancement — fails Frascati criteria."
        ),
        frascati_novelty=False,
        frascati_creativity=False,
        frascati_uncertainty=False,
        frascati_systematicity=True,
        frascati_transferability=False,
    ),
)


# ── R&D Employee definitions ─────────────────────────────────────────────────

@dataclass(frozen=True)
class RDEmployeeEU:
    """An R&D employee at a European entity."""

    employee_id: str
    name: str
    entity_code: str    # "CM" or "CP"
    role: str
    annual_gross_salary_eur: int
    employer_social_charges_eur: int  # ~45% for FR, ~20% for DE

    @property
    def total_employer_cost_eur(self) -> int:
        return self.annual_gross_salary_eur + self.employer_social_charges_eur


def _cm_employees() -> list[RDEmployeeEU]:
    """30 R&D researchers/engineers at CM (Lyon).

    French cotisations sociales ≈ 45% of gross.
    Mix of senior researchers, engineers, and lab technicians.
    """
    employees: list[RDEmployeeEU] = []
    # Names - French-sounding for Lyon entity
    # Salary levels calibrated so aggregate ≈ €2.8M to match design target.
    # French R&D salaries in Lyon: senior researchers €125-145k, engineers
    # €88-105k, scientists €105-115k, lab techs €52-62k, associates €72-78k.
    names_roles_salaries = [
        ("CM-E001", "Dr. Marie Dupont", "Senior Research Scientist", 130000),
        ("CM-E002", "Dr. Pierre Laurent", "Senior Research Scientist", 125000),
        ("CM-E003", "Dr. Isabelle Moreau", "Senior Research Scientist", 135000),
        ("CM-E004", "Dr. Jean-Philippe Bernard", "Principal Researcher", 148000),
        ("CM-E005", "Dr. Claire Lefevre", "Principal Researcher", 142000),
        ("CM-E006", "Antoine Rousseau", "Research Engineer", 98000),
        ("CM-E007", "Sophie Martin", "Research Engineer", 92000),
        ("CM-E008", "Nicolas Girard", "Research Engineer", 101000),
        ("CM-E009", "Camille Dubois", "Research Engineer", 95000),
        ("CM-E010", "Julien Petit", "Research Engineer", 97000),
        ("CM-E011", "Mathilde Simon", "Research Engineer", 93000),
        ("CM-E012", "Romain Fontaine", "Research Engineer", 99000),
        ("CM-E013", "Léa Bonnet", "Research Engineer", 90000),
        ("CM-E014", "Thomas Mercier", "Research Engineer", 103000),
        ("CM-E015", "Chloé Lambert", "Research Engineer", 88000),
        ("CM-E016", "Hugo Roux", "Materials Scientist", 108000),
        ("CM-E017", "Emma Fournier", "Materials Scientist", 105000),
        ("CM-E018", "Lucas Morel", "Materials Scientist", 112000),
        ("CM-E019", "Manon André", "Process Engineer", 96000),
        ("CM-E020", "Alexandre Garnier", "Process Engineer", 99000),
        ("CM-E021", "Pauline Blanc", "Lab Technician", 58000),
        ("CM-E022", "Maxime Guérin", "Lab Technician", 55000),
        ("CM-E023", "Inès Chevalier", "Lab Technician", 60000),
        ("CM-E024", "Florian Muller", "Lab Technician", 56000),
        ("CM-E025", "Anaïs Lemaire", "Lab Technician", 62000),
        ("CM-E026", "Quentin Rivière", "Lab Technician", 53000),
        ("CM-E027", "Marine Garcia", "Lab Technician", 59000),
        ("CM-E028", "Bastien Henry", "Lab Technician", 54000),
        ("CM-E029", "Élodie Robin", "Research Associate", 76000),
        ("CM-E030", "Valentin Picard", "Research Associate", 72000),
    ]
    for emp_id, name, role, salary in names_roles_salaries:
        charges = round(salary * 0.45)
        employees.append(RDEmployeeEU(
            employee_id=emp_id,
            name=name,
            entity_code="CM",
            role=role,
            annual_gross_salary_eur=salary,
            employer_social_charges_eur=charges,
        ))
    return employees


def _cp_employees() -> list[RDEmployeeEU]:
    """12 R&D engineers at CP (Munich).

    German Sozialversicherungsbeiträge ≈ 20% employer share.
    """
    employees: list[RDEmployeeEU] = []
    # Salary levels calibrated so time-weighted qualifying personnel cost
    # ≈ €680k (matching design target for ~€170k Forschungszulage at 25%).
    # German R&D salaries in Munich: senior engineers €95-105k,
    # engineers €78-88k, lab techs €55-62k.
    names_roles_salaries = [
        ("CP-E001", "Dr. Stefan Weber", "Senior R&D Engineer", 102000),
        ("CP-E002", "Dr. Katharina Müller", "Senior R&D Engineer", 98000),
        ("CP-E003", "Markus Schneider", "R&D Engineer", 85000),
        ("CP-E004", "Anna Fischer", "R&D Engineer", 82000),
        ("CP-E005", "Tobias Wagner", "R&D Engineer", 88000),
        ("CP-E006", "Julia Becker", "R&D Engineer", 80000),
        ("CP-E007", "Florian Hoffmann", "R&D Engineer", 84000),
        ("CP-E008", "Sabine Schäfer", "Process Engineer", 86000),
        ("CP-E009", "Michael Koch", "Process Engineer", 90000),
        ("CP-E010", "Claudia Bauer", "Lab Technician", 58000),
        ("CP-E011", "Andreas Richter", "Lab Technician", 55000),
        ("CP-E012", "Nina Klein", "Lab Technician", 57000),
    ]
    for emp_id, name, role, salary in names_roles_salaries:
        charges = round(salary * 0.20)
        employees.append(RDEmployeeEU(
            employee_id=emp_id,
            name=name,
            entity_code="CP",
            role=role,
            annual_gross_salary_eur=salary,
            employer_social_charges_eur=charges,
        ))
    return employees


def generate_rd_employees_eu() -> list[RDEmployeeEU]:
    """Return the complete R&D employee roster for CM + CP."""
    employees = _cm_employees() + _cp_employees()
    employees.sort(key=lambda e: e.employee_id)
    return employees


# ── Time records ──────────────────────────────────────────────────────────────

@dataclass(frozen=True)
class TimeRecordEU:
    """Weekly time record for an EU R&D employee."""

    entity_code: str
    employee_id: str
    employee_name: str
    week_ending: str         # DD.MM.YYYY format (European)
    project_code: str
    hours: float
    activity_description: str


def _week_ending_dates_eu(year: int) -> list[datetime.date]:
    """Return Friday dates for the year, capped at 52."""
    jan1 = datetime.date(year, 1, 1)
    days_to_friday = (4 - jan1.weekday()) % 7
    first_friday = jan1 + datetime.timedelta(days=days_to_friday) if days_to_friday else jan1
    fridays: list[datetime.date] = []
    current = first_friday
    dec31 = datetime.date(year, 12, 31)
    while current <= dec31 and len(fridays) < WEEKS_IN_YEAR:
        fridays.append(current)
        current += datetime.timedelta(days=7)
    return fridays


# Activity descriptions for qualifying R&D work
_EU_QUALIFYING_ACTIVITIES: list[str] = [
    "Experimental design and parameter optimisation",
    "Prototype fabrication and testing",
    "Data analysis and characterisation",
    "Literature review and hypothesis development",
    "Simulation and computational modelling",
    "Laboratory testing and measurement",
    "Process development and scale-up trials",
    "Failure analysis and root cause investigation",
    "Material property characterisation",
    "Design of experiments (DOE) execution",
]

_EU_BORDERLINE_ACTIVITIES: list[str] = [
    "Client specification adaptation testing",
    "Standard testing per existing specifications",
    "Regulatory compliance screening",
    "Sample preparation for qualification testing",
    "Data collection against known benchmarks",
]

_EU_NON_QUALIFYING_ACTIVITIES: list[str] = [
    "Market research and competitive analysis",
    "Customer interviews and requirements gathering",
    "SPC monitoring and Cpk trend analysis",
    "Project management and status reporting",
    "Trade show preparation and attendance",
]

_EU_OVERHEAD_CODES: list[tuple[str, str]] = [
    ("GEN-EU-01", "Production support and troubleshooting"),
    ("GEN-EU-02", "Team meetings and administrative duties"),
    ("GEN-EU-03", "Training and professional development"),
    ("GEN-EU-04", "Quality assurance documentation"),
    ("GEN-EU-05", "Customer technical support"),
    ("GEN-EU-06", "Equipment maintenance"),
]


def _assign_projects_deterministic(
    employees: list[RDEmployeeEU],
) -> dict[str, list[str]]:
    """Assign each employee to R&D projects deterministically.

    CM employees work on CM projects, CP employees on CP projects.
    Some employees also log time to non-qualifying projects.
    """
    cm_projects = [p for p in RD_PROJECTS_EU if p.entity_code == "CM"]
    cp_projects = [p for p in RD_PROJECTS_EU if p.entity_code == "CP"]

    assignments: dict[str, list[str]] = {}
    for i, emp in enumerate(employees):
        if emp.entity_code == "CM":
            # Rotate through qualifying CM projects as primary
            qualifying_cm = [p for p in cm_projects if p.qualifies in ("yes", "borderline")]
            primary_idx = i % len(qualifying_cm)
            projects = [qualifying_cm[primary_idx].code]
            # Every 3rd CM employee gets a secondary project
            if i % 3 == 0 and len(qualifying_cm) > 1:
                secondary_idx = (primary_idx + 1) % len(qualifying_cm)
                projects.append(qualifying_cm[secondary_idx].code)
            # First 2 CM employees also log time to non-qualifying market research
            if i < 2:
                projects.append("CM-RD-10")
        else:
            # CP employees
            qualifying_cp = [p for p in cp_projects if p.qualifies == "yes"]
            cp_idx = (i - 30) % len(qualifying_cp)  # offset by CM employee count
            projects = [qualifying_cp[cp_idx].code]
            # First CP employee also logs to non-qualifying SPC project
            if emp.employee_id == "CP-E001":
                projects.append("CP-RD-04")

        assignments[emp.employee_id] = projects
    return assignments


def generate_time_records_eu() -> list[TimeRecordEU]:
    """Generate weekly time records for all EU R&D employees.

    ~2,184 rows (42 employees × 52 weeks).
    Each employee logs 1 row per week.
    R&D fraction varies by seniority:
    - Senior/Principal: 45-55% of weeks on R&D
    - Engineer/Scientist: 32-48%
    - Lab Tech/Associate: 20-38%

    Deterministic: uses employee index + week index for decisions.
    """
    employees = generate_rd_employees_eu()
    fridays = _week_ending_dates_eu(FY)
    project_by_code = {p.code: p for p in RD_PROJECTS_EU}
    assignments = _assign_projects_deterministic(employees)

    # Deterministic R&D fraction per employee based on role.
    # These are dedicated R&D employees, so fractions are higher than
    # the US TC-08 (where AM employees split time with production).
    def _rd_fraction(emp: RDEmployeeEU, emp_idx: int) -> float:
        """Return deterministic R&D fraction based on role and index."""
        base = hash(emp.employee_id) % 1000 / 10000.0  # 0.0 to 0.1 jitter
        if "Senior" in emp.role or "Principal" in emp.role:
            return 0.72 + base
        if "Engineer" in emp.role or "Scientist" in emp.role:
            return 0.62 + base
        return 0.48 + base  # Lab Tech, Associate

    records: list[TimeRecordEU] = []

    for emp_idx, emp in enumerate(employees):
        rd_frac = _rd_fraction(emp, emp_idx)
        emp_projects = assignments[emp.employee_id]

        for week_idx, friday in enumerate(fridays):
            # Deterministic: use (emp_idx * 53 + week_idx) to decide
            decision_val = (emp_idx * 53 + week_idx) % 100 / 100.0
            is_rd_week = decision_val < rd_frac

            # Hours: deterministic based on indices (36-44 range)
            hours = 38.0 + ((emp_idx * 7 + week_idx * 3) % 13) * 0.5

            if is_rd_week:
                # Pick project deterministically
                proj_idx = (emp_idx + week_idx) % len(emp_projects)
                proj_code = emp_projects[proj_idx]
                proj = project_by_code[proj_code]

                if proj.qualifies == "yes":
                    act_idx = (emp_idx + week_idx) % len(_EU_QUALIFYING_ACTIVITIES)
                    activity = _EU_QUALIFYING_ACTIVITIES[act_idx]
                elif proj.qualifies == "borderline":
                    act_idx = (emp_idx + week_idx) % len(_EU_BORDERLINE_ACTIVITIES)
                    activity = _EU_BORDERLINE_ACTIVITIES[act_idx]
                else:
                    act_idx = (emp_idx + week_idx) % len(_EU_NON_QUALIFYING_ACTIVITIES)
                    activity = _EU_NON_QUALIFYING_ACTIVITIES[act_idx]
            else:
                # Overhead week
                overhead_idx = (emp_idx + week_idx) % len(_EU_OVERHEAD_CODES)
                proj_code, activity = _EU_OVERHEAD_CODES[overhead_idx]

            records.append(TimeRecordEU(
                entity_code=emp.entity_code,
                employee_id=emp.employee_id,
                employee_name=emp.name,
                week_ending=friday.strftime("%d.%m.%Y"),
                project_code=proj_code,
                hours=hours,
                activity_description=activity,
            ))

    records.sort(key=lambda r: (r.week_ending, r.entity_code, r.employee_id))
    return records


# ── Supply expenses ───────────────────────────────────────────────────────────

@dataclass(frozen=True)
class RDSupplyExpenseEU:
    """R&D supply/materials expense line for CM or CP."""

    entity_code: str
    expense_date: str     # DD.MM.YYYY format
    description: str
    amount_eur: Decimal
    cost_center: str
    project_code: str


# Planted error: ERR-EU-008
# This line at CP is coded to CP-RD-03 (qualifying) but is actually a
# maintenance contract, not an R&D supply. For Forschungszulage, supplies
# don't qualify anyway — but the misclassification is the error to detect.
_ERR_EU_008_EXPENSE = RDSupplyExpenseEU(
    entity_code="CP",
    expense_date="15.06.2025",
    description="Calibration service — annual maintenance contract",
    amount_eur=Decimal("12400"),
    cost_center="4500",
    project_code="CP-RD-03",
)


def generate_supply_expenses_eu() -> list[RDSupplyExpenseEU]:
    """Generate R&D supply expenses for CM and CP.

    CM supplies: eligible for CIR (~€280k total across qualifying projects).
    CP supplies: NOT eligible for Forschungszulage.
    Includes ERR-EU-008 (maintenance contract miscoded as R&D supply at CP).
    """
    expenses: list[RDSupplyExpenseEU] = []

    # CM supply expenses (~€280k across 12 months)
    cm_supplies = [
        ("15.01.2025", "High-purity ceramic powders — Al2O3", Decimal("18500"), "CM-RD-01"),
        ("22.01.2025", "Carbon fibre prepreg material", Decimal("32000"), "CM-RD-01"),
        ("05.02.2025", "Fluoroelastomer monomers — batch FE-2025-02", Decimal("15800"), "CM-RD-02"),
        ("18.02.2025", "Sol-gel precursors — cerium nitrate", Decimal("8900"), "CM-RD-03"),
        ("28.02.2025", "Laboratory chemicals and reagents", Decimal("6200"), "CM-RD-04"),
        ("12.03.2025", "Titanium rod stock — Grade 5", Decimal("24500"), "CM-RD-05"),
        ("25.03.2025", "Thermal spray powder — YSZ 8%", Decimal("19800"), "CM-RD-06"),
        ("08.04.2025", "Vitrimer resin components", Decimal("12300"), "CM-RD-07"),
        ("22.04.2025", "SEM sample preparation supplies", Decimal("4200"), "CM-RD-01"),
        ("10.05.2025", "Catalyst synthesis precursors — FeCl3/melamine", Decimal("9600"), "CM-RD-04"),
        ("28.05.2025", "Hydroxyapatite powder — medical grade", Decimal("16700"), "CM-RD-05"),
        ("15.06.2025", "Carbon nanotubes — multi-wall, >95%", Decimal("22100"), "CM-RD-03"),
        ("30.06.2025", "Compression testing consumables", Decimal("3800"), "CM-RD-02"),
        ("18.07.2025", "PEM membrane material — Nafion 212", Decimal("11400"), "CM-RD-04"),
        ("05.08.2025", "Plasma spray nozzle assemblies", Decimal("8500"), "CM-RD-06"),
        ("22.08.2025", "Recycled polymer pellets — validation batch", Decimal("5600"), "CM-RD-07"),
        ("10.09.2025", "X-ray diffraction sample holders", Decimal("2800"), "CM-RD-03"),
        ("25.09.2025", "Corrosion test panels — AA2024-T3", Decimal("7200"), "CM-RD-03"),
        ("15.10.2025", "Sintering furnace consumables", Decimal("14300"), "CM-RD-01"),
        ("30.10.2025", "Polymer characterisation reagents — DSC/TGA", Decimal("6800"), "CM-RD-02"),
        ("12.11.2025", "Biocompatibility test media — cell culture", Decimal("9200"), "CM-RD-05"),
        ("28.11.2025", "Metallographic preparation consumables", Decimal("3400"), "CM-RD-06"),
        ("10.12.2025", "Accelerated ageing test chamber supplies", Decimal("7900"), "CM-RD-07"),
        ("20.12.2025", "Year-end lab consumables order", Decimal("12500"), "CM-RD-01"),
    ]

    for date_str, desc, amount, proj in cm_supplies:
        expenses.append(RDSupplyExpenseEU(
            entity_code="CM",
            expense_date=date_str,
            description=desc,
            amount_eur=amount,
            cost_center="3800",  # CM R&D cost centre
            project_code=proj,
        ))

    # CP supply expenses (NOT eligible for Forschungszulage, but present in data)
    cp_supplies = [
        ("20.01.2025", "Laser optics — replacement lens assembly", Decimal("8700"), "CP-RD-01"),
        ("15.03.2025", "Metal powder feedstock — 316L, 20 µm", Decimal("18500"), "CP-RD-02"),
        ("10.05.2025", "Camera modules — high-speed industrial", Decimal("14200"), "CP-RD-03"),
        ("25.07.2025", "Precision measurement fixtures", Decimal("6800"), "CP-RD-01"),
        ("15.09.2025", "GPU modules — NVIDIA A100 for training", Decimal("22000"), "CP-RD-03"),
        ("20.11.2025", "Test artefacts — CMM reference standards", Decimal("5400"), "CP-RD-02"),
    ]

    for date_str, desc, amount, proj in cp_supplies:
        expenses.append(RDSupplyExpenseEU(
            entity_code="CP",
            expense_date=date_str,
            description=desc,
            amount_eur=amount,
            cost_center="4500",  # CP R&D cost centre
            project_code=proj,
        ))

    # ERR-EU-008: maintenance contract miscoded as R&D supply
    expenses.append(_ERR_EU_008_EXPENSE)

    expenses.sort(key=lambda e: (e.entity_code, e.expense_date, e.project_code))
    return expenses


# ── Subcontractor invoices ────────────────────────────────────────────────────

@dataclass(frozen=True)
class SubcontractorInvoiceEU:
    """Subcontracted R&D invoice for CM (CIR eligible)."""

    invoice_id: str
    subcontractor_name: str
    project_code: str
    description: str
    amount_eur: Decimal
    subcontractor_type: str  # "public" or "private"


def generate_subcontractor_invoices_eu() -> list[SubcontractorInvoiceEU]:
    """Generate subcontracted R&D invoices for CM (Lyon).

    4 invoices totalling ~€420k.
    - 3 private subcontractors
    - 1 public research organism (university) → gets 2× multiplier in CIR
    - 1 invoice is for the non-qualifying market research project (must exclude)
    """
    invoices = [
        SubcontractorInvoiceEU(
            invoice_id="SUB-CM-001",
            subcontractor_name="Laboratoire de Mécanique des Contacts et Structures (INSA Lyon)",
            project_code="CM-RD-01",
            description="Subcontracted fatigue testing and microstructural analysis of composite specimens",
            amount_eur=Decimal("80000"),
            subcontractor_type="public",
        ),
        SubcontractorInvoiceEU(
            invoice_id="SUB-CM-002",
            subcontractor_name="Materia Nova ASBL",
            project_code="CM-RD-03",
            description="Nano-coating characterisation and corrosion testing programme",
            amount_eur=Decimal("120000"),
            subcontractor_type="private",
        ),
        SubcontractorInvoiceEU(
            invoice_id="SUB-CM-003",
            subcontractor_name="Polymat Ingénierie SAS",
            project_code="CM-RD-07",
            description="Vitrimer rheological testing and recyclability assessment",
            amount_eur=Decimal("140000"),
            subcontractor_type="private",
        ),
        # Non-qualifying: market research firm for CM-RD-10 (must be excluded)
        SubcontractorInvoiceEU(
            invoice_id="SUB-CM-004",
            subcontractor_name="EuroChemInsight Consulting GmbH",
            project_code="CM-RD-10",
            description="European specialty chemicals market sizing and competitive landscape analysis",
            amount_eur=Decimal("80000"),
            subcontractor_type="private",
        ),
    ]
    invoices.sort(key=lambda i: i.invoice_id)
    return invoices


# ── CIR computation ──────────────────────────────────────────────────────────

@dataclass(frozen=True)
class CIRResult:
    """CIR computation result for CM (Lyon)."""

    qualifying_researcher_salary_eur: Decimal
    forfaitaire_overhead_eur: Decimal
    personnel_total_eur: Decimal
    qualifying_supplies_eur: Decimal
    subcontracted_private_eur: Decimal
    subcontracted_public_eur: Decimal
    subcontracted_public_doubled_eur: Decimal  # 2× multiplier applied
    total_subcontracted_eur: Decimal
    total_eligible_base_eur: Decimal
    cir_credit_eur: Decimal


def compute_cir(
    employees: list[RDEmployeeEU],
    time_records: list[TimeRecordEU],
    supply_expenses: list[RDSupplyExpenseEU],
    subcontractor_invoices: list[SubcontractorInvoiceEU],
) -> CIRResult:
    """Compute French CIR for CM (Lyon).

    CIR = 30% × eligible base, where eligible base =
      + Personnel: researcher gross salaries × 1.43 (forfaitaire)
      + Supplies: consumables on qualifying projects
      + Subcontracted: private at face value, public × 2
      - Exclude: non-qualifying projects (CM-RD-10, borderline included for gold)
    """
    qualifying_codes = {
        p.code for p in RD_PROJECTS_EU
        if p.entity_code == "CM" and p.qualifies in ("yes", "borderline")
    }

    # Personnel: only CM employees who logged time to qualifying projects
    cm_employees = [e for e in employees if e.entity_code == "CM"]
    cm_time = [r for r in time_records if r.entity_code == "CM"]

    # Identify employees with qualifying R&D time
    emp_has_qualifying: set[str] = set()
    for r in cm_time:
        if r.project_code in qualifying_codes:
            emp_has_qualifying.add(r.employee_id)

    qualifying_salary = Decimal(0)
    for emp in cm_employees:
        if emp.employee_id in emp_has_qualifying:
            qualifying_salary += Decimal(emp.annual_gross_salary_eur)

    forfaitaire = (qualifying_salary * CIR_FORFAITAIRE).quantize(
        Decimal("1"), rounding=ROUND_HALF_UP,
    )
    personnel_total = qualifying_salary + forfaitaire

    # Supplies: CM supplies on qualifying projects
    qualifying_supplies = Decimal(0)
    for exp in supply_expenses:
        if exp.entity_code == "CM" and exp.project_code in qualifying_codes:
            qualifying_supplies += exp.amount_eur

    # Subcontractors: exclude CM-RD-10 (non-qualifying)
    sub_private = Decimal(0)
    sub_public = Decimal(0)
    for inv in subcontractor_invoices:
        if inv.project_code not in qualifying_codes:
            continue
        if inv.subcontractor_type == "public":
            sub_public += inv.amount_eur
        else:
            sub_private += inv.amount_eur

    sub_public_doubled = sub_public * 2  # 2× multiplier for public organisms
    total_sub = sub_private + sub_public_doubled

    total_eligible = personnel_total + qualifying_supplies + total_sub

    # CIR: 30% (all under €100M threshold)
    cir_credit = (total_eligible * CIR_RATE).quantize(
        Decimal("1"), rounding=ROUND_HALF_UP,
    )

    return CIRResult(
        qualifying_researcher_salary_eur=qualifying_salary,
        forfaitaire_overhead_eur=forfaitaire,
        personnel_total_eur=personnel_total,
        qualifying_supplies_eur=qualifying_supplies,
        subcontracted_private_eur=sub_private,
        subcontracted_public_eur=sub_public,
        subcontracted_public_doubled_eur=sub_public_doubled,
        total_subcontracted_eur=total_sub,
        total_eligible_base_eur=total_eligible,
        cir_credit_eur=cir_credit,
    )


# ── Forschungszulage computation ──────────────────────────────────────────────

@dataclass(frozen=True)
class ForschungszulageResult:
    """Forschungszulage computation result for CP (Munich)."""

    qualifying_personnel_cost_eur: Decimal
    assessment_basis_eur: Decimal  # min(personnel, €2M cap)
    benefit_eur: Decimal           # 25% of assessment basis


def compute_forschungszulage(
    employees: list[RDEmployeeEU],
    time_records: list[TimeRecordEU],
) -> ForschungszulageResult:
    """Compute German Forschungszulage for CP (Munich).

    Forschungszulage = 25% × min(eligible personnel costs, €2M).
    Only personnel costs of employees directly performing qualifying R&D.
    NO supplies, NO subcontractors.
    """
    qualifying_codes = {
        p.code for p in RD_PROJECTS_EU
        if p.entity_code == "CP" and p.qualifies == "yes"
    }

    cp_employees = [e for e in employees if e.entity_code == "CP"]
    cp_time = [r for r in time_records if r.entity_code == "CP"]

    # Compute time-weighted personnel costs for qualifying projects
    # For each employee: (qualifying hours / total hours) × total employer cost
    emp_total_hours: dict[str, float] = {}
    emp_qualifying_hours: dict[str, float] = {}

    for r in cp_time:
        emp_total_hours[r.employee_id] = emp_total_hours.get(r.employee_id, 0.0) + r.hours
        if r.project_code in qualifying_codes:
            emp_qualifying_hours[r.employee_id] = (
                emp_qualifying_hours.get(r.employee_id, 0.0) + r.hours
            )

    qualifying_personnel = Decimal(0)
    for emp in cp_employees:
        total_h = emp_total_hours.get(emp.employee_id, 0.0)
        qual_h = emp_qualifying_hours.get(emp.employee_id, 0.0)
        if total_h <= 0 or qual_h <= 0:
            continue
        fraction = Decimal(str(round(qual_h / total_h, 6)))
        cost = Decimal(emp.total_employer_cost_eur)
        qualifying_personnel += (cost * fraction).quantize(
            Decimal("1"), rounding=ROUND_HALF_UP,
        )

    assessment_basis = min(qualifying_personnel, FZ_MAX_BASIS)
    benefit = (assessment_basis * FZ_RATE).quantize(
        Decimal("1"), rounding=ROUND_HALF_UP,
    )

    return ForschungszulageResult(
        qualifying_personnel_cost_eur=qualifying_personnel,
        assessment_basis_eur=assessment_basis,
        benefit_eur=benefit,
    )


# ── Consolidated computation ──────────────────────────────────────────────────

@dataclass(frozen=True)
class ConsolidatedRDBenefit:
    """Consolidated R&D incentive summary for the Cascade Europe group."""

    cir: CIRResult
    forschungszulage: ForschungszulageResult
    total_benefit_eur: Decimal


def compute_consolidated_rd_benefit() -> ConsolidatedRDBenefit:
    """Compute the full R&D incentive for CM + CP."""
    employees = generate_rd_employees_eu()
    time_records = generate_time_records_eu()
    supply_expenses = generate_supply_expenses_eu()
    subcontractor_invoices = generate_subcontractor_invoices_eu()

    cir = compute_cir(employees, time_records, supply_expenses, subcontractor_invoices)
    fz = compute_forschungszulage(employees, time_records)

    return ConsolidatedRDBenefit(
        cir=cir,
        forschungszulage=fz,
        total_benefit_eur=cir.cir_credit_eur + fz.benefit_eur,
    )


# ── Canary keys ───────────────────────────────────────────────────────────────

ALL_CANARY_KEYS_TC08EU: list[str] = sorted([
    "tc08eu_time_records",
    "tc08eu_payroll",
    "tc08eu_supply_expenses",
    "tc08eu_subcontractor_invoices",
    "tc08eu_prior_year_rd",
] + [f"tc08eu_project_{i:03d}" for i in range(1, 15)])
