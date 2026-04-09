"""Cascade Industries R&D projects and time records for TC-08 (Section 41).

Generates 12 R&D projects for Advanced Materials, weekly time records for
45 R&D-eligible employees across 52 weeks, and supply expenses coded to R&D
cost centers.  Computes QREs and the ASC credit (~$185K).

Key constraints from prompt.md TC-08:
- 12 projects: 8 qualify under 4-part test, 2 borderline, 2 don't qualify.
- ~2,340 time record rows (45 employees × 52 weeks, multiple allocations).
- Prior year QREs: FY2023 = $3.1M, FY2024 = $3.4M.
- ASC credit ≈ $185,000 (acceptance: within $500).
- QRE_2025 target ≈ $2,885,714 to hit the credit.
"""

from __future__ import annotations

import datetime
import random
from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal

from generator.model.employees import Employee

# ── Constants ─────────────────────────────────────────────────────────────────

PRIOR_YEAR_QRES: dict[int, Decimal] = {
    2023: Decimal("3_100_000"),
    2024: Decimal("3_400_000"),
}

TARGET_CREDIT = Decimal("185_000")
ASC_RATE = Decimal("0.14")

# Back-computed: Credit = 14% × (QRE_2025 − 50% × avg(QRE_23, QRE_24, QRE_25))
# Solving for QRE_2025 to hit $185,000 → $2,885,714.29 (we'll target $2,885,714)
QRE_2025_TARGET = Decimal("2_885_714")

# Number of R&D-eligible employees who log time to projects.
RD_EMPLOYEE_COUNT = 45

# Fiscal year parameters.
FY = 2025
WEEKS_IN_YEAR = 52
FY_START = datetime.date(FY, 1, 1)


# ── R&D Project definitions ──────────────────────────────────────────────────

@dataclass(frozen=True)
class RDProject:
    """A single R&D project at Cascade Advanced Materials."""

    code: str           # e.g. "RD-001"
    name: str
    objective: str
    uncertainty: str    # Technical uncertainty addressed
    methodology: str
    status: str         # "Active", "Completed", "On Hold"
    qualifies: str      # "yes", "borderline", "no"
    disqualification_reason: str  # empty if qualifies == "yes"
    qre_weight: float   # Relative share of total QREs (0.0 for non-qualifying)


# 12 projects: 8 qualify, 2 borderline, 2 don't qualify.
# QRE weights are normalized so qualifying projects sum to 1.0.
# Borderline projects contribute to QREs (they'd be flagged for manager review
# but their expenses still flow through the computation for the agent to assess).
RD_PROJECTS: tuple[RDProject, ...] = (
    RDProject(
        code="RD-001",
        name="High-Temperature Ceramic Composite Development",
        objective=(
            "Develop ceramic matrix composites capable of sustained "
            "operation above 1,400\u00b0C for turbine blade applications"
        ),
        uncertainty=(
            "Whether sintering parameters can achieve target density "
            "(>97%) without microcracking at grain boundaries"
        ),
        methodology=(
            "Systematic variation of sintering temperature, pressure, "
            "and atmosphere using DOE methodology with SEM analysis"
        ),
        status="Active",
        qualifies="yes",
        disqualification_reason="",
        qre_weight=0.18,
    ),
    RDProject(
        code="RD-002",
        name="Nano-Structured Thermal Barrier Coatings",
        objective=(
            "Create thermal barrier coatings with 40% lower thermal "
            "conductivity using nanostructured yttria-stabilized zirconia"
        ),
        uncertainty=(
            "Whether plasma spray deposition can produce consistent "
            "nanostructure without agglomeration at production scale"
        ),
        methodology=(
            "Iterative testing of spray parameters (standoff distance, "
            "feed rate, power) with XRD and TEM characterization"
        ),
        status="Active",
        qualifies="yes",
        disqualification_reason="",
        qre_weight=0.16,
    ),
    RDProject(
        code="RD-003",
        name="Additive Manufacturing Process for Ti-6Al-4V",
        objective=(
            "Develop selective laser melting parameters for "
            "aerospace-grade titanium alloy with equivalent fatigue "
            "life to wrought material"
        ),
        uncertainty=(
            "Whether post-processing heat treatment can eliminate "
            "porosity-induced fatigue initiation sites below 50\u03bcm"
        ),
        methodology=(
            "Parametric study of laser power, scan speed, and hatch "
            "spacing combined with HIP post-processing evaluation"
        ),
        status="Active",
        qualifies="yes",
        disqualification_reason="",
        qre_weight=0.15,
    ),
    RDProject(
        code="RD-004",
        name="Bio-Inspired Self-Healing Polymer Matrix",
        objective=(
            "Incorporate microencapsulated healing agents into epoxy "
            "matrices for autonomous crack repair in composites"
        ),
        uncertainty=(
            "Whether microcapsule survival during composite layup and "
            "cure cycle can exceed 85% without compromising strength"
        ),
        methodology=(
            "Capsule formulation trials with varying shell thickness, "
            "followed by mechanical testing (DCB, ENF) of healed specimens"
        ),
        status="Active",
        qualifies="yes",
        disqualification_reason="",
        qre_weight=0.12,
    ),
    RDProject(
        code="RD-005",
        name="Graphene-Enhanced Aluminum Alloy Conductors",
        objective=(
            "Integrate graphene nanoplatelets into 6061-T6 aluminum "
            "to improve electrical conductivity by 15% for busbars"
        ),
        uncertainty=(
            "Whether ball milling dispersion can prevent graphene "
            "restacking while maintaining alloy ductility > 8%"
        ),
        methodology=(
            "Varying milling time, graphene loading (0.1-2.0 wt%), "
            "and sintering conditions with 4-point probe and tensile testing"
        ),
        status="Active",
        qualifies="yes",
        disqualification_reason="",
        qre_weight=0.10,
    ),
    RDProject(
        code="RD-006",
        name="Corrosion-Resistant Surface Treatment for Mg Alloys",
        objective=(
            "Develop a non-chromate conversion coating for AZ91D "
            "magnesium that passes 500-hour salt spray testing"
        ),
        uncertainty=(
            "Whether rare-earth-based sol-gel chemistry can achieve "
            "adhesion and corrosion performance equivalent to chromate"
        ),
        methodology=(
            "Screening of Ce/La-based sol-gel formulations with "
            "EIS monitoring and ASTM B117 salt spray validation"
        ),
        status="Completed",
        qualifies="yes",
        disqualification_reason="",
        qre_weight=0.09,
    ),
    RDProject(
        code="RD-007",
        name="Piezoelectric Energy Harvesting from Vibration",
        objective=(
            "Design embedded piezoelectric transducer arrays that "
            "harvest > 5mW from typical aerospace structural vibration"
        ),
        uncertainty=(
            "Whether lead-free piezoceramics (BNT-BT) can achieve "
            "coupling coefficient (k33 > 0.45) at operating temp range"
        ),
        methodology=(
            "FEA modal analysis combined with prototype testing on "
            "vibration table across 10-2000 Hz spectrum"
        ),
        status="Active",
        qualifies="yes",
        disqualification_reason="",
        qre_weight=0.08,
    ),
    RDProject(
        code="RD-008",
        name="Machine Learning for In-Situ Process Monitoring",
        objective=(
            "Develop real-time ML models to detect porosity defects "
            "during laser powder bed fusion using acoustic sensors"
        ),
        uncertainty=(
            "Whether acoustic emission signals contain sufficient "
            "info to classify defects with >90% accuracy in production"
        ),
        methodology=(
            "Collect labeled acoustic data from controlled defect "
            "builds, train CNN classifiers, validate on blind builds"
        ),
        status="Active",
        qualifies="yes",
        disqualification_reason="",
        qre_weight=0.07,
    ),
    # ── Borderline projects (2) — flagged for manager review ─────────
    RDProject(
        code="RD-009",
        name="Automated Quality Inspection System",
        objective=(
            "Implement automated visual inspection using commercial "
            "machine vision cameras for surface defect detection"
        ),
        uncertainty=(
            "Integration challenge: whether commercial algorithms can "
            "be tuned for subtle surface defects in our materials"
        ),
        methodology=(
            "Configure vendor software parameters, run detection "
            "accuracy tests against known defect samples"
        ),
        status="Active",
        qualifies="borderline",
        disqualification_reason=(
            "Primarily involves adapting commercial off-the-shelf "
            "technology; may not meet technological uncertainty test"
        ),
        qre_weight=0.03,
    ),
    RDProject(
        code="RD-010",
        name="Recycled Feedstock Qualification for Powder Metallurgy",
        objective=(
            "Determine whether recycled titanium powder (up to 30% "
            "reclaimed) can meet aerospace material specifications"
        ),
        uncertainty=(
            "Whether oxygen pickup during powder recycling degrades "
            "mechanical properties below spec minimum"
        ),
        methodology=(
            "Blend ratios testing (0%, 10%, 20%, 30% recycled) with "
            "full mechanical property characterization per AMS specs"
        ),
        status="Active",
        qualifies="borderline",
        disqualification_reason=(
            "Testing known materials against existing specs may be "
            "routine quality testing rather than experimentation"
        ),
        qre_weight=0.02,
    ),
    # ── Non-qualifying projects (2) ──────────────────────────────────
    RDProject(
        code="RD-011",
        name="Market Analysis for European Aerospace Applications",
        objective=(
            "Assess European OEM demand for advanced ceramic "
            "components and establish pricing strategy for EU entry"
        ),
        uncertainty=(
            "None \u2014 this is market research, not technological "
            "research"
        ),
        methodology=(
            "Competitive landscape analysis, customer interviews, "
            "trade show attendance, pricing benchmarks"
        ),
        status="Active",
        qualifies="no",
        disqualification_reason=(
            "Market research does not meet the permitted purpose "
            "test; no technological uncertainty or experimentation"
        ),
        qre_weight=0.0,
    ),
    RDProject(
        code="RD-012",
        name="ERP System Migration for R&D Lab Management",
        objective=(
            "Migrate R&D laboratory inventory and equipment tracking "
            "from legacy system to SAP S/4HANA module"
        ),
        uncertainty=(
            "None \u2014 this is an IT implementation project using "
            "established software"
        ),
        methodology=(
            "Gap analysis, data migration, user acceptance testing, "
            "go-live cutover"
        ),
        status="On Hold",
        qualifies="no",
        disqualification_reason=(
            "Internal IT system implementation does not involve "
            "technological uncertainty or experimentation in the "
            "physical/biological/computer sciences"
        ),
        qre_weight=0.0,
    ),
)


# ── Time record dataclass ────────────────────────────────────────────────────

@dataclass(frozen=True)
class TimeRecord:
    """A single weekly time record row for an R&D employee."""

    employee_id: str
    employee_name: str
    week_ending: datetime.date  # Always a Friday
    project_code: str
    hours: float
    activity_description: str


# ── Supply expense dataclass ─────────────────────────────────────────────────

@dataclass(frozen=True)
class RDSupplyExpense:
    """An R&D supply/materials expense line."""

    date: datetime.date
    vendor: str
    description: str
    amount: Decimal
    project_code: str
    cost_center: str  # "3500" for AM R&D


# ── Activity descriptions by project type ────────────────────────────────────

_QUALIFYING_ACTIVITIES: list[str] = [
    "Experimental design and parameter optimization",
    "Prototype fabrication and testing",
    "Data analysis and characterization",
    "Literature review and hypothesis development",
    "Simulation and computational modeling",
    "Lab testing and measurement",
    "Process development and scale-up trials",
    "Failure analysis and root cause investigation",
    "Material property characterization",
    "Design of experiments (DOE) execution",
]

_BORDERLINE_ACTIVITIES: list[str] = [
    "Vendor software configuration and tuning",
    "Standard testing per existing specifications",
    "Calibration and threshold adjustment",
    "Sample preparation for qualification testing",
    "Data collection against known benchmarks",
]

_NON_QUALIFYING_ACTIVITIES: list[str] = [
    "Market research and competitive analysis",
    "Customer interviews and requirements gathering",
    "System administration and data migration",
    "Project management and status reporting",
    "Trade show preparation and attendance",
]

# General / overhead activity codes (not R&D projects — non-qualifying time).
_OVERHEAD_CODES: list[tuple[str, str]] = [
    ("GEN-001", "Production support and troubleshooting"),
    ("GEN-002", "Staff meetings and administrative duties"),
    ("GEN-003", "Training and professional development"),
    ("GEN-004", "Quality assurance and compliance documentation"),
    ("GEN-005", "Customer technical support"),
    ("GEN-006", "Equipment maintenance and calibration"),
]


# ── Supply vendors ───────────────────────────────────────────────────────────

_SUPPLY_VENDORS: list[tuple[str, str]] = [
    ("Alfa Aesar", "High-purity ceramic powders"),
    ("Thermo Fisher Scientific", "Laboratory chemicals and reagents"),
    ("McMaster-Carr", "Precision tooling and fixtures"),
    ("Sigma-Aldrich", "Specialty polymers and solvents"),
    ("Praxair Surface Technologies", "Thermal spray powders"),
    ("AP&C Advanced Powders", "Titanium alloy powder feedstock"),
    ("Hexcel Corporation", "Carbon fiber prepreg materials"),
    ("Oerlikon Metco", "Coating materials and consumables"),
    ("Ted Pella Inc", "SEM/TEM sample preparation supplies"),
    ("Buehler", "Metallographic preparation supplies"),
]


# ── Generator functions ──────────────────────────────────────────────────────

def _select_rd_employees(
    employees: list[Employee],
    rng: random.Random,
) -> list[Employee]:
    """Select exactly 45 active R&D-eligible employees from the roster.

    Filters to AM employees who are R&D-eligible and active in FY2025,
    then takes the first 45 sorted by employee_id for determinism.
    If fewer than 45 are available, raises ValueError.
    """
    fy_start = datetime.date(FY, 1, 1)

    eligible = [
        e for e in employees
        if e.is_rd_eligible
        and e.entity_code == "AM"
        and (e.termination_date is None or e.termination_date >= fy_start)
    ]
    # Sort by employee_id for determinism.
    eligible.sort(key=lambda e: e.employee_id)

    if len(eligible) < RD_EMPLOYEE_COUNT:
        raise ValueError(
            f"Need {RD_EMPLOYEE_COUNT} R&D-eligible AM employees, "
            f"but only {len(eligible)} are active in FY{FY}."
        )

    return eligible[:RD_EMPLOYEE_COUNT]


def _week_ending_dates(year: int) -> list[datetime.date]:
    """Return all Friday dates that serve as week-ending dates for the year."""
    # First Friday of the year.
    jan1 = datetime.date(year, 1, 1)
    days_to_friday = (4 - jan1.weekday()) % 7
    if days_to_friday == 0:
        first_friday = jan1
    else:
        first_friday = jan1 + datetime.timedelta(days=days_to_friday)

    fridays: list[datetime.date] = []
    current = first_friday
    dec31 = datetime.date(year, 12, 31)
    while current <= dec31:
        fridays.append(current)
        current += datetime.timedelta(days=7)

    return fridays[:WEEKS_IN_YEAR]  # Cap at 52


def generate_rd_projects() -> list[RDProject]:
    """Return the canonical list of 12 R&D projects."""
    return list(RD_PROJECTS)


def generate_time_records(
    employees: list[Employee],
    rng: random.Random,
) -> list[TimeRecord]:
    """Generate weekly time records for 45 R&D-eligible employees.

    Each employee gets exactly 1 row per week = 45 × 52 = 2,340 rows.
    Each row shows the primary activity for that week: either an R&D project
    or general/overhead work. The agent computes QRE wages by:
      (hours on qualifying projects / total hours) × W-2 wages.

    Employees spend a varying fraction of weeks on R&D projects (20-55%)
    and the remainder on production support, admin, training, etc.
    The overall qualifying fraction is calibrated so wage QREs ≈ $2.1M,
    leaving ~$800K for supplies to hit QRE_2025_TARGET ≈ $2.885M.
    """
    rd_employees = _select_rd_employees(employees, rng)
    fridays = _week_ending_dates(FY)

    # All 12 R&D projects (qualifying + borderline + non-qualifying).
    # Non-qualifying projects (RD-011, RD-012) also get time allocated —
    # the agent must recognize they don't qualify and exclude those hours.
    project_by_code = {p.code: p for p in RD_PROJECTS}

    # Assign each employee an R&D fraction (what % of weeks they work on
    # R&D projects vs. general overhead). Target average ≈ 38%.
    # Senior scientists: 45-55%, junior/lab techs: 20-35%.
    emp_rd_fraction: dict[str, float] = {}
    for emp in rd_employees:
        if "Senior" in emp.title or "Director" in emp.title:
            frac = rng.uniform(0.42, 0.55)
        elif "Scientist" in emp.title or "Engineer" in emp.title:
            frac = rng.uniform(0.32, 0.48)
        else:  # Lab Tech, Research Associate, etc.
            frac = rng.uniform(0.20, 0.38)
        emp_rd_fraction[emp.employee_id] = frac

    # Assign each employee a primary R&D project (weighted by qre_weight)
    # and optionally a secondary. Non-qualifying projects also get assigned
    # to some employees (RD-011 to a couple, RD-012 to a couple).
    qualifying_projects = [p for p in RD_PROJECTS if p.qualifies in ("yes", "borderline")]
    emp_projects: dict[str, list[str]] = {}
    for i, emp in enumerate(rd_employees):
        # Most employees get qualifying projects.
        primary = rng.choices(
            qualifying_projects,
            weights=[p.qre_weight for p in qualifying_projects],
            k=1,
        )[0]
        projects = [primary.code]

        # 40% chance of a secondary qualifying project.
        if rng.random() < 0.40:
            others = [p for p in qualifying_projects if p.code != primary.code]
            secondary = rng.choices(
                others,
                weights=[p.qre_weight for p in others],
                k=1,
            )[0]
            projects.append(secondary.code)

        # A few employees also log time to non-qualifying projects.
        # This tests whether the agent correctly excludes these.
        if i < 3:  # First 3 employees get RD-011 (market research)
            projects.append("RD-011")
        elif i < 5:  # Next 2 get RD-012 (ERP migration)
            projects.append("RD-012")

        emp_projects[emp.employee_id] = projects

    # Generate exactly 1 record per employee per week.
    records: list[TimeRecord] = []

    for friday in fridays:
        for emp in rd_employees:
            # Skip if not yet hired or already terminated.
            if emp.hire_date > friday:
                continue
            if (emp.termination_date is not None
                    and emp.termination_date < friday - datetime.timedelta(days=6)):
                continue

            # Hours this week (36-44).
            hours = round(max(32.0, min(48.0, rng.gauss(40.0, 2.0))), 1)

            # Decide if this is an R&D week or overhead week.
            rd_frac = emp_rd_fraction[emp.employee_id]
            is_rd_week = rng.random() < rd_frac

            if is_rd_week:
                # Pick one of the employee's assigned R&D projects.
                proj_code = rng.choice(emp_projects[emp.employee_id])
                proj = project_by_code[proj_code]

                if proj.qualifies == "yes":
                    activity = rng.choice(_QUALIFYING_ACTIVITIES)
                elif proj.qualifies == "borderline":
                    activity = rng.choice(_BORDERLINE_ACTIVITIES)
                else:
                    activity = rng.choice(_NON_QUALIFYING_ACTIVITIES)
            else:
                # Overhead / non-R&D week.
                proj_code, activity = rng.choice(_OVERHEAD_CODES)

            records.append(TimeRecord(
                employee_id=emp.employee_id,
                employee_name=emp.name,
                week_ending=friday,
                project_code=proj_code,
                hours=hours,
                activity_description=activity,
            ))

    records.sort(key=lambda r: (r.week_ending, r.employee_id, r.project_code))
    return records


def generate_supply_expenses(
    rng: random.Random,
) -> list[RDSupplyExpense]:
    """Generate R&D supply and materials expenses for FY2025.

    Expenses are allocated to qualifying projects proportionally to their
    QRE weight. Total supply expenses are calibrated so that
    wages + supplies = QRE_2025_TARGET.

    Returns sorted list of RDSupplyExpense records.
    """
    # We generate ~120 supply expense lines spread across the year.
    # Total supply amount is set after wage calculation in compute_qres().
    # Here we generate the line items with relative amounts that will
    # be scaled by the caller.
    expenses: list[RDSupplyExpense] = []
    allocable = [p for p in RD_PROJECTS if p.qualifies == "yes"]

    for month in range(1, 13):
        # ~10 purchases per month
        n_purchases = rng.randint(8, 12)
        for _ in range(n_purchases):
            day = rng.randint(1, 28)
            date = datetime.date(FY, month, day)

            proj = rng.choices(
                allocable,
                weights=[p.qre_weight for p in allocable],
                k=1,
            )[0]

            vendor, default_desc = rng.choice(_SUPPLY_VENDORS)
            amount = Decimal(str(round(rng.uniform(500, 15_000), 2)))

            expenses.append(RDSupplyExpense(
                date=date,
                vendor=vendor,
                description=default_desc,
                amount=amount,
                project_code=proj.code,
                cost_center="3500",  # AM R&D cost center
            ))

    expenses.sort(key=lambda e: (e.date, e.project_code, e.vendor))
    return expenses


# ── QRE computation ──────────────────────────────────────────────────────────

@dataclass
class QREResult:
    """Qualified Research Expenses computation for a fiscal year."""

    year: int
    wage_qres: Decimal          # Wages allocated to qualifying activities
    supply_qres: Decimal        # Qualifying supply expenses
    total_qres: Decimal         # wage_qres + supply_qres
    qres_by_project: dict[str, Decimal]  # project_code → QRE amount
    credit: Decimal             # ASC credit amount


def compute_qres(
    time_records: list[TimeRecord],
    supply_expenses: list[RDSupplyExpense],
    employees: list[Employee],
) -> QREResult:
    """Compute Qualified Research Expenses and ASC credit for FY2025.

    QRE wages = for each employee on qualifying/borderline projects,
    (hours on project / total hours) × W-2 wages (annual_salary).

    Supply QREs = supply expenses on qualifying projects.

    The credit is computed using the Alternative Simplified Credit method:
    Credit = 14% × (QRE_2025 − 50% × average(QRE_2023, QRE_2024, QRE_2025))
    """
    emp_by_id = {e.employee_id: e for e in employees if e.is_rd_eligible}

    # Compute hours by employee × project for FY2025.
    emp_project_hours: dict[str, dict[str, float]] = {}  # emp_id → {proj → hours}
    emp_total_hours: dict[str, float] = {}

    for rec in time_records:
        emp_project_hours.setdefault(rec.employee_id, {})
        emp_project_hours[rec.employee_id][rec.project_code] = (
            emp_project_hours[rec.employee_id].get(rec.project_code, 0.0) + rec.hours
        )
        emp_total_hours[rec.employee_id] = (
            emp_total_hours.get(rec.employee_id, 0.0) + rec.hours
        )

    # Qualifying project codes (yes + borderline for flagging).
    qualifying_codes = {p.code for p in RD_PROJECTS if p.qualifies in ("yes", "borderline")}

    # Wage QREs by project.
    wage_by_project: dict[str, Decimal] = {}
    total_wage_qres = Decimal(0)

    for emp_id in sorted(emp_project_hours.keys()):
        emp = emp_by_id.get(emp_id)
        if emp is None:
            continue

        total_h = emp_total_hours.get(emp_id, 0.0)
        if total_h <= 0:
            continue

        salary = Decimal(emp.annual_salary)

        for proj_code in sorted(emp_project_hours[emp_id].keys()):
            if proj_code not in qualifying_codes:
                continue

            proj_hours = emp_project_hours[emp_id][proj_code]
            fraction = Decimal(str(round(proj_hours / total_h, 6)))
            wage_qre = (salary * fraction).quantize(
                Decimal("0.01"), rounding=ROUND_HALF_UP
            )
            wage_by_project[proj_code] = (
                wage_by_project.get(proj_code, Decimal(0)) + wage_qre
            )
            total_wage_qres += wage_qre

    # Supply QREs (only qualifying projects).
    supply_by_project: dict[str, Decimal] = {}
    total_supply_qres = Decimal(0)

    for exp in supply_expenses:
        if exp.project_code not in qualifying_codes:
            continue
        supply_by_project[exp.project_code] = (
            supply_by_project.get(exp.project_code, Decimal(0)) + exp.amount
        )
        total_supply_qres += exp.amount

    # Scale supply expenses so total QREs hit the target.
    # QRE_target = wages + scaled_supplies
    # scaled_supplies = QRE_target - wages
    supply_target = QRE_2025_TARGET - total_wage_qres
    if total_supply_qres > 0 and supply_target > 0:
        scale_factor = supply_target / total_supply_qres
        # Re-scale supply_by_project.
        for proj_code in sorted(supply_by_project.keys()):
            supply_by_project[proj_code] = (
                supply_by_project[proj_code] * scale_factor
            ).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        total_supply_qres = sum(supply_by_project.values())

    total_qres = total_wage_qres + total_supply_qres

    # Merge into by-project totals.
    all_projects = sorted(set(list(wage_by_project.keys()) + list(supply_by_project.keys())))
    qres_by_project: dict[str, Decimal] = {}
    for proj_code in all_projects:
        qres_by_project[proj_code] = (
            wage_by_project.get(proj_code, Decimal(0))
            + supply_by_project.get(proj_code, Decimal(0))
        )

    # ASC credit computation.
    avg_3yr = (PRIOR_YEAR_QRES[2023] + PRIOR_YEAR_QRES[2024] + total_qres) / 3
    credit = (ASC_RATE * (total_qres - Decimal("0.50") * avg_3yr)).quantize(
        Decimal("0.01"), rounding=ROUND_HALF_UP
    )

    return QREResult(
        year=FY,
        wage_qres=total_wage_qres,
        supply_qres=total_supply_qres,
        total_qres=total_qres,
        qres_by_project=qres_by_project,
        credit=credit,
    )
