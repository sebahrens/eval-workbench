"""Cascade Industries 850-employee roster (§1.4 of prompt.md).

Deterministic generation of the full employee database. Every employee has:
employee_id, name, entity, department, title, hire_date, annual_salary,
state, cost_center, is_r&d_eligible, termination_date.

Key constraints:
- 850 total employees across CI, PC, AM, DS (default).
- R&D eligibility only for Advanced Materials (AM) R&D and Engineering staff.
- ~8% annual turnover → ~68 terminated employees.
- Hire dates distributed across 3 years (2022-01-01 to 2024-12-31).
- Deterministic via seeded random + Faker.

When a :class:`~generator.config.Config` is provided, headcounts, turnover
rate, remote states, and entity codes are derived from it.  Without config,
hardcoded Cascade defaults are used (backward compatible).
"""

from __future__ import annotations

import datetime
import random
from dataclasses import dataclass
from typing import TYPE_CHECKING

from faker import Faker

from generator.model.entities import ENTITIES, entities_from_config

if TYPE_CHECKING:
    from generator.config import Config

# ── Department / title / salary configuration ─────────────────────────────────

# Department → list of (title, base_salary_low, base_salary_high)
# Salaries are base ranges; location multipliers adjust them.
DEPARTMENTS: dict[str, list[tuple[str, int, int]]] = {
    "Engineering": [
        ("Engineering Manager", 120_000, 160_000),
        ("Senior Engineer", 95_000, 130_000),
        ("Engineer", 70_000, 100_000),
        ("Junior Engineer", 55_000, 75_000),
    ],
    "Manufacturing": [
        ("Plant Manager", 110_000, 145_000),
        ("Production Supervisor", 65_000, 85_000),
        ("Machine Operator", 38_000, 52_000),
        ("Quality Inspector", 45_000, 62_000),
        ("Assembly Technician", 35_000, 48_000),
    ],
    "Sales": [
        ("VP Sales", 140_000, 180_000),
        ("Sales Manager", 90_000, 120_000),
        ("Account Executive", 65_000, 95_000),
        ("Sales Representative", 45_000, 65_000),
    ],
    "G&A": [
        ("Chief Executive Officer", 250_000, 350_000),
        ("Chief Financial Officer", 200_000, 280_000),
        ("Chief Operating Officer", 200_000, 280_000),
        ("VP Operations", 130_000, 170_000),
        ("Office Manager", 50_000, 70_000),
        ("Administrative Assistant", 35_000, 50_000),
        ("Executive Assistant", 45_000, 65_000),
    ],
    "R&D": [
        ("R&D Director", 140_000, 180_000),
        ("Senior Research Scientist", 110_000, 145_000),
        ("Research Scientist", 85_000, 115_000),
        ("Lab Technician", 50_000, 70_000),
        ("Research Associate", 60_000, 85_000),
    ],
    "Warehouse": [
        ("Warehouse Manager", 70_000, 95_000),
        ("Warehouse Supervisor", 50_000, 65_000),
        ("Forklift Operator", 32_000, 42_000),
        ("Shipping Clerk", 30_000, 40_000),
        ("Receiving Clerk", 30_000, 40_000),
    ],
    "Finance": [
        ("Controller", 120_000, 160_000),
        ("Senior Accountant", 75_000, 100_000),
        ("Staff Accountant", 55_000, 75_000),
        ("Accounts Payable Specialist", 40_000, 55_000),
        ("Payroll Specialist", 42_000, 58_000),
        ("Financial Analyst", 65_000, 90_000),
    ],
}

# Location salary multipliers (cost-of-living adjustment).
_LOCATION_MULTIPLIER: dict[str, float] = {
    "OR": 1.00,
    "TX": 1.05,
    "IL": 0.95,
    "CA": 1.15,
    "WA": 1.10,
    "NY": 1.12,
}

# Entity → department weight distribution.
# Weights are relative; they get normalized to produce the target headcount.
_ENTITY_DEPT_WEIGHTS: dict[str, dict[str, float]] = {
    "CI": {
        "G&A": 0.45,
        "Finance": 0.35,
        "Sales": 0.20,
    },
    "PC": {
        "Engineering": 0.15,
        "Manufacturing": 0.40,
        "Sales": 0.10,
        "G&A": 0.08,
        "Finance": 0.07,
        "Warehouse": 0.15,
        "R&D": 0.05,
    },
    "AM": {
        "Engineering": 0.18,
        "Manufacturing": 0.25,
        "R&D": 0.22,
        "Sales": 0.08,
        "G&A": 0.07,
        "Finance": 0.06,
        "Warehouse": 0.14,
    },
    "DS": {
        "Warehouse": 0.45,
        "Sales": 0.12,
        "G&A": 0.10,
        "Finance": 0.08,
        "Manufacturing": 0.15,
        "Engineering": 0.10,
    },
}

# Target headcounts per entity (must sum to 850).
_ENTITY_HEADCOUNTS: dict[str, int] = {
    "CI": 40,
    "PC": 330,
    "AM": 270,
    "DS": 210,
}

# Cost center prefixes by entity.
_COST_CENTER_BASE: dict[str, int] = {
    "CI": 1000,
    "PC": 2000,
    "AM": 3000,
    "DS": 4000,
}

# Department → cost center offset within entity.
_DEPT_CC_OFFSET: dict[str, int] = {
    "Engineering": 100,
    "Manufacturing": 200,
    "Sales": 300,
    "G&A": 400,
    "R&D": 500,
    "Warehouse": 600,
    "Finance": 700,
}

# Remote states (some employees work from CA, WA, NY).
_REMOTE_STATES = ["CA", "WA", "NY"]
_REMOTE_PROBABILITY = 0.06  # 6% of employees are remote


@dataclass(frozen=True)
class Employee:
    """A single employee in the Cascade Industries group."""

    employee_id: str
    name: str
    entity_code: str
    entity_name: str
    department: str
    title: str
    hire_date: datetime.date
    annual_salary: int
    state: str
    cost_center: str
    is_rd_eligible: bool
    termination_date: datetime.date | None


def generate_employees(
    rng: random.Random,
    config: Config | None = None,
) -> list[Employee]:
    """Generate the full employee roster deterministically.

    Parameters
    ----------
    rng : random.Random
        Seeded PRNG for determinism.
    config : Config, optional
        When provided, headcounts come from subsidiary ``employee_count``
        fields, turnover rate from ``config.company.employees``, and
        remote states from the same section.  Without config the
        hardcoded Cascade defaults (850 employees) are used.

    Returns
    -------
    list[Employee]
        Sorted by employee_id.
    """
    fake = Faker()
    Faker.seed(rng.randint(0, 2**31))  # Derive Faker seed from our RNG

    # ── Resolve parameters from config or hardcoded defaults ────────
    if config is not None:
        all_entities, _ = entities_from_config(config.company)
        entity_headcounts = _headcounts_from_config(config)
        turnover_rate = config.company.employees.annual_turnover_rate
        remote_states = list(config.company.employees.remote_states)
    else:
        all_entities = ENTITIES
        entity_headcounts = dict(_ENTITY_HEADCOUNTS)
        turnover_rate = 0.08
        remote_states = list(_REMOTE_STATES)

    employees: list[Employee] = []

    # Process entities in sorted order for determinism.
    for entity_code in sorted(entity_headcounts.keys()):
        entity = all_entities[entity_code]
        headcount = entity_headcounts[entity_code]
        dept_weights = _ENTITY_DEPT_WEIGHTS.get(entity_code)
        if dept_weights is None or headcount == 0:
            continue

        # Distribute headcount across departments proportionally.
        dept_counts = _distribute_headcount(headcount, dept_weights, rng)

        seq = 1  # Sequential employee counter within entity.
        for dept_name in sorted(dept_counts.keys()):
            count = dept_counts[dept_name]
            titles_pool = DEPARTMENTS[dept_name]

            for _ in range(count):
                # Pick title (weighted toward more junior roles).
                title_idx = min(
                    rng.randint(0, len(titles_pool) - 1),
                    rng.randint(0, len(titles_pool) - 1),
                )
                title_name, sal_low, sal_high = titles_pool[title_idx]

                # State: default is entity's state, some remote.
                state = entity.state
                if not entity.is_parent and rng.random() < _REMOTE_PROBABILITY:
                    state = rng.choice(remote_states)

                # Salary: base range × location multiplier, rounded to nearest $1000.
                multiplier = _LOCATION_MULTIPLIER.get(state, 1.0)
                raw_salary = rng.randint(sal_low, sal_high) * multiplier
                salary = round(raw_salary / 1000) * 1000

                # Hire date: distributed across 2022-01-01 to 2024-12-31.
                hire_date = _random_hire_date(rng)

                # Termination: ~8% annual turnover.
                termination_date = _maybe_terminate(rng, hire_date, turnover_rate)

                # R&D eligibility: only AM R&D and Engineering staff.
                is_rd_eligible = (
                    entity_code == "AM"
                    and dept_name in ("R&D", "Engineering")
                )

                # Cost center: entity base + department offset.
                cc_base = _COST_CENTER_BASE.get(entity_code, 5000)
                cc_offset = _DEPT_CC_OFFSET[dept_name]
                cost_center = str(cc_base + cc_offset)

                employee_id = f"{entity_code}-{seq:04d}"

                employees.append(
                    Employee(
                        employee_id=employee_id,
                        name=fake.name(),
                        entity_code=entity_code,
                        entity_name=entity.name,
                        department=dept_name,
                        title=title_name,
                        hire_date=hire_date,
                        annual_salary=salary,
                        state=state,
                        cost_center=cost_center,
                        is_rd_eligible=is_rd_eligible,
                        termination_date=termination_date,
                    )
                )
                seq += 1

    # Sort by employee_id for deterministic output order.
    employees.sort(key=lambda e: e.employee_id)
    return employees


def _headcounts_from_config(config: Config) -> dict[str, int]:
    """Derive per-entity headcounts from config.

    Subsidiary headcounts come from each subsidiary's ``employee_count``.
    Parent headcount is ``total_count - sum(subsidiary counts)``.
    """
    from generator.config import CompanyConfig

    company: CompanyConfig = config.company
    total = company.employees.total_count

    # Build parent entity code (same logic as entities_from_config)
    words = company.name.replace(",", "").replace(".", "").split()
    parent_code = "".join(w[0] for w in words if w[0].isupper())[:2]

    headcounts: dict[str, int] = {}
    sub_total = 0
    for _key, sub in sorted(company.subsidiaries.items()):
        headcounts[sub.entity_code] = sub.employee_count
        sub_total += sub.employee_count

    parent_headcount = total - sub_total
    if parent_headcount > 0:
        headcounts[parent_code] = parent_headcount

    return headcounts


def _distribute_headcount(
    total: int,
    weights: dict[str, float],
    rng: random.Random,
) -> dict[str, int]:
    """Distribute total headcount across departments by weight.

    Uses largest-remainder method for exact total. Departments processed
    in sorted order for determinism.
    """
    sorted_depts = sorted(weights.keys())
    total_weight = sum(weights[d] for d in sorted_depts)

    # Compute ideal (fractional) counts.
    ideal = {d: total * weights[d] / total_weight for d in sorted_depts}

    # Floor each and track remainders.
    floored = {d: int(ideal[d]) for d in sorted_depts}
    remainders = {d: ideal[d] - floored[d] for d in sorted_depts}

    # Distribute remaining slots by largest remainder.
    remaining = total - sum(floored.values())
    # Sort by remainder descending, then by name for stability.
    by_remainder = sorted(sorted_depts, key=lambda d: (-remainders[d], d))
    for i in range(remaining):
        floored[by_remainder[i]] += 1

    return floored


def _random_hire_date(rng: random.Random) -> datetime.date:
    """Generate a hire date between 2022-01-01 and 2024-12-31.

    Slightly weighted toward earlier dates (established workforce).
    """
    start = datetime.date(2022, 1, 1)
    end = datetime.date(2024, 12, 31)
    days_range = (end - start).days

    # Weight toward earlier: take min of two uniform draws.
    day_offset = min(rng.randint(0, days_range), rng.randint(0, days_range))
    return start + datetime.timedelta(days=day_offset)


def _maybe_terminate(
    rng: random.Random,
    hire_date: datetime.date,
    turnover_rate: float = 0.08,
) -> datetime.date | None:
    """Decide if an employee is terminated, and if so, when.

    Termination date is after hire_date and before 2025-12-31.
    Returns None for active employees.
    """
    # Calculate tenure in years to determine termination probability.
    reference = datetime.date(2025, 12, 31)
    tenure_days = (reference - hire_date).days
    tenure_years = tenure_days / 365.25

    # Probability of having been terminated at some point during tenure.
    # Using 1 - (1 - rate)^years for compound probability.
    prob_terminated = 1.0 - (1.0 - turnover_rate) ** tenure_years

    if rng.random() >= prob_terminated:
        return None

    # Termination date: random point after hire, before end of FY2025.
    earliest = hire_date + datetime.timedelta(days=90)  # Min 90-day tenure
    latest = datetime.date(2025, 12, 31)
    if earliest >= latest:
        return None

    term_days = rng.randint(0, (latest - earliest).days)
    return earliest + datetime.timedelta(days=term_days)
