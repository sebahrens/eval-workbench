"""Cascade Industries entity definitions (§1.1–1.2 of prompt.md).

Provides the canonical Entity dataclass and a pre-built dict of all four
entities keyed by their two-letter entity code.

The ``entities_from_config`` adapter builds identical Entity dicts from a
:class:`~generator.config.CompanyConfig`, allowing custom scenarios to
override the hardcoded Cascade defaults below.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from generator.config import CompanyConfig


@dataclass(frozen=True)
class Entity:
    """A single legal entity in the Cascade Industries group."""

    code: str  # Two-letter prefix (CI, PC, AM, DS)
    name: str  # Legal name
    location: str  # City, State
    state: str  # Two-letter state abbreviation
    revenue_target: int  # FY2025 target revenue in dollars
    gross_margin: float  # Gross margin as a decimal (e.g. 0.35)
    revenue_mix: str  # Short description of business type
    is_parent: bool = False


# ── Constants from prompt.md §1.1–1.2 ──────────────────────────────────────

PARENT = Entity(
    code="CI",
    name="Cascade Industries, Inc.",
    location="Portland, OR",
    state="OR",
    revenue_target=200_000_000,
    gross_margin=0.0,  # consolidated; not meaningful at parent level
    revenue_mix="US C-Corporation, mid-market manufacturer (holding company)",
    is_parent=True,
)

PRECISION_COMPONENTS = Entity(
    code="PC",
    name="Cascade Precision Components LLC",
    location="Portland, OR",
    state="OR",
    revenue_target=95_000_000,
    gross_margin=0.35,
    revenue_mix="Core manufacturing (industrial parts)",
)

ADVANCED_MATERIALS = Entity(
    code="AM",
    name="Cascade Advanced Materials, Inc.",
    location="Austin, TX",
    state="TX",
    revenue_target=65_000_000,
    gross_margin=0.52,
    revenue_mix="Specialty materials R&D and manufacturing",
)

DISTRIBUTION_SERVICES = Entity(
    code="DS",
    name="Cascade Distribution Services LLC",
    location="Chicago, IL",
    state="IL",
    revenue_target=40_000_000,
    gross_margin=0.18,
    revenue_mix="Warehousing and logistics",
)

# Canonical lookup: entity code → Entity
ENTITIES: dict[str, Entity] = {
    e.code: e
    for e in [PARENT, PRECISION_COMPONENTS, ADVANCED_MATERIALS, DISTRIBUTION_SERVICES]
}

# Subsidiaries only (excludes parent)
SUBSIDIARIES: dict[str, Entity] = {
    code: entity for code, entity in ENTITIES.items() if not entity.is_parent
}


# ── Config-to-Entity adapter ────────────────────────────────────────────────

def entities_from_config(company: CompanyConfig) -> tuple[dict[str, Entity], dict[str, Entity]]:
    """Derive canonical entity dicts from a :class:`CompanyConfig`.

    Returns ``(all_entities, subsidiaries)`` — the same shape as the
    module-level ``ENTITIES`` and ``SUBSIDIARIES`` constants, but built
    from config rather than hardcoded values.

    The parent entity is synthesized from company-level fields
    (``name``, ``headquarters``, ``consolidated_revenue``).  Subsidiary
    entities come directly from each :class:`SubsidiaryConfig`.
    """
    # Parse "City, State" from headquarters (e.g. "Portland, Oregon" → "OR")
    # Config uses full state names in headquarters; derive the two-letter
    # abbreviation from the first subsidiary in that state, or fall back to
    # the first two letters of the last word.
    hq_parts = company.headquarters.rsplit(",", 1)
    hq_city_state = company.headquarters
    if len(hq_parts) == 2:
        hq_city_state = f"{hq_parts[0].strip()}, {hq_parts[1].strip()[:2].upper()}"

    # Try to find the parent state from a subsidiary at the same city
    parent_state = hq_parts[1].strip()[:2].upper() if len(hq_parts) == 2 else "XX"
    for sub in company.subsidiaries.values():
        if sub.location.startswith(hq_parts[0].strip()):
            parent_state = sub.state
            hq_city_state = f"{hq_parts[0].strip()}, {parent_state}"
            break

    # Build parent entity code from company name initials
    words = company.name.replace(",", "").replace(".", "").split()
    parent_code = "".join(w[0] for w in words if w[0].isupper())[:2]

    parent = Entity(
        code=parent_code,
        name=company.name,
        location=hq_city_state,
        state=parent_state,
        revenue_target=company.consolidated_revenue,
        gross_margin=0.0,
        revenue_mix=f"{company.type}, {company.industry} (holding company)",
        is_parent=True,
    )

    all_entities: dict[str, Entity] = {parent_code: parent}
    subsidiaries: dict[str, Entity] = {}

    for _key, sub in sorted(company.subsidiaries.items()):
        entity = Entity(
            code=sub.entity_code,
            name=sub.legal_name,
            location=sub.location,
            state=sub.state,
            revenue_target=sub.revenue,
            gross_margin=sub.gross_margin,
            revenue_mix=sub.type,
        )
        all_entities[sub.entity_code] = entity
        subsidiaries[sub.entity_code] = entity

    return all_entities, subsidiaries
