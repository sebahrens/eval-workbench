"""Cascade Industries entity definitions (§1.1–1.2 of prompt.md).

Provides the canonical Entity dataclass and a pre-built dict of all four
entities keyed by their two-letter entity code.
"""

from __future__ import annotations

from dataclasses import dataclass


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
