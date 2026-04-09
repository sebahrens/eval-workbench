"""Configuration loader and validator for the Cascade Industries test suite generator."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml

# ---------------------------------------------------------------------------
# Dataclasses — typed, immutable representations of the config tree
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class SubsidiaryConfig:
    legal_name: str
    location: str
    state: str
    entity_code: str
    revenue: int
    type: str
    gross_margin: float
    employee_count: int
    rd_spend_pct: float = 0.0


@dataclass(frozen=True)
class GrowthRates:
    fy2023_to_fy2024: float
    fy2024_to_fy2025: float


@dataclass(frozen=True)
class IntercompanyConfig:
    raw_materials_markup: float
    management_fee_pct: float
    intercompany_loan_principal: int
    intercompany_loan_rate: float


@dataclass(frozen=True)
class EmployeeConfig:
    total_count: int
    annual_turnover_rate: float
    remote_states: list[str] = field(default_factory=lambda: ["CA", "WA", "NY"])


@dataclass(frozen=True)
class SeasonalWeights:
    Q1: float
    Q2: float
    Q3: float
    Q4: float

    def __post_init__(self) -> None:
        total = self.Q1 + self.Q2 + self.Q3 + self.Q4
        if abs(total - 1.0) > 1e-6:
            raise ConfigError(f"Seasonal weights must sum to 1.0, got {total}")


@dataclass(frozen=True)
class CompanyConfig:
    name: str
    type: str
    industry: str
    headquarters: str
    fiscal_year_end: str
    years: list[int]
    current_year: int
    consolidated_revenue: int
    subsidiaries: dict[str, SubsidiaryConfig]
    growth_rates: GrowthRates
    intercompany: IntercompanyConfig
    employees: EmployeeConfig
    seasonal_weights: SeasonalWeights


@dataclass(frozen=True)
class Config:
    seed: int
    output_dir: str
    company: CompanyConfig
    canary_assignments: dict[str, str]
    error_injections: dict[str, Any]


# ---------------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------------

class ConfigError(Exception):
    """Raised when the configuration file is missing, malformed, or invalid."""


# ---------------------------------------------------------------------------
# Validation helpers
# ---------------------------------------------------------------------------

_REQUIRED_TOP = {"seed", "output_dir", "company"}
_REQUIRED_COMPANY = {
    "name", "type", "industry", "headquarters", "fiscal_year_end",
    "years", "current_year", "consolidated_revenue", "subsidiaries",
    "growth_rates", "intercompany", "employees", "seasonal_weights",
}
_REQUIRED_SUBSIDIARY = {
    "legal_name", "location", "state", "entity_code", "revenue",
    "type", "gross_margin", "employee_count",
}


def _require_keys(data: dict, required: set[str], context: str) -> None:
    missing = required - set(data)
    if missing:
        raise ConfigError(f"Missing required keys in {context}: {sorted(missing)}")


def _parse_subsidiary(key: str, raw: dict) -> SubsidiaryConfig:
    _require_keys(raw, _REQUIRED_SUBSIDIARY, f"subsidiaries.{key}")
    return SubsidiaryConfig(
        legal_name=raw["legal_name"],
        location=raw["location"],
        state=raw["state"],
        entity_code=raw["entity_code"],
        revenue=int(raw["revenue"]),
        type=raw["type"],
        gross_margin=float(raw["gross_margin"]),
        employee_count=int(raw["employee_count"]),
        rd_spend_pct=float(raw.get("rd_spend_pct", 0.0)),
    )


def _parse_company(raw: dict) -> CompanyConfig:
    _require_keys(raw, _REQUIRED_COMPANY, "company")

    subs = {k: _parse_subsidiary(k, v) for k, v in raw["subsidiaries"].items()}
    if not subs:
        raise ConfigError("At least one subsidiary is required")

    return CompanyConfig(
        name=raw["name"],
        type=raw["type"],
        industry=raw["industry"],
        headquarters=raw["headquarters"],
        fiscal_year_end=raw["fiscal_year_end"],
        years=sorted(raw["years"]),
        current_year=int(raw["current_year"]),
        consolidated_revenue=int(raw["consolidated_revenue"]),
        subsidiaries=subs,
        growth_rates=GrowthRates(**raw["growth_rates"]),
        intercompany=IntercompanyConfig(**raw["intercompany"]),
        employees=EmployeeConfig(**raw["employees"]),
        seasonal_weights=SeasonalWeights(**raw["seasonal_weights"]),
    )


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def load_config(path: str | Path) -> Config:
    """Load, validate, and return a typed Config from a YAML file.

    Raises ConfigError on any structural or semantic problem.
    """
    path = Path(path)
    if not path.exists():
        raise ConfigError(f"Config file not found: {path}")

    try:
        with open(path) as f:
            raw = yaml.safe_load(f)
    except yaml.YAMLError as exc:
        raise ConfigError(f"Invalid YAML in {path}: {exc}") from exc

    if not isinstance(raw, dict):
        raise ConfigError(f"Config root must be a mapping, got {type(raw).__name__}")

    _require_keys(raw, _REQUIRED_TOP, "config root")

    company = _parse_company(raw["company"])

    return Config(
        seed=int(raw["seed"]),
        output_dir=raw["output_dir"],
        company=company,
        canary_assignments=raw.get("canary_assignments") or {},
        error_injections=raw.get("error_injections") or {},
    )
