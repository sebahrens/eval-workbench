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
class AugmentationConfig:
    enabled: bool = False
    model: str = ""
    cache_dir: str = ".augmentation_cache"
    warm_on_miss: bool = False


# ---------------------------------------------------------------------------
# v1 customization profile dataclasses (synth-data-2u6.2)
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class DifficultyProfile:
    """Controls error/canary/trap density for scenario difficulty tuning."""
    error_density: float = 1.0
    canary_visibility: str = "visible"
    judgment_trap_density: float = 1.0

    def __post_init__(self) -> None:
        if not 0.0 <= self.error_density <= 1.0:
            raise ConfigError(
                f"difficulty.error_density must be 0.0–1.0, got {self.error_density}"
            )
        if self.canary_visibility not in ("visible", "subtle", "hidden"):
            raise ConfigError(
                f"difficulty.canary_visibility must be 'visible', 'subtle', or 'hidden', "
                f"got {self.canary_visibility!r}"
            )
        if not 0.0 <= self.judgment_trap_density <= 1.0:
            raise ConfigError(
                f"difficulty.judgment_trap_density must be 0.0–1.0, "
                f"got {self.judgment_trap_density}"
            )


@dataclass(frozen=True)
class OutputProfile:
    """Controls which test cases and packs the generator emits."""
    enabled_test_cases: list[str] = field(default_factory=list)
    enabled_packs: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class ErrorProfile:
    """Overrides which planted errors are injected."""
    include: list[str] = field(default_factory=list)
    exclude: list[str] = field(default_factory=list)
    density_override: float | None = None

    def __post_init__(self) -> None:
        if self.density_override is not None and not 0.0 <= self.density_override <= 1.0:
            raise ConfigError(
                f"errors.density_override must be 0.0–1.0 or null, "
                f"got {self.density_override}"
            )
        overlap = set(self.include) & set(self.exclude)
        if overlap:
            raise ConfigError(
                f"errors.include and errors.exclude overlap: {sorted(overlap)}"
            )


@dataclass(frozen=True)
class Config:
    seed: int
    output_dir: str
    company: CompanyConfig
    canary_assignments: dict[str, str]
    error_injections: dict[str, Any]
    augmentation: AugmentationConfig = field(default_factory=AugmentationConfig)
    difficulty: DifficultyProfile = field(default_factory=DifficultyProfile)
    output: OutputProfile = field(default_factory=OutputProfile)
    errors: ErrorProfile = field(default_factory=ErrorProfile)


# ---------------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------------

class ConfigError(Exception):
    """Raised when the configuration file is missing, malformed, or invalid."""


# ---------------------------------------------------------------------------
# Validation helpers
# ---------------------------------------------------------------------------

_REQUIRED_TOP = {"seed", "output_dir", "company"}
_ALLOWED_TOP = _REQUIRED_TOP | {
    "canary_assignments", "error_injections", "augmentation",
    "difficulty", "output", "errors",
}
_ALLOWED_DIFFICULTY = {"error_density", "canary_visibility", "judgment_trap_density"}
_ALLOWED_OUTPUT = {"enabled_test_cases", "enabled_packs"}
_ALLOWED_ERRORS = {"include", "exclude", "density_override"}
_REQUIRED_COMPANY = {
    "name", "type", "industry", "headquarters", "fiscal_year_end",
    "years", "current_year", "consolidated_revenue", "subsidiaries",
    "growth_rates", "intercompany", "employees", "seasonal_weights",
}
_ALLOWED_COMPANY = _REQUIRED_COMPANY  # No optional company-level keys in v1
_REQUIRED_SUBSIDIARY = {
    "legal_name", "location", "state", "entity_code", "revenue",
    "type", "gross_margin", "employee_count",
}
_ALLOWED_SUBSIDIARY = _REQUIRED_SUBSIDIARY | {"rd_spend_pct"}


def _require_keys(data: dict, required: set[str], context: str) -> None:
    missing = required - set(data)
    if missing:
        raise ConfigError(f"Missing required keys in {context}: {sorted(missing)}")


def _reject_unknown_keys(
    data: dict, allowed: set[str], context: str,
) -> None:
    unknown = set(data) - allowed
    if unknown:
        raise ConfigError(
            f"Unknown key(s) in {context}: {sorted(unknown)}. "
            f"Only these keys are supported in v1: {sorted(allowed)}"
        )


def _parse_difficulty(raw: dict) -> DifficultyProfile:
    _reject_unknown_keys(raw, _ALLOWED_DIFFICULTY, "difficulty")
    return DifficultyProfile(
        error_density=float(raw.get("error_density", 1.0)),
        canary_visibility=str(raw.get("canary_visibility", "visible")),
        judgment_trap_density=float(raw.get("judgment_trap_density", 1.0)),
    )


def _parse_output(raw: dict) -> OutputProfile:
    _reject_unknown_keys(raw, _ALLOWED_OUTPUT, "output")
    return OutputProfile(
        enabled_test_cases=list(raw.get("enabled_test_cases") or []),
        enabled_packs=list(raw.get("enabled_packs") or []),
    )


def _parse_errors(raw: dict) -> ErrorProfile:
    _reject_unknown_keys(raw, _ALLOWED_ERRORS, "errors")
    density = raw.get("density_override")
    return ErrorProfile(
        include=list(raw.get("include") or []),
        exclude=list(raw.get("exclude") or []),
        density_override=float(density) if density is not None else None,
    )


# ---------------------------------------------------------------------------
# Deep merge for layered config (synth-data-2u6.2)
# ---------------------------------------------------------------------------

def deep_merge(base: dict, overlay: dict) -> dict:
    """Deep-merge *overlay* into *base*, returning a new dict.

    Merge semantics (per customization-schema-v1.md):
    - Dicts are recursively merged (not replaced).
    - Lists are replaced wholesale (not appended).
    - A value of ``None`` in the overlay deletes the key from the result.
    - Scalars in overlay override base.
    """
    result = dict(base)
    for key, value in overlay.items():
        if value is None:
            result.pop(key, None)
        elif isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def _parse_subsidiary(key: str, raw: dict) -> SubsidiaryConfig:
    _require_keys(raw, _REQUIRED_SUBSIDIARY, f"subsidiaries.{key}")
    _reject_unknown_keys(raw, _ALLOWED_SUBSIDIARY, f"subsidiaries.{key}")
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
    _reject_unknown_keys(raw, _ALLOWED_COMPANY, "company")

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
    _reject_unknown_keys(raw, _ALLOWED_TOP, "config root")

    company = _parse_company(raw["company"])

    aug_raw = raw.get("augmentation") or {}
    augmentation = AugmentationConfig(
        enabled=bool(aug_raw.get("enabled", False)),
        model=str(aug_raw.get("model", "")),
        cache_dir=str(aug_raw.get("cache_dir", ".augmentation_cache")),
        warm_on_miss=bool(aug_raw.get("warm_on_miss", False)),
    )

    difficulty = _parse_profile_section(raw, "difficulty", _parse_difficulty, DifficultyProfile)
    output = _parse_profile_section(raw, "output", _parse_output, OutputProfile)
    errors = _parse_profile_section(raw, "errors", _parse_errors, ErrorProfile)

    return Config(
        seed=int(raw["seed"]),
        output_dir=raw["output_dir"],
        company=company,
        canary_assignments=raw.get("canary_assignments") or {},
        error_injections=raw.get("error_injections") or {},
        augmentation=augmentation,
        difficulty=difficulty,
        output=output,
        errors=errors,
    )


def _parse_profile_section(raw: dict, key: str, parser: Any, default_cls: type) -> Any:
    """Parse an optional profile section, validating it is a mapping if present."""
    if key not in raw:
        return default_cls()
    section = raw[key]
    if not isinstance(section, dict):
        raise ConfigError(f"'{key}' must be a mapping, got {type(section).__name__}")
    return parser(section)


def _load_yaml(path: Path) -> dict:
    """Load a YAML file and return its contents as a dict."""
    if not path.exists():
        raise ConfigError(f"Config file not found: {path}")
    try:
        with open(path) as f:
            raw = yaml.safe_load(f)
    except yaml.YAMLError as exc:
        raise ConfigError(f"Invalid YAML in {path}: {exc}") from exc
    if not isinstance(raw, dict):
        raise ConfigError(f"Config root must be a mapping, got {type(raw).__name__}")
    return raw


def load_layered_config(
    base_path: str | Path,
    layers: list[str | Path] | None = None,
    set_overrides: dict | None = None,
) -> Config:
    """Load a base config and merge zero or more overlay layers on top.

    Each layer is deep-merged onto the base in order (last wins). After
    merging, the combined result is validated and returned as a Config.

    Parameters
    ----------
    base_path : path to the base YAML config file.
    layers : optional list of overlay YAML file paths, merged in order.
    set_overrides : optional dict of leaf-level overrides (e.g. from
        ``--set key=value`` CLI flags).  Applied after file layers.

    Layer files may contain any subset of v1 supported fields.
    """
    base_raw = _load_yaml(Path(base_path))

    if layers:
        for layer_path in layers:
            layer_raw = _load_yaml(Path(layer_path))
            base_raw = deep_merge(base_raw, layer_raw)

    if set_overrides:
        base_raw = deep_merge(base_raw, set_overrides)

    # Validate and parse the merged result through the same path as load_config
    _require_keys(base_raw, _REQUIRED_TOP, "config root (merged)")
    _reject_unknown_keys(base_raw, _ALLOWED_TOP, "config root (merged)")

    company = _parse_company(base_raw["company"])

    aug_raw = base_raw.get("augmentation") or {}
    augmentation = AugmentationConfig(
        enabled=bool(aug_raw.get("enabled", False)),
        model=str(aug_raw.get("model", "")),
        cache_dir=str(aug_raw.get("cache_dir", ".augmentation_cache")),
        warm_on_miss=bool(aug_raw.get("warm_on_miss", False)),
    )

    difficulty = _parse_profile_section(base_raw, "difficulty", _parse_difficulty, DifficultyProfile)
    output_prof = _parse_profile_section(base_raw, "output", _parse_output, OutputProfile)
    errors = _parse_profile_section(base_raw, "errors", _parse_errors, ErrorProfile)

    return Config(
        seed=int(base_raw["seed"]),
        output_dir=base_raw["output_dir"],
        company=company,
        canary_assignments=base_raw.get("canary_assignments") or {},
        error_injections=base_raw.get("error_injections") or {},
        augmentation=augmentation,
        difficulty=difficulty,
        output=output_prof,
        errors=errors,
    )
