"""Schema metadata for TUI field editing.

Exposes editable config field paths with labels, help text, input types,
choices/ranges, grouping, and validation hints.  Aligned to the v1
customization schema (customization-schema-v1.md) and the dataclasses
in generator/config.py.
"""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum

# ---------------------------------------------------------------------------
# Field metadata types
# ---------------------------------------------------------------------------

class InputType(Enum):
    """The kind of UI widget a field should render as."""
    TEXT = "text"
    INTEGER = "integer"
    FLOAT = "float"
    CHOICE = "choice"
    MULTI_CHOICE = "multi_choice"
    LIST_TEXT = "list_text"
    LIST_INT = "list_int"


class FieldGroup(Enum):
    """Logical grouping of fields for TUI screen layout."""
    COMPANY = "company"
    SUBSIDIARY = "subsidiary"
    FINANCIAL = "financial"
    DIFFICULTY = "difficulty"
    OUTPUT = "output"
    ERRORS = "errors"
    GENERATOR = "generator"


@dataclass(frozen=True)
class FieldMeta:
    """Metadata for a single editable config field."""

    path: str
    """Dot-separated path in the config tree, e.g. 'company.name'."""

    label: str
    """Short, human-readable label for TUI display."""

    help_text: str
    """One-line description shown in the TUI help panel."""

    input_type: InputType
    """Determines the UI widget and value parsing."""

    group: FieldGroup
    """Which TUI screen/section this field belongs to."""

    required: bool = True
    """Whether the field must be present in the final merged config."""

    choices: tuple[str, ...] = ()
    """For CHOICE/MULTI_CHOICE: allowed values."""

    range_min: float | None = None
    """For INTEGER/FLOAT: minimum allowed value (inclusive)."""

    range_max: float | None = None
    """For INTEGER/FLOAT: maximum allowed value (inclusive)."""

    default: object = None
    """Default value used when the field is omitted."""

    template: str = ""
    """For subsidiary fields: path template with '{key}' placeholder."""


# ---------------------------------------------------------------------------
# Field registry
# ---------------------------------------------------------------------------

# Company profile fields
_COMPANY_FIELDS: list[FieldMeta] = [
    FieldMeta(
        path="company.name",
        label="Company Name",
        help_text="Legal name of the parent entity",
        input_type=InputType.TEXT,
        group=FieldGroup.COMPANY,
    ),
    FieldMeta(
        path="company.type",
        label="Entity Type",
        help_text="Corporate structure (e.g. US C-Corporation)",
        input_type=InputType.TEXT,
        group=FieldGroup.COMPANY,
    ),
    FieldMeta(
        path="company.industry",
        label="Industry",
        help_text="Industry description (e.g. Mid-market manufacturer)",
        input_type=InputType.TEXT,
        group=FieldGroup.COMPANY,
    ),
    FieldMeta(
        path="company.headquarters",
        label="Headquarters",
        help_text="City and state of the headquarters",
        input_type=InputType.TEXT,
        group=FieldGroup.COMPANY,
    ),
    FieldMeta(
        path="company.fiscal_year_end",
        label="Fiscal Year End",
        help_text="Month-day of fiscal year end (MM-DD format)",
        input_type=InputType.TEXT,
        group=FieldGroup.COMPANY,
    ),
    FieldMeta(
        path="company.years",
        label="Fiscal Years",
        help_text="Three consecutive fiscal years to generate data for",
        input_type=InputType.LIST_INT,
        group=FieldGroup.COMPANY,
    ),
    FieldMeta(
        path="company.current_year",
        label="Current Year",
        help_text="The most recent fiscal year (must be max of years)",
        input_type=InputType.INTEGER,
        group=FieldGroup.COMPANY,
    ),
    FieldMeta(
        path="company.consolidated_revenue",
        label="Consolidated Revenue",
        help_text="Total consolidated revenue in dollars",
        input_type=InputType.INTEGER,
        group=FieldGroup.COMPANY,
        range_min=1,
    ),
]

# Subsidiary fields — templated with {key} for the subsidiary slug
_SUBSIDIARY_FIELDS: list[FieldMeta] = [
    FieldMeta(
        path="company.subsidiaries.{key}.legal_name",
        label="Legal Name",
        help_text="Full legal name of the subsidiary",
        input_type=InputType.TEXT,
        group=FieldGroup.SUBSIDIARY,
        template="company.subsidiaries.{key}.legal_name",
    ),
    FieldMeta(
        path="company.subsidiaries.{key}.location",
        label="Location",
        help_text="City and state (e.g. Portland, OR)",
        input_type=InputType.TEXT,
        group=FieldGroup.SUBSIDIARY,
        template="company.subsidiaries.{key}.location",
    ),
    FieldMeta(
        path="company.subsidiaries.{key}.state",
        label="State",
        help_text="Two-letter state code",
        input_type=InputType.TEXT,
        group=FieldGroup.SUBSIDIARY,
        template="company.subsidiaries.{key}.state",
    ),
    FieldMeta(
        path="company.subsidiaries.{key}.revenue",
        label="Revenue",
        help_text="Annual revenue in dollars",
        input_type=InputType.INTEGER,
        group=FieldGroup.SUBSIDIARY,
        template="company.subsidiaries.{key}.revenue",
        range_min=0,
    ),
    FieldMeta(
        path="company.subsidiaries.{key}.type",
        label="Business Type",
        help_text="Description of what the subsidiary does",
        input_type=InputType.TEXT,
        group=FieldGroup.SUBSIDIARY,
        template="company.subsidiaries.{key}.type",
    ),
    FieldMeta(
        path="company.subsidiaries.{key}.gross_margin",
        label="Gross Margin",
        help_text="Gross margin as a decimal (0.0–1.0)",
        input_type=InputType.FLOAT,
        group=FieldGroup.SUBSIDIARY,
        template="company.subsidiaries.{key}.gross_margin",
        range_min=0.0,
        range_max=1.0,
    ),
    FieldMeta(
        path="company.subsidiaries.{key}.employee_count",
        label="Employee Count",
        help_text="Number of employees at this subsidiary",
        input_type=InputType.INTEGER,
        group=FieldGroup.SUBSIDIARY,
        template="company.subsidiaries.{key}.employee_count",
        range_min=1,
    ),
    FieldMeta(
        path="company.subsidiaries.{key}.rd_spend_pct",
        label="R&D Spend %",
        help_text="R&D spending as fraction of revenue (0.0–1.0)",
        input_type=InputType.FLOAT,
        group=FieldGroup.SUBSIDIARY,
        template="company.subsidiaries.{key}.rd_spend_pct",
        required=False,
        range_min=0.0,
        range_max=1.0,
        default=0.0,
    ),
]

# Financial parameter fields
_FINANCIAL_FIELDS: list[FieldMeta] = [
    FieldMeta(
        path="company.growth_rates.fy2023_to_fy2024",
        label="FY2023→FY2024 Growth",
        help_text="Year-over-year revenue growth rate",
        input_type=InputType.FLOAT,
        group=FieldGroup.FINANCIAL,
    ),
    FieldMeta(
        path="company.growth_rates.fy2024_to_fy2025",
        label="FY2024→FY2025 Growth",
        help_text="Year-over-year revenue growth rate",
        input_type=InputType.FLOAT,
        group=FieldGroup.FINANCIAL,
    ),
    FieldMeta(
        path="company.intercompany.raw_materials_markup",
        label="Raw Materials Markup",
        help_text="Cost-plus markup on intercompany raw materials",
        input_type=InputType.FLOAT,
        group=FieldGroup.FINANCIAL,
        range_min=0.0,
    ),
    FieldMeta(
        path="company.intercompany.management_fee_pct",
        label="Management Fee %",
        help_text="Parent management fee as fraction of sub revenue",
        input_type=InputType.FLOAT,
        group=FieldGroup.FINANCIAL,
        range_min=0.0,
        range_max=1.0,
    ),
    FieldMeta(
        path="company.intercompany.intercompany_loan_principal",
        label="IC Loan Principal",
        help_text="Intercompany loan principal amount in dollars",
        input_type=InputType.INTEGER,
        group=FieldGroup.FINANCIAL,
        range_min=0,
    ),
    FieldMeta(
        path="company.intercompany.intercompany_loan_rate",
        label="IC Loan Rate",
        help_text="Annual interest rate on intercompany loan",
        input_type=InputType.FLOAT,
        group=FieldGroup.FINANCIAL,
        range_min=0.0,
        range_max=1.0,
    ),
    FieldMeta(
        path="company.employees.total_count",
        label="Total Employees",
        help_text="Company-wide employee headcount",
        input_type=InputType.INTEGER,
        group=FieldGroup.FINANCIAL,
        range_min=1,
    ),
    FieldMeta(
        path="company.employees.annual_turnover_rate",
        label="Turnover Rate",
        help_text="Annual employee turnover as a decimal",
        input_type=InputType.FLOAT,
        group=FieldGroup.FINANCIAL,
        range_min=0.0,
        range_max=1.0,
    ),
    FieldMeta(
        path="company.employees.remote_states",
        label="Remote States",
        help_text="Two-letter state codes where remote employees reside",
        input_type=InputType.LIST_TEXT,
        group=FieldGroup.FINANCIAL,
    ),
    FieldMeta(
        path="company.seasonal_weights.Q1",
        label="Q1 Weight",
        help_text="Fraction of annual revenue in Q1 (all Qs must sum to 1.0)",
        input_type=InputType.FLOAT,
        group=FieldGroup.FINANCIAL,
        range_min=0.0,
        range_max=1.0,
    ),
    FieldMeta(
        path="company.seasonal_weights.Q2",
        label="Q2 Weight",
        help_text="Fraction of annual revenue in Q2",
        input_type=InputType.FLOAT,
        group=FieldGroup.FINANCIAL,
        range_min=0.0,
        range_max=1.0,
    ),
    FieldMeta(
        path="company.seasonal_weights.Q3",
        label="Q3 Weight",
        help_text="Fraction of annual revenue in Q3",
        input_type=InputType.FLOAT,
        group=FieldGroup.FINANCIAL,
        range_min=0.0,
        range_max=1.0,
    ),
    FieldMeta(
        path="company.seasonal_weights.Q4",
        label="Q4 Weight",
        help_text="Fraction of annual revenue in Q4",
        input_type=InputType.FLOAT,
        group=FieldGroup.FINANCIAL,
        range_min=0.0,
        range_max=1.0,
    ),
]

# Difficulty profile fields
_DIFFICULTY_FIELDS: list[FieldMeta] = [
    FieldMeta(
        path="difficulty.error_density",
        label="Error Density",
        help_text="Fraction of possible errors to inject (0.0–1.0)",
        input_type=InputType.FLOAT,
        group=FieldGroup.DIFFICULTY,
        required=False,
        range_min=0.0,
        range_max=1.0,
        default=1.0,
    ),
    FieldMeta(
        path="difficulty.canary_visibility",
        label="Canary Visibility",
        help_text="How visible canary values are in generated files",
        input_type=InputType.CHOICE,
        group=FieldGroup.DIFFICULTY,
        required=False,
        choices=("visible", "subtle", "hidden"),
        default="visible",
    ),
    FieldMeta(
        path="difficulty.judgment_trap_density",
        label="Judgment Trap Density",
        help_text="Fraction of judgment traps to include (0.0–1.0)",
        input_type=InputType.FLOAT,
        group=FieldGroup.DIFFICULTY,
        required=False,
        range_min=0.0,
        range_max=1.0,
        default=1.0,
    ),
]

# Output profile fields
_OUTPUT_FIELDS: list[FieldMeta] = [
    FieldMeta(
        path="output.enabled_test_cases",
        label="Enabled Test Cases",
        help_text="Which TCs to generate (e.g. TC-01, TC-06); empty = all",
        input_type=InputType.MULTI_CHOICE,
        group=FieldGroup.OUTPUT,
        required=False,
        choices=(
            "TC-01", "TC-02", "TC-03", "TC-04", "TC-05", "TC-06",
            "TC-07", "TC-08", "TC-09", "TC-10", "TC-11", "TC-12",
            "TC-13", "TC-14", "TC-15", "TC-16", "TC-17", "TC-18",
        ),
        default=(),
    ),
    FieldMeta(
        path="output.enabled_packs",
        label="Enabled Packs",
        help_text="Scenario packs to generate; empty = all",
        input_type=InputType.LIST_TEXT,
        group=FieldGroup.OUTPUT,
        required=False,
        default=(),
    ),
]

# Error profile fields
_ERROR_FIELDS: list[FieldMeta] = [
    FieldMeta(
        path="errors.include",
        label="Force Include Errors",
        help_text="Error IDs to always inject regardless of density",
        input_type=InputType.LIST_TEXT,
        group=FieldGroup.ERRORS,
        required=False,
        default=(),
    ),
    FieldMeta(
        path="errors.exclude",
        label="Force Exclude Errors",
        help_text="Error IDs to never inject regardless of density",
        input_type=InputType.LIST_TEXT,
        group=FieldGroup.ERRORS,
        required=False,
        default=(),
    ),
    FieldMeta(
        path="errors.density_override",
        label="Error Density Override",
        help_text="Overrides difficulty.error_density (0.0–1.0, or empty for none)",
        input_type=InputType.FLOAT,
        group=FieldGroup.ERRORS,
        required=False,
        range_min=0.0,
        range_max=1.0,
        default=None,
    ),
]

# Generator control fields
_GENERATOR_FIELDS: list[FieldMeta] = [
    FieldMeta(
        path="seed",
        label="Random Seed",
        help_text="Master seed for deterministic generation",
        input_type=InputType.INTEGER,
        group=FieldGroup.GENERATOR,
    ),
    FieldMeta(
        path="output_dir",
        label="Output Directory",
        help_text="Directory path for generated test suite",
        input_type=InputType.TEXT,
        group=FieldGroup.GENERATOR,
    ),
]


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def get_all_fields() -> list[FieldMeta]:
    """Return all non-template field metadata entries."""
    return (
        _GENERATOR_FIELDS
        + _COMPANY_FIELDS
        + _FINANCIAL_FIELDS
        + _DIFFICULTY_FIELDS
        + _OUTPUT_FIELDS
        + _ERROR_FIELDS
    )


def get_subsidiary_fields() -> list[FieldMeta]:
    """Return template field metadata for subsidiary editing.

    Each entry has a ``template`` with ``{key}`` placeholder.  The TUI
    should call :func:`expand_subsidiary_fields` with actual subsidiary
    keys to get concrete paths.
    """
    return list(_SUBSIDIARY_FIELDS)


def expand_subsidiary_fields(sub_key: str) -> list[FieldMeta]:
    """Return concrete field metadata for a specific subsidiary.

    Replaces ``{key}`` in both ``path`` and ``template`` with *sub_key*.
    """
    result = []
    for tmpl in _SUBSIDIARY_FIELDS:
        concrete_path = tmpl.path.replace("{key}", sub_key)
        result.append(FieldMeta(
            path=concrete_path,
            label=tmpl.label,
            help_text=tmpl.help_text,
            input_type=tmpl.input_type,
            group=tmpl.group,
            required=tmpl.required,
            choices=tmpl.choices,
            range_min=tmpl.range_min,
            range_max=tmpl.range_max,
            default=tmpl.default,
            template=concrete_path,
        ))
    return result


def get_fields_by_group(group: FieldGroup) -> list[FieldMeta]:
    """Return all non-template fields belonging to *group*."""
    return [f for f in get_all_fields() if f.group == group]


def get_field_by_path(path: str) -> FieldMeta | None:
    """Look up a field by its exact dot-path.  Returns None if not found."""
    for f in get_all_fields():
        if f.path == path:
            return f
    return None


def is_unsupported_path(path: str) -> bool:
    """Return True if *path* targets a field that is not user-editable in v1.

    The v1 schema explicitly forbids:
    - ``canary_assignments`` — computed from seed
    - ``error_injections`` — derived from error registry + profile
    - ``company.subsidiaries.*.entity_code`` — join key, breaks cross-refs
    """
    if path in ("canary_assignments", "error_injections"):
        return True
    parts = path.split(".")
    if (
        len(parts) >= 4
        and parts[0] == "company"
        and parts[1] == "subsidiaries"
        and parts[3] == "entity_code"
    ):
        return True
    return False
