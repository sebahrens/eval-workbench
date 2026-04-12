"""Tests for the company profile configuration screen.

Covers: valid and invalid company/subsidiary/financial updates via
the DraftService, field grouping, subsidiary template expansion,
unsupported-field rejection, and the screen-level helper.

Bead: synth-data-2u6.6.5
"""

from __future__ import annotations

import textwrap
from pathlib import Path

import pytest

from generator.schema_metadata import (
    FieldGroup,
    InputType,
    expand_subsidiary_fields,
    get_fields_by_group,
)
from generator.tui.draft_service import DraftService

# ---------------------------------------------------------------------------
# Fixtures
# ---------------------------------------------------------------------------


def _base_config_yaml() -> str:
    """Minimal valid base config with company/subsidiary/financial sections."""
    return textwrap.dedent("""\
        seed: 42
        output_dir: out
        company:
          name: TestCo
          type: C-Corp
          industry: Manufacturing
          headquarters: Portland, OR
          fiscal_year_end: "12-31"
          years: [2023, 2024, 2025]
          current_year: 2025
          consolidated_revenue: 100000000
          subsidiaries:
            sub_a:
              legal_name: Sub A LLC
              location: Portland, OR
              state: OR
              entity_code: SA
              revenue: 60000000
              type: Manufacturing
              gross_margin: 0.35
              employee_count: 200
            sub_b:
              legal_name: Sub B Inc.
              location: Austin, TX
              state: TX
              entity_code: SB
              revenue: 40000000
              type: R&D
              gross_margin: 0.50
              employee_count: 100
              rd_spend_pct: 0.10
          growth_rates:
            fy2023_to_fy2024: 0.06
            fy2024_to_fy2025: 0.09
          intercompany:
            raw_materials_markup: 0.08
            management_fee_pct: 0.015
            intercompany_loan_principal: 5000000
            intercompany_loan_rate: 0.05
          employees:
            total_count: 300
            annual_turnover_rate: 0.08
            remote_states: [CA, WA]
          seasonal_weights:
            Q1: 0.20
            Q2: 0.25
            Q3: 0.25
            Q4: 0.30
        difficulty:
          error_density: 1.0
          canary_visibility: visible
          judgment_trap_density: 1.0
        output:
          enabled_test_cases: []
          enabled_packs: []
    """)


@pytest.fixture()
def base_yaml(tmp_path: Path) -> Path:
    p = tmp_path / "config.yaml"
    p.write_text(_base_config_yaml())
    return p


@pytest.fixture()
def service(base_yaml: Path) -> DraftService:
    return DraftService(base_yaml)


# ---------------------------------------------------------------------------
# Company field metadata tests
# ---------------------------------------------------------------------------


class TestCompanyFieldMetadata:
    """Verify company group fields are registered correctly."""

    def test_company_fields_count(self) -> None:
        fields = get_fields_by_group(FieldGroup.COMPANY)
        assert len(fields) == 8

    def test_company_name_meta(self) -> None:
        fields = get_fields_by_group(FieldGroup.COMPANY)
        by_path = {f.path: f for f in fields}
        f = by_path["company.name"]
        assert f.input_type == InputType.TEXT
        assert f.required

    def test_fiscal_years_meta(self) -> None:
        fields = get_fields_by_group(FieldGroup.COMPANY)
        by_path = {f.path: f for f in fields}
        f = by_path["company.years"]
        assert f.input_type == InputType.LIST_INT

    def test_current_year_meta(self) -> None:
        fields = get_fields_by_group(FieldGroup.COMPANY)
        by_path = {f.path: f for f in fields}
        f = by_path["company.current_year"]
        assert f.input_type == InputType.INTEGER

    def test_consolidated_revenue_meta(self) -> None:
        fields = get_fields_by_group(FieldGroup.COMPANY)
        by_path = {f.path: f for f in fields}
        f = by_path["company.consolidated_revenue"]
        assert f.input_type == InputType.INTEGER
        assert f.range_min == 1


# ---------------------------------------------------------------------------
# Subsidiary field metadata tests
# ---------------------------------------------------------------------------


class TestSubsidiaryFieldMetadata:
    """Verify subsidiary template fields expand correctly."""

    def test_expand_subsidiary_fields(self) -> None:
        fields = expand_subsidiary_fields("sub_a")
        paths = [f.path for f in fields]
        assert "company.subsidiaries.sub_a.legal_name" in paths
        assert "company.subsidiaries.sub_a.revenue" in paths
        assert "company.subsidiaries.sub_a.gross_margin" in paths
        assert "company.subsidiaries.sub_a.employee_count" in paths

    def test_expanded_fields_have_correct_group(self) -> None:
        fields = expand_subsidiary_fields("sub_a")
        for f in fields:
            assert f.group == FieldGroup.SUBSIDIARY

    def test_expand_different_keys_produce_different_paths(self) -> None:
        fields_a = expand_subsidiary_fields("sub_a")
        fields_b = expand_subsidiary_fields("sub_b")
        paths_a = {f.path for f in fields_a}
        paths_b = {f.path for f in fields_b}
        assert paths_a.isdisjoint(paths_b)


# ---------------------------------------------------------------------------
# Financial field metadata tests
# ---------------------------------------------------------------------------


class TestFinancialFieldMetadata:
    """Verify financial group fields are registered correctly."""

    def test_financial_fields_count(self) -> None:
        fields = get_fields_by_group(FieldGroup.FINANCIAL)
        assert len(fields) == 13

    def test_growth_rate_fields_present(self) -> None:
        fields = get_fields_by_group(FieldGroup.FINANCIAL)
        paths = {f.path for f in fields}
        assert "company.growth_rates.fy2023_to_fy2024" in paths
        assert "company.growth_rates.fy2024_to_fy2025" in paths

    def test_seasonal_weight_fields_present(self) -> None:
        fields = get_fields_by_group(FieldGroup.FINANCIAL)
        paths = {f.path for f in fields}
        for q in ("Q1", "Q2", "Q3", "Q4"):
            assert f"company.seasonal_weights.{q}" in paths


# ---------------------------------------------------------------------------
# Valid company profile updates
# ---------------------------------------------------------------------------


class TestValidCompanyUpdates:
    """Service accepts valid company field changes."""

    def test_set_company_name(self, service: DraftService) -> None:
        service.set_field("company.name", "New Corp, Inc.")
        assert service.get_field_value("company.name") == "New Corp, Inc."
        assert service.is_modified("company.name")

    def test_set_current_year(self, service: DraftService) -> None:
        service.set_field("company.current_year", 2024)
        assert service.get_field_value("company.current_year") == 2024

    def test_set_current_year_from_string(self, service: DraftService) -> None:
        service.set_field("company.current_year", "2024")
        assert service.get_field_value("company.current_year") == 2024

    def test_set_consolidated_revenue(self, service: DraftService) -> None:
        service.set_field("company.consolidated_revenue", 250000000)
        assert service.get_field_value("company.consolidated_revenue") == 250000000

    def test_set_fiscal_year_end(self, service: DraftService) -> None:
        service.set_field("company.fiscal_year_end", "06-30")
        assert service.get_field_value("company.fiscal_year_end") == "06-30"

    def test_set_years_list(self, service: DraftService) -> None:
        service.set_field("company.years", [2024, 2025, 2026])
        assert service.get_field_value("company.years") == [2024, 2025, 2026]

    def test_company_diff_tracks_changes(self, service: DraftService) -> None:
        service.set_field("company.name", "Changed Inc.")
        diffs = service.diff()
        paths = [d.path for d in diffs]
        assert "company.name" in paths

    def test_reset_company_field(self, service: DraftService) -> None:
        service.set_field("company.name", "Changed Inc.")
        service.reset_field("company.name")
        assert not service.is_modified("company.name")
        assert service.get_field_value("company.name") == "TestCo"


# ---------------------------------------------------------------------------
# Invalid company profile updates
# ---------------------------------------------------------------------------


class TestInvalidCompanyUpdates:
    """Service rejects invalid company field values."""

    def test_consolidated_revenue_below_min(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="below minimum"):
            service.set_field("company.consolidated_revenue", 0)

    def test_entity_code_rejected(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="not user-editable"):
            service.set_field("company.subsidiaries.sub_a.entity_code", "XX")


# ---------------------------------------------------------------------------
# Valid subsidiary profile updates
# ---------------------------------------------------------------------------


class TestValidSubsidiaryUpdates:
    """Service accepts valid subsidiary field changes."""

    def test_set_subsidiary_legal_name(self, service: DraftService) -> None:
        service.set_field("company.subsidiaries.sub_a.legal_name", "New Sub LLC")
        assert (
            service.get_field_value("company.subsidiaries.sub_a.legal_name")
            == "New Sub LLC"
        )

    def test_set_subsidiary_revenue(self, service: DraftService) -> None:
        service.set_field("company.subsidiaries.sub_a.revenue", 80000000)
        assert service.get_field_value("company.subsidiaries.sub_a.revenue") == 80000000

    def test_set_subsidiary_gross_margin(self, service: DraftService) -> None:
        service.set_field("company.subsidiaries.sub_b.gross_margin", 0.45)
        assert service.get_field_value("company.subsidiaries.sub_b.gross_margin") == 0.45

    def test_set_subsidiary_employee_count(self, service: DraftService) -> None:
        service.set_field("company.subsidiaries.sub_a.employee_count", 500)
        assert service.get_field_value("company.subsidiaries.sub_a.employee_count") == 500

    def test_set_subsidiary_rd_spend_pct(self, service: DraftService) -> None:
        service.set_field("company.subsidiaries.sub_b.rd_spend_pct", 0.15)
        assert service.get_field_value("company.subsidiaries.sub_b.rd_spend_pct") == 0.15

    def test_subsidiary_diff_tracks_changes(self, service: DraftService) -> None:
        service.set_field("company.subsidiaries.sub_a.revenue", 70000000)
        diffs = service.diff()
        paths = [d.path for d in diffs]
        assert "company.subsidiaries.sub_a.revenue" in paths

    def test_reset_subsidiary_field(self, service: DraftService) -> None:
        service.set_field("company.subsidiaries.sub_a.revenue", 70000000)
        service.reset_field("company.subsidiaries.sub_a.revenue")
        assert not service.is_modified("company.subsidiaries.sub_a.revenue")
        assert service.get_field_value("company.subsidiaries.sub_a.revenue") == 60000000


# ---------------------------------------------------------------------------
# Valid financial profile updates
# ---------------------------------------------------------------------------


class TestValidFinancialUpdates:
    """Service accepts valid financial field changes."""

    def test_set_growth_rate(self, service: DraftService) -> None:
        service.set_field("company.growth_rates.fy2023_to_fy2024", 0.10)
        assert service.get_field_value("company.growth_rates.fy2023_to_fy2024") == 0.10

    def test_set_management_fee_pct(self, service: DraftService) -> None:
        service.set_field("company.intercompany.management_fee_pct", 0.02)
        assert service.get_field_value("company.intercompany.management_fee_pct") == 0.02

    def test_set_seasonal_weight(self, service: DraftService) -> None:
        service.set_field("company.seasonal_weights.Q1", 0.25)
        assert service.get_field_value("company.seasonal_weights.Q1") == 0.25

    def test_set_total_employees(self, service: DraftService) -> None:
        service.set_field("company.employees.total_count", 1000)
        assert service.get_field_value("company.employees.total_count") == 1000

    def test_set_remote_states(self, service: DraftService) -> None:
        service.set_field("company.employees.remote_states", ["CA", "WA", "NY", "TX"])
        assert service.get_field_value("company.employees.remote_states") == [
            "CA", "WA", "NY", "TX"
        ]


# ---------------------------------------------------------------------------
# Invalid financial profile updates
# ---------------------------------------------------------------------------


class TestInvalidFinancialUpdates:
    """Service rejects invalid financial field values."""

    def test_management_fee_above_max(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="above maximum"):
            service.set_field("company.intercompany.management_fee_pct", 1.5)

    def test_seasonal_weight_above_max(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="above maximum"):
            service.set_field("company.seasonal_weights.Q1", 1.1)

    def test_total_employees_below_min(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="below minimum"):
            service.set_field("company.employees.total_count", 0)

    def test_turnover_rate_above_max(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="above maximum"):
            service.set_field("company.employees.annual_turnover_rate", 1.5)

    def test_subsidiary_fields_accept_any_value_without_registry(
        self, service: DraftService
    ) -> None:
        """Subsidiary paths are dynamic (templated), so the static registry
        doesn't contain concrete paths like 'company.subsidiaries.sub_a.gross_margin'.
        The DraftService accepts these without per-field range validation —
        full validation happens through the config loader on save/validate."""
        service.set_field("company.subsidiaries.sub_a.gross_margin", 1.5)
        assert service.get_field_value("company.subsidiaries.sub_a.gross_margin") == 1.5


# ---------------------------------------------------------------------------
# Unsupported field rejection
# ---------------------------------------------------------------------------


class TestUnsupportedFieldRejection:
    """Unsupported paths are rejected by the service."""

    def test_canary_assignments_rejected(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="not user-editable"):
            service.set_field("canary_assignments", {"x": "y"})

    def test_error_injections_rejected(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="not user-editable"):
            service.set_field("error_injections", [])

    def test_entity_code_rejected(self, service: DraftService) -> None:
        with pytest.raises(ValueError, match="not user-editable"):
            service.set_field("company.subsidiaries.sub_a.entity_code", "XX")


# ---------------------------------------------------------------------------
# Screen-level helper tests
# ---------------------------------------------------------------------------


class TestScreenHelpers:
    """Test the is_unsupported_company_path helper."""

    def test_company_path_supported(self) -> None:
        from generator.tui.company_profile_screen import is_unsupported_company_path

        assert not is_unsupported_company_path("company.name")

    def test_financial_path_supported(self) -> None:
        from generator.tui.company_profile_screen import is_unsupported_company_path

        assert not is_unsupported_company_path("company.growth_rates.fy2023_to_fy2024")

    def test_difficulty_path_unsupported_in_this_screen(self) -> None:
        from generator.tui.company_profile_screen import is_unsupported_company_path

        assert is_unsupported_company_path("difficulty.error_density")

    def test_output_path_unsupported_in_this_screen(self) -> None:
        from generator.tui.company_profile_screen import is_unsupported_company_path

        assert is_unsupported_company_path("output.enabled_test_cases")

    def test_computed_path_unsupported(self) -> None:
        from generator.tui.company_profile_screen import is_unsupported_company_path

        assert is_unsupported_company_path("canary_assignments")

    def test_unknown_path_unsupported(self) -> None:
        from generator.tui.company_profile_screen import is_unsupported_company_path

        assert is_unsupported_company_path("foo.bar.baz")


# ---------------------------------------------------------------------------
# Combined company profile workflow
# ---------------------------------------------------------------------------


class TestCombinedCompanyWorkflow:
    """End-to-end workflow: set company + subsidiary + financial, validate, diff, save."""

    def test_set_multiple_then_validate(self, service: DraftService) -> None:
        service.set_field("company.name", "Acme Corp")
        service.set_field("company.subsidiaries.sub_a.revenue", 70000000)
        service.set_field("company.growth_rates.fy2023_to_fy2024", 0.08)
        result = service.validate()
        assert result.valid

    def test_set_multiple_then_diff(self, service: DraftService) -> None:
        service.set_field("company.name", "Acme Corp")
        service.set_field("company.subsidiaries.sub_b.gross_margin", 0.60)
        diffs = service.diff()
        paths = {d.path for d in diffs}
        assert "company.name" in paths
        assert "company.subsidiaries.sub_b.gross_margin" in paths

    def test_set_save_reload_roundtrip(
        self, base_yaml: Path, service: DraftService, tmp_path: Path
    ) -> None:
        service.set_field("company.name", "Roundtrip Corp")
        service.set_field("company.consolidated_revenue", 300000000)
        service.set_field("company.subsidiaries.sub_a.legal_name", "RT Sub A")
        service.set_field("company.seasonal_weights.Q4", 0.35)

        out = tmp_path / "override.yaml"
        service.save(out)

        svc2 = DraftService(base_yaml)
        svc2.load_draft(out)
        assert svc2.get_field_value("company.name") == "Roundtrip Corp"
        assert svc2.get_field_value("company.consolidated_revenue") == 300000000
        assert (
            svc2.get_field_value("company.subsidiaries.sub_a.legal_name") == "RT Sub A"
        )
        assert svc2.get_field_value("company.seasonal_weights.Q4") == 0.35

    def test_reset_all_company_fields(self, service: DraftService) -> None:
        service.set_field("company.name", "Changed")
        service.set_field("company.subsidiaries.sub_a.revenue", 99999)

        service.reset_field("company.name")
        service.reset_field("company.subsidiaries.sub_a.revenue")

        assert not service.is_modified("company.name")
        assert not service.is_modified("company.subsidiaries.sub_a.revenue")
        assert service.get_field_value("company.name") == "TestCo"
        assert service.get_field_value("company.subsidiaries.sub_a.revenue") == 60000000
