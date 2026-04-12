"""Tests for generator.schema_metadata — TUI field metadata registry."""

from generator.schema_metadata import (
    FieldGroup,
    InputType,
    expand_subsidiary_fields,
    get_all_fields,
    get_field_by_path,
    get_fields_by_group,
    get_subsidiary_fields,
    is_unsupported_path,
)

# ---------------------------------------------------------------------------
# Company group coverage
# ---------------------------------------------------------------------------

class TestCompanyFields:
    def test_company_name_present(self):
        f = get_field_by_path("company.name")
        assert f is not None
        assert f.label == "Company Name"
        assert f.input_type == InputType.TEXT
        assert f.group == FieldGroup.COMPANY
        assert f.required is True

    def test_consolidated_revenue_has_range(self):
        f = get_field_by_path("company.consolidated_revenue")
        assert f is not None
        assert f.input_type == InputType.INTEGER
        assert f.range_min == 1

    def test_fiscal_years_is_list_int(self):
        f = get_field_by_path("company.years")
        assert f is not None
        assert f.input_type == InputType.LIST_INT

    def test_current_year_is_integer(self):
        f = get_field_by_path("company.current_year")
        assert f is not None
        assert f.input_type == InputType.INTEGER

    def test_all_company_fields_in_group(self):
        fields = get_fields_by_group(FieldGroup.COMPANY)
        paths = {f.path for f in fields}
        expected = {
            "company.name", "company.type", "company.industry",
            "company.headquarters", "company.fiscal_year_end",
            "company.years", "company.current_year",
            "company.consolidated_revenue",
        }
        assert paths == expected


# ---------------------------------------------------------------------------
# Subsidiary group coverage
# ---------------------------------------------------------------------------

class TestSubsidiaryFields:
    def test_templates_have_key_placeholder(self):
        for f in get_subsidiary_fields():
            assert "{key}" in f.path
            assert f.group == FieldGroup.SUBSIDIARY

    def test_expand_replaces_placeholder(self):
        expanded = expand_subsidiary_fields("precision_components")
        for f in expanded:
            assert "{key}" not in f.path
            assert "precision_components" in f.path

    def test_expand_includes_all_required_sub_fields(self):
        expanded = expand_subsidiary_fields("test_sub")
        paths = {f.path for f in expanded}
        assert "company.subsidiaries.test_sub.legal_name" in paths
        assert "company.subsidiaries.test_sub.revenue" in paths
        assert "company.subsidiaries.test_sub.gross_margin" in paths
        assert "company.subsidiaries.test_sub.employee_count" in paths

    def test_rd_spend_optional(self):
        expanded = expand_subsidiary_fields("x")
        rd = [f for f in expanded if f.path.endswith(".rd_spend_pct")]
        assert len(rd) == 1
        assert rd[0].required is False
        assert rd[0].default == 0.0

    def test_gross_margin_range(self):
        expanded = expand_subsidiary_fields("x")
        gm = [f for f in expanded if f.path.endswith(".gross_margin")]
        assert len(gm) == 1
        assert gm[0].range_min == 0.0
        assert gm[0].range_max == 1.0


# ---------------------------------------------------------------------------
# Difficulty / output group coverage
# ---------------------------------------------------------------------------

class TestDifficultyFields:
    def test_error_density_range(self):
        f = get_field_by_path("difficulty.error_density")
        assert f is not None
        assert f.range_min == 0.0
        assert f.range_max == 1.0
        assert f.default == 1.0

    def test_canary_visibility_choices(self):
        f = get_field_by_path("difficulty.canary_visibility")
        assert f is not None
        assert f.input_type == InputType.CHOICE
        assert set(f.choices) == {"visible", "subtle", "hidden"}

    def test_judgment_trap_density(self):
        f = get_field_by_path("difficulty.judgment_trap_density")
        assert f is not None
        assert f.range_min == 0.0
        assert f.range_max == 1.0

    def test_all_difficulty_fields_optional(self):
        fields = get_fields_by_group(FieldGroup.DIFFICULTY)
        assert all(not f.required for f in fields)


class TestOutputFields:
    def test_enabled_test_cases_multi_choice(self):
        f = get_field_by_path("output.enabled_test_cases")
        assert f is not None
        assert f.input_type == InputType.MULTI_CHOICE
        assert "TC-01" in f.choices
        assert "TC-18" in f.choices
        assert len(f.choices) == 18

    def test_enabled_packs_is_list(self):
        f = get_field_by_path("output.enabled_packs")
        assert f is not None
        assert f.input_type == InputType.LIST_TEXT


# ---------------------------------------------------------------------------
# Error profile and TC override group coverage
# ---------------------------------------------------------------------------

class TestErrorFields:
    def test_include_exclude_are_lists(self):
        inc = get_field_by_path("errors.include")
        exc = get_field_by_path("errors.exclude")
        assert inc is not None and inc.input_type == InputType.LIST_TEXT
        assert exc is not None and exc.input_type == InputType.LIST_TEXT

    def test_density_override_range(self):
        f = get_field_by_path("errors.density_override")
        assert f is not None
        assert f.range_min == 0.0
        assert f.range_max == 1.0
        assert f.default is None

    def test_all_error_fields_optional(self):
        fields = get_fields_by_group(FieldGroup.ERRORS)
        assert all(not f.required for f in fields)


# ---------------------------------------------------------------------------
# Generator control
# ---------------------------------------------------------------------------

class TestGeneratorFields:
    def test_seed_present(self):
        f = get_field_by_path("seed")
        assert f is not None
        assert f.input_type == InputType.INTEGER

    def test_output_dir_present(self):
        f = get_field_by_path("output_dir")
        assert f is not None
        assert f.input_type == InputType.TEXT


# ---------------------------------------------------------------------------
# Unsupported path detection
# ---------------------------------------------------------------------------

class TestUnsupportedPaths:
    def test_canary_assignments_unsupported(self):
        assert is_unsupported_path("canary_assignments") is True

    def test_error_injections_unsupported(self):
        assert is_unsupported_path("error_injections") is True

    def test_entity_code_unsupported(self):
        assert is_unsupported_path(
            "company.subsidiaries.precision_components.entity_code"
        ) is True

    def test_company_name_supported(self):
        assert is_unsupported_path("company.name") is False

    def test_difficulty_supported(self):
        assert is_unsupported_path("difficulty.error_density") is False


# ---------------------------------------------------------------------------
# Structural invariants
# ---------------------------------------------------------------------------

class TestStructuralInvariants:
    def test_no_duplicate_paths(self):
        paths = [f.path for f in get_all_fields()]
        assert len(paths) == len(set(paths)), f"Duplicate paths: {paths}"

    def test_all_fields_have_labels(self):
        for f in get_all_fields():
            assert f.label, f"Field {f.path} missing label"

    def test_all_fields_have_help_text(self):
        for f in get_all_fields():
            assert f.help_text, f"Field {f.path} missing help_text"

    def test_choice_fields_have_choices(self):
        for f in get_all_fields():
            if f.input_type in (InputType.CHOICE, InputType.MULTI_CHOICE):
                assert f.choices, f"Field {f.path} is {f.input_type} but has no choices"

    def test_every_group_has_fields(self):
        for group in FieldGroup:
            if group == FieldGroup.SUBSIDIARY:
                # Subsidiary fields are templates, accessed via get_subsidiary_fields
                assert len(get_subsidiary_fields()) > 0
            else:
                assert len(get_fields_by_group(group)) > 0, (
                    f"FieldGroup.{group.name} has no fields"
                )
