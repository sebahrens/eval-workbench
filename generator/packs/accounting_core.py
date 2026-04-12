"""Cascade accounting-core scenario pack — TC-01 through TC-18."""

from __future__ import annotations

from generator.formatters.tc01 import emit_tc01
from generator.formatters.tc02 import emit_tc02
from generator.formatters.tc03 import emit_tc03
from generator.formatters.tc04 import emit_tc04
from generator.formatters.tc05 import emit_tc05
from generator.formatters.tc06 import emit_tc06
from generator.formatters.tc07 import emit_tc07
from generator.formatters.tc08 import emit_tc08
from generator.formatters.tc09 import emit_tc09
from generator.formatters.tc10 import emit_tc10
from generator.formatters.tc11 import emit_tc11
from generator.formatters.tc12 import emit_tc12
from generator.formatters.tc13 import emit_tc13
from generator.formatters.tc14 import emit_tc14
from generator.formatters.tc15 import emit_tc15
from generator.formatters.tc16 import emit_tc16
from generator.formatters.tc17 import emit_tc17
from generator.formatters.tc18 import emit_tc18
from generator.formatters.templates import emit_templates as _emit_templates_raw
from generator.packs import ScenarioPack


def _emit_templates_compat(
    model: object,
    output_dir: object,
    canaries: object,
    errors: object,
    manifest: object,
    **kwargs: object,
) -> None:
    """Adapter: emit_templates ignores model and errors."""
    _emit_templates_raw(output_dir, canaries, manifest)  # type: ignore[arg-type]


# Emitters in the exact order currently used by generate_test_suite.py.
# emit_templates is placed between tc16 and tc17 to match the existing
# execution sequence (TC-17 depends on templates being written first).
_EMITTERS = [
    emit_tc01,
    emit_tc02,
    emit_tc03,
    emit_tc04,
    emit_tc05,
    emit_tc06,
    emit_tc07,
    emit_tc08,
    emit_tc09,
    emit_tc10,
    emit_tc11,
    emit_tc12,
    emit_tc13,
    emit_tc14,
    emit_tc15,
    emit_tc16,
    _emit_templates_compat,
    emit_tc17,
    emit_tc18,
]

# File keys that need canaries — sorted for deterministic registry.
_CANARY_FILE_KEYS: list[str] = sorted([
    "cascade_tb_fy2025",
    "cascade_tb_fy2024_workpaper",
    "cascade_financials_fy2024_signed",
    "ar_aging_fy2025",
    "ar_confirmations_summary",
    "allowance_analysis",
    "workpaper_memo_template",
    # TC-03 files
    "tc03_revenue_by_product",
    "tc03_industry_benchmark",
    "tc03_mgmt_rep_letter",
    # TC-02 files
    "bank_confirmation_fy2025",
    "bank_statement_dec2025",
    "cascade_gl_cash_dec2025",
    # TC-06 files
    "cascade_consolidated_tb_fy2025",
    "tax_provision_fy2024_workpaper",
    "perm_temp_differences_fy2025",
    "statutory_rates",
    # TC-04 files (15 lease PDFs + partial schedule)
    "tc04_lease_001",
    "tc04_lease_002",
    "tc04_lease_003",
    "tc04_lease_004",
    "tc04_lease_005",
    "tc04_lease_006",
    "tc04_lease_007",
    "tc04_lease_008",
    "tc04_lease_009",
    "tc04_lease_010",
    "tc04_lease_011",
    "tc04_lease_012",
    "tc04_lease_013",
    "tc04_lease_014",
    "tc04_lease_015",
    "tc04_lease_schedule_partial",
    # TC-07 files (8 K-1 PDFs + org chart)
    "tc07_k1_001",
    "tc07_k1_002",
    "tc07_k1_003",
    "tc07_k1_004",
    "tc07_k1_005",
    "tc07_k1_006",
    "tc07_k1_007",
    "tc07_k1_008",
    "tc07_entity_org_chart",
    # TC-09 files (1 IC transactions xlsx + 1 comparables xlsx + 1 TP report pdf)
    "tc09_ic_transactions",
    "tc09_comparable_companies",
    "tc09_tp_report_fy2024",
    # TC-08 files (1 CSV + 12 project docx + 1 payroll xlsx + 1 supply xlsx)
    "tc08_payroll_data_fy2025",
    "tc08_rd_employee_time_records",
    "tc08_rd_project_001",
    "tc08_rd_project_002",
    "tc08_rd_project_003",
    "tc08_rd_project_004",
    "tc08_rd_project_005",
    "tc08_rd_project_006",
    "tc08_rd_project_007",
    "tc08_rd_project_008",
    "tc08_rd_project_009",
    "tc08_rd_project_010",
    "tc08_rd_project_011",
    "tc08_rd_project_012",
    "tc08_rd_supply_expenses",
    # TC-10 files (1 consolidated P&L + 1 state factors + 1 rules docx)
    "tc10_consolidated_pl",
    "tc10_state_factors",
    "tc10_apportionment_rules",
    # TC-11 files (1 P&L xlsx + 1 adjustments xlsx + 8 contract PDFs + 1 docx)
    "tc11_monthly_pl",
    "tc11_mgmt_adjustments",
    "tc11_contract_001",
    "tc11_contract_002",
    "tc11_contract_003",
    "tc11_contract_004",
    "tc11_contract_005",
    "tc11_contract_006",
    "tc11_contract_007",
    "tc11_contract_008",
    "tc11_interview_notes",
    # TC-12 data room files (32) + DD checklist (1)
    "tc12_articles_of_incorporation",
    "tc12_audited_financials_fy2023",
    "tc12_audited_financials_fy2024",
    "tc12_benefits_summary",
    "tc12_board_minutes_2024",
    "tc12_board_minutes_2025",
    "tc12_budget_fy2025",
    "tc12_ceo_employment_agreement",
    "tc12_cfo_employment_agreement",
    "tc12_cto_employment_agreement",
    "tc12_customer_agreement_acme",
    "tc12_customer_agreement_globex",
    "tc12_customer_list_with_revenue",
    "tc12_bylaws",
    "tc12_dd_checklist",
    "tc12_debt_schedule",
    "tc12_employee_census",
    "tc12_equipment_list",
    "tc12_facility_leases",
    "tc12_federal_returns_fy2023",
    "tc12_federal_returns_fy2024",
    "tc12_insurance_policies_summary",
    "tc12_ip_assignment_agreements",
    "tc12_it_infrastructure_overview",
    "tc12_management_financials_fy2025",
    "tc12_org_chart",
    "tc12_org_chart_detailed",
    "tc12_patent_portfolio",
    "tc12_pending_litigation_summary",
    "tc12_software_licenses",
    "tc12_state_returns_summary",
    "tc12_supplier_agreement_initech",
    "tc12_tax_notices",
    "tc12_vendor_list",
    # TC-13 files (1 CSV)
    "tc13_ap_transactions",
    # TC-14 files (3 xlsx + 1 docx)
    "tc14_balance_sheet_current",
    "tc14_ar_aging_report",
    "tc14_ap_aging_report",
    "tc14_committed_expenses",
    # TC-15 files (3 xlsx + 1 pdf)
    "tc15_historical_financials",
    "tc15_management_projections",
    "tc15_comparable_companies",
    "tc15_industry_overview",
    # TC-16 files (1 client profile docx + 1 fee schedule xlsx + 1 template docx)
    "tc16_client_profile",
    "tc16_fee_schedule",
    "tc16_engagement_template",
    # TC-17 files (4 docx + 2 xlsx workpaper sections)
    "tc17_executive_summary",
    "tc17_financial_analysis",
    "tc17_industry_overview",
    "tc17_risk_assessment",
    "tc17_detailed_findings",
    "tc17_recommendations",
    # Template files (used by TC-17)
    "cover_page_template",
    "formatting_guide",
    # TC-18 files (6 prior-year xlsx + 4 prior-year docx + 5 current-year)
    "tc18_wp_revenue_fy2024",
    "tc18_wp_expenses_fy2024",
    "tc18_wp_balance_sheet_fy2024",
    "tc18_wp_cash_fy2024",
    "tc18_wp_fixed_assets_fy2024",
    "tc18_wp_leases_fy2024",
    "tc18_memo_planning_fy2024",
    "tc18_memo_risk_assessment_fy2024",
    "tc18_memo_summary_fy2024",
    "tc18_memo_management_letter_fy2024",
    "tc18_cy_trial_balance_fy2025",
    "tc18_cy_bank_statements_fy2025",
    "tc18_cy_lease_schedule_fy2025",
    "tc18_cy_mgmt_projections_fy2025",
    "tc18_cy_goodwill_impairment_fy2025",
])

PACK = ScenarioPack(
    pack_id="cascade_accounting_core",
    display_name="Cascade Accounting Core",
    test_cases=[f"TC-{i:02d}" for i in range(1, 19)],
    canary_file_keys=_CANARY_FILE_KEYS,
    emitters=_EMITTERS,
    dependencies=[],
)
