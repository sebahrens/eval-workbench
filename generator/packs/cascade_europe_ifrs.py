"""Cascade Europe IFRS scenario pack — European variants.

Depends on cascade_accounting_core (the base Cascade model provides entity
context that European variants build upon).
"""

from __future__ import annotations

from generator.formatters.tc04_eu import emit_tc04_eu
from generator.formatters.tc06_eu import emit_tc06_eu
from generator.formatters.tc07_eu import emit_tc07_eu
from generator.formatters.tc08_eu import emit_tc08_eu
from generator.formatters.tc09_eu import emit_tc09_eu
from generator.formatters.tc10_eu import emit_tc10_eu
from generator.formatters.tc12_eu import emit_tc12_eu
from generator.formatters.tc16_eu import emit_tc16_eu
from generator.packs import ScenarioPack

_EMITTERS = [
    emit_tc04_eu,
    emit_tc06_eu,
    emit_tc07_eu,
    emit_tc08_eu,
    emit_tc09_eu,
    emit_tc10_eu,
    emit_tc12_eu,
    emit_tc16_eu,
]

# Canary file keys — TC-04-EU: 15 lease PDFs + 1 schedule = 16; TC-06-EU: 4 files;
# TC-07-EU: 8 allocation PDFs + 1 investment register + 1 WHT summary = 10;
# TC-08-EU: 1 CSV + 14 DOCX + 3 XLSX + 1 XLSX = 19; TC-09-EU: 3 XLSX + 2 PDF = 5;
# TC-10-EU: 3 XLSX + 1 DOCX = 4; TC-12-EU: 36 data room files + 1 checklist = 37;
# TC-16-EU: 1 DOCX + 1 XLSX + 1 DOCX = 3
_CANARY_FILE_KEYS: list[str] = sorted([
    # TC-04-EU
    "tc04eu_lease_001",
    "tc04eu_lease_002",
    "tc04eu_lease_003",
    "tc04eu_lease_004",
    "tc04eu_lease_005",
    "tc04eu_lease_006",
    "tc04eu_lease_007",
    "tc04eu_lease_008",
    "tc04eu_lease_009",
    "tc04eu_lease_010",
    "tc04eu_lease_011",
    "tc04eu_lease_012",
    "tc04eu_lease_013",
    "tc04eu_lease_014",
    "tc04eu_lease_015",
    "tc04eu_lease_schedule_partial",
    # TC-06-EU
    "tc06eu_consolidated_tb_fy2025",
    "tc06eu_tax_provision_fy2024_workpaper",
    "tc06eu_perm_temp_differences_fy2025",
    "tc06eu_statutory_rates",
    # TC-07-EU
    "tc07eu_alloc_001",
    "tc07eu_alloc_002",
    "tc07eu_alloc_003",
    "tc07eu_alloc_004",
    "tc07eu_alloc_005",
    "tc07eu_alloc_006",
    "tc07eu_alloc_007",
    "tc07eu_alloc_008",
    "tc07eu_investment_register",
    "tc07eu_wht_summary",
    # TC-08-EU
    "tc08eu_time_records",
    "tc08eu_payroll",
    "tc08eu_supply_expenses",
    "tc08eu_subcontractor_invoices",
    "tc08eu_prior_year_rd",
    "tc08eu_project_001",
    "tc08eu_project_002",
    "tc08eu_project_003",
    "tc08eu_project_004",
    "tc08eu_project_005",
    "tc08eu_project_006",
    "tc08eu_project_007",
    "tc08eu_project_008",
    "tc08eu_project_009",
    "tc08eu_project_010",
    "tc08eu_project_011",
    "tc08eu_project_012",
    "tc08eu_project_013",
    "tc08eu_project_014",
    # TC-09-EU
    "tc09eu_comparable_companies",
    "tc09eu_intercompany_transactions",
    "tc09eu_interest_rate_benchmarks",
    "tc09eu_local_file_cp_fy2024",
    "tc09eu_master_file_fy2024",
    # TC-10-EU
    "tc10eu_eu_vat_rules_reference",
    "tc10eu_intercompany_sales_fy2025",
    "tc10eu_vat_registrations",
    "tc10eu_vat_returns_summary_fy2025",
    # TC-12-EU (36 data room files + 1 checklist = 37 canary keys)
    "tc12eu_benefits_summary_by_country",
    "tc12eu_cd_companies_house_filing",
    "tc12eu_cd_ct600_fy2024",
    "tc12eu_ce_articles_of_association",
    "tc12eu_ce_board_minutes_2024",
    "tc12eu_ce_board_minutes_2025",
    "tc12eu_ce_kvk_extract",
    "tc12eu_ce_vpb_aangifte_fy2024",
    "tc12eu_cm_kbis_extract",
    "tc12eu_cm_liasse_fiscale_fy2024",
    "tc12eu_cp_handelsregister_auszug",
    "tc12eu_cp_korperschaftsteuer_fy2023",
    "tc12eu_cp_korperschaftsteuer_fy2024",
    "tc12eu_customer_agreement_autohaus",
    "tc12eu_customer_agreement_renault",
    "tc12eu_customer_list_with_revenue",
    "tc12eu_dd_checklist_european",
    "tc12eu_equipment_list",
    "tc12eu_facility_leases",
    "tc12eu_finance_director_agreement_ce",
    "tc12eu_gdpr_data_processing_register",
    "tc12eu_group_budget_fy2025",
    "tc12eu_group_debt_schedule",
    "tc12eu_group_employee_census",
    "tc12eu_group_ifrs_financials_fy2023",
    "tc12eu_group_ifrs_financials_fy2024",
    "tc12eu_group_management_accounts_ytd",
    "tc12eu_group_org_chart",
    "tc12eu_group_org_chart_detailed",
    "tc12eu_group_vat_returns_summary",
    "tc12eu_insurance_policies_summary",
    "tc12eu_ip_assignment_agreements",
    "tc12eu_it_infrastructure_overview",
    "tc12eu_managing_director_agreement_ce",
    "tc12eu_patent_portfolio",
    "tc12eu_pending_litigation_summary",
    "tc12eu_software_licenses",
    "tc12eu_supplier_agreement_thyssenkrupp",
    "tc12eu_tax_notices_and_assessments",
    "tc12eu_technical_director_agreement_cm",
    "tc12eu_vendor_list",
    "tc12eu_works_council_agreements_cp",
    # TC-16-EU (3 input files = 3 canary keys)
    "tc16eu_client_profile",
    "tc16eu_engagement_template",
    "tc16eu_fee_schedule",
])

PACK = ScenarioPack(
    pack_id="cascade_europe_ifrs",
    display_name="Cascade Europe -- IFRS/OECD Variants",
    test_cases=["TC-04-EU", "TC-06-EU", "TC-07-EU", "TC-08-EU", "TC-09-EU", "TC-10-EU", "TC-12-EU", "TC-16-EU"],
    canary_file_keys=_CANARY_FILE_KEYS,
    emitters=_EMITTERS,
    dependencies=["cascade_accounting_core"],
)
