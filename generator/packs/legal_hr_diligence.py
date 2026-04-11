"""Legal/HR diligence scenario pack — TC-19 through TC-21.

Depends on cascade_accounting_core (TC-12 data room provides the baseline
company context that legal and HR diligence builds upon).
"""

from __future__ import annotations

from generator.formatters.tc19 import emit_tc19
from generator.formatters.tc20 import emit_tc20
from generator.formatters.tc21 import emit_tc21
from generator.packs import ScenarioPack

# Canary file keys — merged from TC-19 (15) + TC-20 (11) + TC-21 (11) = 37
_CANARY_FILE_KEYS: list[str] = sorted([
    # TC-19: 10 contracts + 3 amendments + 1 memo + 1 request list
    "tc19_contract_lctr_001",
    "tc19_contract_lctr_002",
    "tc19_contract_lctr_003",
    "tc19_contract_lctr_004",
    "tc19_contract_lctr_005",
    "tc19_contract_lctr_006",
    "tc19_contract_lctr_007",
    "tc19_contract_lctr_008",
    "tc19_contract_lctr_009",
    "tc19_contract_lctr_010",
    "tc19_amendment_amd_001",
    "tc19_amendment_amd_002",
    "tc19_amendment_amd_003",
    "tc19_management_summary_memo",
    "tc19_diligence_request_list",
    # TC-20: 7 agreements + 1 census + 1 severance + 1 retention + 1 contractor
    "tc20_agreement_ea_001",
    "tc20_agreement_ea_002",
    "tc20_agreement_ea_003",
    "tc20_agreement_ea_004",
    "tc20_agreement_ea_005",
    "tc20_agreement_ea_006",
    "tc20_agreement_ea_007",
    "tc20_employee_census",
    "tc20_severance_schedule",
    "tc20_retention_plan",
    "tc20_contractor_roster",
    # TC-21: 3 contracts + 2 amendments + 2 agreements + 2 schedules + 1 QA + 1 tracker
    "tc21_contract_lctr_001",
    "tc21_contract_lctr_003",
    "tc21_contract_lctr_004",
    "tc21_amendment_amd_002",
    "tc21_amendment_amd_003",
    "tc21_agreement_ea_001",
    "tc21_agreement_ea_006",
    "tc21_severance_schedule",
    "tc21_retention_schedule",
    "tc21_management_qa_summary",
    "tc21_diligence_request_tracker",
])

_EMITTERS = [
    emit_tc19,
    emit_tc20,
    emit_tc21,
]

PACK = ScenarioPack(
    pack_id="cascade_legal_hr_diligence",
    display_name="Cascade Legal/HR Diligence",
    test_cases=["TC-19", "TC-20", "TC-21"],
    canary_file_keys=_CANARY_FILE_KEYS,
    emitters=_EMITTERS,
    dependencies=["cascade_accounting_core"],
)
