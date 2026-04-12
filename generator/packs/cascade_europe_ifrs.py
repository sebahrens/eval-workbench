"""Cascade Europe IFRS scenario pack — European variants.

Depends on cascade_accounting_core (the base Cascade model provides entity
context that European variants build upon).
"""

from __future__ import annotations

from generator.formatters.tc04_eu import emit_tc04_eu
from generator.formatters.tc06_eu import emit_tc06_eu
from generator.formatters.tc07_eu import emit_tc07_eu
from generator.packs import ScenarioPack

_EMITTERS = [
    emit_tc04_eu,
    emit_tc06_eu,
    emit_tc07_eu,
]

# Canary file keys — TC-04-EU: 15 lease PDFs + 1 schedule = 16; TC-06-EU: 4 files;
# TC-07-EU: 8 allocation PDFs + 1 investment register + 1 WHT summary = 10
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
])

PACK = ScenarioPack(
    pack_id="cascade_europe_ifrs",
    display_name="Cascade Europe -- IFRS/OECD Variants",
    test_cases=["TC-04-EU", "TC-06-EU", "TC-07-EU"],
    canary_file_keys=_CANARY_FILE_KEYS,
    emitters=_EMITTERS,
    dependencies=["cascade_accounting_core"],
)
