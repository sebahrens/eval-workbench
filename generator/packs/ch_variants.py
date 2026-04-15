"""Swiss (CH) scenario pack — TC-23+.

Swiss-jurisdiction test cases for Cascade Precision Instruments AG,
a Zurich-based subsidiary of Cascade Industries.
"""

from __future__ import annotations

from generator.formatters.tc23 import emit_tc23
from generator.packs import ScenarioPack

_EMITTERS = [
    emit_tc23,
]

# Canary file keys — TC-23: 6 input files
_CANARY_FILE_KEYS: list[str] = sorted([
    # TC-23
    "tc23_bank_chf",
    "tc23_bank_eur",
    "tc23_bank_usd",
    "tc23_gl_cash",
    "tc23_bank_confirm",
    "tc23_snb_fx_rates",
])

PACK = ScenarioPack(
    pack_id="ch_variants",
    display_name="Swiss Variants (CH)",
    test_cases=[
        "TC-23",
    ],
    canary_file_keys=_CANARY_FILE_KEYS,
    emitters=_EMITTERS,
    dependencies=["cascade_accounting_core"],
)
