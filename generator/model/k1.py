"""Cascade Industries K-1 partnership investment data (TC-07 of prompt.md).

Generates 8 partnership investments with Schedule K-1 box data (boxes 1–13
plus Box 20 codes).  Amounts range from $5,000 to $2.3M.  One K-1 is marked
as corrected/amended: ordinary income changed from $340,000 to $285,000 and
a $55,000 guaranteed payment was added.  Section 199A amounts are present but
flagged N/A because Cascade Industries is a C-corporation.

Three K-1s are "system-generated" (clean layout); five come from different
partnerships with varying PDF layouts.

Determinism: uses only the passed ``rng``; no unordered sets or wall-clock reads.
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import Decimal
from enum import Enum


class K1LayoutType(Enum):
    """PDF layout style for the K-1."""
    SYSTEM_CLEAN = "system_clean"
    VARYING = "varying"


@dataclass(frozen=True)
class K1Box20Code:
    """A single Box 20 (other information) code entry."""
    code: str       # e.g. "A", "Z", "AH"
    description: str
    amount: Decimal


@dataclass(frozen=True)
class K1Amendment:
    """Tracks what changed between original and amended K-1."""
    field_changed: str
    original_value: Decimal
    amended_value: Decimal
    description: str


@dataclass(frozen=True)
class K1Investment:
    """A single Schedule K-1 from a partnership investment."""

    k1_id: str                    # K1-001 through K1-008
    partnership_name: str
    partnership_ein: str
    entity_code: str              # Which Cascade entity holds this investment
    tax_year: int
    layout_type: K1LayoutType
    is_amended: bool

    # Box data — None means not applicable / no amount
    box_1_ordinary_income: Decimal | None        # Ordinary business income (loss)
    box_2_net_rental_income: Decimal | None       # Net rental real estate income (loss)
    box_3_other_rental_income: Decimal | None     # Other net rental income (loss)
    box_4a_guaranteed_payments_services: Decimal | None
    box_4b_guaranteed_payments_capital: Decimal | None
    box_4c_total_guaranteed_payments: Decimal | None
    box_5_interest_income: Decimal | None
    box_6a_ordinary_dividends: Decimal | None
    box_6b_qualified_dividends: Decimal | None
    box_7_royalties: Decimal | None
    box_8_net_st_capital_gain: Decimal | None      # Net short-term capital gain (loss)
    box_9a_net_lt_capital_gain: Decimal | None     # Net long-term capital gain (loss)
    box_9b_collectibles_gain: Decimal | None
    box_9c_unrecaptured_1250: Decimal | None
    box_10_net_1231_gain: Decimal | None           # Net section 1231 gain (loss)
    box_11_other_income: Decimal | None            # Other income (loss)
    box_12_section_179: Decimal | None             # Section 179 deduction
    box_13_other_deductions: Decimal | None        # Other deductions

    # Box 20 codes
    box_20_codes: tuple[K1Box20Code, ...]

    # Section 199A — present on some K-1s but N/A for C-corp
    section_199a_qbi: Decimal | None
    section_199a_wages: Decimal | None
    section_199a_ubia: Decimal | None

    # Amendment details (only populated if is_amended)
    amendments: tuple[K1Amendment, ...] = ()

    @property
    def total_income(self) -> Decimal:
        """Sum of all income boxes (1–11)."""
        boxes = [
            self.box_1_ordinary_income,
            self.box_2_net_rental_income,
            self.box_3_other_rental_income,
            self.box_4c_total_guaranteed_payments,
            self.box_5_interest_income,
            self.box_6a_ordinary_dividends,
            self.box_7_royalties,
            self.box_8_net_st_capital_gain,
            self.box_9a_net_lt_capital_gain,
            self.box_10_net_1231_gain,
            self.box_11_other_income,
        ]
        return sum((b for b in boxes if b is not None), Decimal(0))

    @property
    def total_deductions(self) -> Decimal:
        """Sum of deduction boxes (12–13)."""
        boxes = [
            self.box_12_section_179,
            self.box_13_other_deductions,
        ]
        return sum((b for b in boxes if b is not None), Decimal(0))


# ── K-1 templates ──────────────────────────────────────────────────────────
# Each dict defines a partnership and its K-1 box data.
# Amounts are strings to preserve exact decimal values.

_K1_TEMPLATES: list[dict] = [
    # K1-001: Large manufacturing partnership — system-generated, clean
    dict(
        partnership_name="Pacific Manufacturing Partners LP",
        partnership_ein="91-3456789",
        entity_code="PC",
        layout_type=K1LayoutType.SYSTEM_CLEAN,
        is_amended=False,
        box_1="2300000",       # $2.3M — the high end
        box_2=None,
        box_3=None,
        box_4a=None, box_4b=None, box_4c=None,
        box_5="18500",
        box_6a="42000", box_6b="42000",
        box_7=None,
        box_8=None,
        box_9a="156000", box_9b=None, box_9c=None,
        box_10=None,
        box_11=None,
        box_12=None,
        box_13="28000",
        box_20=[
            ("A", "Investment income", "60500"),
            ("AH", "Section 199A QBI", "2300000"),
        ],
        sec_199a_qbi="2300000",
        sec_199a_wages="890000",
        sec_199a_ubia="4200000",
    ),
    # K1-002: Real estate partnership — varying layout
    dict(
        partnership_name="Willamette Real Estate Holdings LLC",
        partnership_ein="93-1122334",
        entity_code="CI",
        layout_type=K1LayoutType.VARYING,
        is_amended=False,
        box_1=None,
        box_2="185000",
        box_3=None,
        box_4a=None, box_4b=None, box_4c=None,
        box_5="7200",
        box_6a=None, box_6b=None,
        box_7=None,
        box_8=None,
        box_9a="92000", box_9b=None, box_9c="35000",
        box_10=None,
        box_11="12500",
        box_12=None,
        box_13="41000",
        box_20=[
            ("Z", "Net income from rental activity", "185000"),
        ],
        sec_199a_qbi="185000",
        sec_199a_wages=None,
        sec_199a_ubia="3100000",
    ),
    # K1-003: Tech venture fund — system-generated, clean
    dict(
        partnership_name="Cascade Ventures Fund III LP",
        partnership_ein="47-8899001",
        entity_code="AM",
        layout_type=K1LayoutType.SYSTEM_CLEAN,
        is_amended=False,
        box_1="78000",
        box_2=None,
        box_3=None,
        box_4a=None, box_4b=None, box_4c=None,
        box_5="5200",
        box_6a="15800", box_6b="11200",
        box_7="32000",
        box_8="24000",
        box_9a="445000", box_9b=None, box_9c=None,
        box_10=None,
        box_11=None,
        box_12=None,
        box_13="19500",
        box_20=[
            ("A", "Investment income", "53000"),
        ],
        sec_199a_qbi=None,
        sec_199a_wages=None,
        sec_199a_ubia=None,
    ),
    # K1-004: Small energy partnership — varying layout (THE AMENDED ONE)
    # Original: ordinary income $340,000, no guaranteed payments
    # Amended: ordinary income $285,000, guaranteed payment $55,000
    dict(
        partnership_name="Columbia Basin Energy Partners LLC",
        partnership_ein="91-5566778",
        entity_code="PC",
        layout_type=K1LayoutType.VARYING,
        is_amended=True,
        box_1="285000",       # Amended from $340,000
        box_2=None,
        box_3=None,
        box_4a="55000", box_4b=None, box_4c="55000",  # Added in amendment
        box_5="8900",
        box_6a=None, box_6b=None,
        box_7=None,
        box_8=None,
        box_9a=None, box_9b=None, box_9c=None,
        box_10="67000",
        box_11=None,
        box_12=None,
        box_13="15200",
        box_20=[
            ("A", "Investment income", "8900"),
            ("AH", "Section 199A QBI", "340000"),
        ],
        sec_199a_qbi="340000",
        sec_199a_wages="125000",
        sec_199a_ubia="780000",
        amendment_details=[
            ("box_1_ordinary_income", "340000", "285000",
             "Ordinary income reduced from $340,000 to $285,000"),
            ("box_4c_total_guaranteed_payments", "0", "55000",
             "Added $55,000 guaranteed payment"),
        ],
    ),
    # K1-005: Logistics JV — varying layout
    dict(
        partnership_name="Great Lakes Logistics JV",
        partnership_ein="36-2233445",
        entity_code="DS",
        layout_type=K1LayoutType.VARYING,
        is_amended=False,
        box_1="520000",
        box_2=None,
        box_3=None,
        box_4a=None, box_4b=None, box_4c=None,
        box_5="3100",
        box_6a="8500", box_6b="8500",
        box_7=None,
        box_8=None,
        box_9a=None, box_9b=None, box_9c=None,
        box_10=None,
        box_11="5000",
        box_12="12000",
        box_13="9800",
        box_20=[
            ("A", "Investment income", "11600"),
        ],
        sec_199a_qbi="520000",
        sec_199a_wages="310000",
        sec_199a_ubia="1450000",
    ),
    # K1-006: Materials research partnership — system-generated, clean
    dict(
        partnership_name="Advanced Polymer Research Partners LP",
        partnership_ein="74-6677889",
        entity_code="AM",
        layout_type=K1LayoutType.SYSTEM_CLEAN,
        is_amended=False,
        box_1="145000",
        box_2=None,
        box_3=None,
        box_4a="25000", box_4b=None, box_4c="25000",
        box_5="2800",
        box_6a=None, box_6b=None,
        box_7="18000",
        box_8=None,
        box_9a="38000", box_9b=None, box_9c=None,
        box_10=None,
        box_11=None,
        box_12="8500",
        box_13="6200",
        box_20=[
            ("A", "Investment income", "2800"),
            ("AH", "Section 199A QBI", "170000"),
        ],
        sec_199a_qbi="170000",
        sec_199a_wages="95000",
        sec_199a_ubia="420000",
    ),
    # K1-007: Small minority interest — varying layout, small amounts
    dict(
        partnership_name="Portland Innovation Hub LLC",
        partnership_ein="93-4455667",
        entity_code="CI",
        layout_type=K1LayoutType.VARYING,
        is_amended=False,
        box_1="5000",          # $5K — the low end
        box_2=None,
        box_3=None,
        box_4a=None, box_4b=None, box_4c=None,
        box_5="800",
        box_6a="1200", box_6b="1200",
        box_7=None,
        box_8="2500",
        box_9a=None, box_9b=None, box_9c=None,
        box_10=None,
        box_11=None,
        box_12=None,
        box_13="1800",
        box_20=[],
        sec_199a_qbi="5000",
        sec_199a_wages="3200",
        sec_199a_ubia="15000",
    ),
    # K1-008: Distribution partnership — varying layout
    dict(
        partnership_name="Midwest Supply Chain Partners LP",
        partnership_ein="36-7788990",
        entity_code="DS",
        layout_type=K1LayoutType.VARYING,
        is_amended=False,
        box_1="890000",
        box_2=None,
        box_3=None,
        box_4a=None, box_4b="35000", box_4c="35000",
        box_5="12400",
        box_6a="22000", box_6b="18000",
        box_7=None,
        box_8=None,
        box_9a="175000", box_9b=None, box_9c=None,
        box_10="28000",
        box_11=None,
        box_12=None,
        box_13="32500",
        box_20=[
            ("A", "Investment income", "34400"),
            ("AH", "Section 199A QBI", "925000"),
        ],
        sec_199a_qbi="925000",
        sec_199a_wages="480000",
        sec_199a_ubia="2100000",
    ),
]


# ── Generation ─────────────────────────────────────────────────────────────

def _dec(val: str | None) -> Decimal | None:
    """Convert a string amount to Decimal, or None."""
    if val is None:
        return None
    return Decimal(val)


def generate_k1_investments() -> list[K1Investment]:
    """Generate the 8 K-1 partnership investments deterministically.

    No RNG needed — all values are fixed per prompt.md TC-07 spec.
    Returns a list sorted by k1_id.
    """
    investments: list[K1Investment] = []

    for idx, tmpl in enumerate(_K1_TEMPLATES, start=1):
        k1_id = f"K1-{idx:03d}"

        # Build Box 20 codes
        box_20_codes = tuple(
            K1Box20Code(code=c, description=d, amount=Decimal(a))
            for c, d, a in tmpl.get("box_20", [])
        )

        # Build amendments if present
        amendments = tuple(
            K1Amendment(
                field_changed=fc,
                original_value=Decimal(ov),
                amended_value=Decimal(av),
                description=desc,
            )
            for fc, ov, av, desc in tmpl.get("amendment_details", [])
        )

        inv = K1Investment(
            k1_id=k1_id,
            partnership_name=tmpl["partnership_name"],
            partnership_ein=tmpl["partnership_ein"],
            entity_code=tmpl["entity_code"],
            tax_year=2025,
            layout_type=tmpl["layout_type"],
            is_amended=tmpl["is_amended"],
            box_1_ordinary_income=_dec(tmpl["box_1"]),
            box_2_net_rental_income=_dec(tmpl["box_2"]),
            box_3_other_rental_income=_dec(tmpl["box_3"]),
            box_4a_guaranteed_payments_services=_dec(tmpl["box_4a"]),
            box_4b_guaranteed_payments_capital=_dec(tmpl["box_4b"]),
            box_4c_total_guaranteed_payments=_dec(tmpl["box_4c"]),
            box_5_interest_income=_dec(tmpl["box_5"]),
            box_6a_ordinary_dividends=_dec(tmpl["box_6a"]),
            box_6b_qualified_dividends=_dec(tmpl["box_6b"]),
            box_7_royalties=_dec(tmpl["box_7"]),
            box_8_net_st_capital_gain=_dec(tmpl["box_8"]),
            box_9a_net_lt_capital_gain=_dec(tmpl["box_9a"]),
            box_9b_collectibles_gain=_dec(tmpl["box_9b"]),
            box_9c_unrecaptured_1250=_dec(tmpl["box_9c"]),
            box_10_net_1231_gain=_dec(tmpl["box_10"]),
            box_11_other_income=_dec(tmpl["box_11"]),
            box_12_section_179=_dec(tmpl["box_12"]),
            box_13_other_deductions=_dec(tmpl["box_13"]),
            box_20_codes=box_20_codes,
            section_199a_qbi=_dec(tmpl["sec_199a_qbi"]),
            section_199a_wages=_dec(tmpl.get("sec_199a_wages")),
            section_199a_ubia=_dec(tmpl.get("sec_199a_ubia")),
            amendments=amendments,
        )
        investments.append(inv)

    investments.sort(key=lambda k: k.k1_id)
    return investments


# ── Consolidation helpers ──────────────────────────────────────────────────

def consolidated_totals(
    investments: list[K1Investment],
) -> dict[str, Decimal]:
    """Consolidate all K-1 data into a summary dict keyed by box label.

    These totals should match TC-07 gold standard exactly.
    """
    totals: dict[str, Decimal] = {}

    box_fields = [
        ("box_1_ordinary_income", "Box 1 - Ordinary business income"),
        ("box_2_net_rental_income", "Box 2 - Net rental real estate income"),
        ("box_3_other_rental_income", "Box 3 - Other net rental income"),
        ("box_4c_total_guaranteed_payments", "Box 4c - Guaranteed payments"),
        ("box_5_interest_income", "Box 5 - Interest income"),
        ("box_6a_ordinary_dividends", "Box 6a - Ordinary dividends"),
        ("box_6b_qualified_dividends", "Box 6b - Qualified dividends"),
        ("box_7_royalties", "Box 7 - Royalties"),
        ("box_8_net_st_capital_gain", "Box 8 - Net ST capital gain"),
        ("box_9a_net_lt_capital_gain", "Box 9a - Net LT capital gain"),
        ("box_9b_collectibles_gain", "Box 9b - Collectibles gain"),
        ("box_9c_unrecaptured_1250", "Box 9c - Unrecaptured Sec 1250 gain"),
        ("box_10_net_1231_gain", "Box 10 - Net section 1231 gain"),
        ("box_11_other_income", "Box 11 - Other income"),
        ("box_12_section_179", "Box 12 - Section 179 deduction"),
        ("box_13_other_deductions", "Box 13 - Other deductions"),
    ]

    for attr, label in box_fields:
        total = Decimal(0)
        for inv in investments:
            val = getattr(inv, attr)
            if val is not None:
                total += val
        if total != 0:
            totals[label] = total

    # Section 199A totals (present but N/A for C-corp)
    for attr, label in [
        ("section_199a_qbi", "Section 199A - QBI (N/A to C-corp)"),
        ("section_199a_wages", "Section 199A - W-2 Wages (N/A to C-corp)"),
        ("section_199a_ubia", "Section 199A - UBIA (N/A to C-corp)"),
    ]:
        total = Decimal(0)
        for inv in investments:
            val = getattr(inv, attr)
            if val is not None:
                total += val
        if total != 0:
            totals[label] = total

    return totals
