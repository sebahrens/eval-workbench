"""European VAT and cross-border tax model for TC-10-EU.

Generates deterministic intercompany sales data, VAT registrations,
quarterly VAT return summaries, and EU VAT rules reference content
for the Cascade Europe Holdings B.V. group.

All data is deterministic (no RNG) per the design bead synth-data-eu.14.
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal

# ── Entity constants (imported from design) ──────────────────────────────────

ENTITY_NAMES_VAT = {
    "CE": "Cascade Europe Holdings B.V.",
    "CP": "Cascade Präzisionsteile GmbH",
    "CM": "Cascade Matériaux Avancés SAS",
    "CD": "Cascade Distribution Services Ltd",
}

ENTITY_JURISDICTIONS_VAT = {
    "CE": "Netherlands",
    "CP": "Germany",
    "CM": "France",
    "CD": "United Kingdom",
}

VAT_IDS = {
    "CE": "NL123456789B01",
    "CP": "DE123456789",
    "CM": "FR12345678901",
    "CD": "GB123456789",
}

VAT_RATES = {
    "NL": Decimal("0.21"),  # 21%
    "DE": Decimal("0.19"),  # 19%
    "FR": Decimal("0.20"),  # 20%
    "UK": Decimal("0.20"),  # 20%
}

# ── Revenue / IC flow constants ──────────────────────────────────────────────

CP_REVENUE = Decimal("45000000")
CM_REVENUE = Decimal("32000000")
CD_REVENUE = Decimal("21000000")

MGMT_FEE_PCT = Decimal("0.015")  # 1.5%
ROYALTY_PCT = Decimal("0.03")    # 3% of CP revenue

# IC goods flows (annual totals from design)
CP_TO_CM_RAW_MATERIALS = Decimal("4200000")
CP_TO_CD_FINISHED_GOODS = Decimal("6800000")
CE_TO_CP_MGMT_FEE = (CP_REVENUE * MGMT_FEE_PCT).quantize(Decimal("1"), ROUND_HALF_UP)
CE_TO_CM_MGMT_FEE = (CM_REVENUE * MGMT_FEE_PCT).quantize(Decimal("1"), ROUND_HALF_UP)
CE_TO_CD_MGMT_FEE = (CD_REVENUE * MGMT_FEE_PCT).quantize(Decimal("1"), ROUND_HALF_UP)
CM_TO_CP_ROYALTY = (CP_REVENUE * ROYALTY_PCT).quantize(Decimal("1"), ROUND_HALF_UP)

# Q2 shipment amounts for proof-of-dispatch gap
# 5 shipments in Q2, ~€1.36M each → total ~€6.8M/5 quarters ≈ €1.36M/Q
CP_TO_CD_Q2_TOTAL = (CP_TO_CD_FINISHED_GOODS / 4).quantize(Decimal("1"), ROUND_HALF_UP)
CP_TO_CD_Q2_PER_SHIPMENT = (CP_TO_CD_Q2_TOTAL / 5).quantize(Decimal("1"), ROUND_HALF_UP)
CP_TO_CD_Q2_UNDOCUMENTED = CP_TO_CD_Q2_PER_SHIPMENT * 2  # 2 of 5 missing

# ERR-EU-010: CE→CM Q3 management fee with incorrect 20% French VAT
ERR_EU_010_QUARTER = "Q3"
ERR_EU_010_AMOUNT = (CE_TO_CM_MGMT_FEE / 4).quantize(Decimal("1"), ROUND_HALF_UP)
ERR_EU_010_VAT_CHARGED = (ERR_EU_010_AMOUNT * VAT_RATES["FR"]).quantize(
    Decimal("1"), ROUND_HALF_UP,
)


# ── Intercompany sales data class ───────────────────────────────────────────

@dataclass(frozen=True)
class ICSaleEU:
    """A single intercompany/third-party sale record."""

    seller: str
    buyer: str
    description: str
    amount_eur: Decimal
    vat_treatment: str  # "Zero-rated (intra-EU)", "Reverse charge", etc.
    invoice_vat_rate: str  # "0%", "20%", "N/A", etc.
    incoterms: str  # "EXW", "DDP", "N/A" for services
    proof_of_dispatch: str  # "Y", "N", "Partial", "N/A"
    quarter: str  # "Q1", "Q2", etc.
    is_error: bool  # True for ERR-EU-010 row


def generate_ic_sales_eu() -> list[ICSaleEU]:
    """Generate intercompany and third-party sales for FY2025."""
    rows: list[ICSaleEU] = []

    # Quarterly split helper
    def q_split(annual: Decimal) -> list[tuple[str, Decimal]]:
        q = (annual / 4).quantize(Decimal("1"), ROUND_HALF_UP)
        return [
            ("Q1", q), ("Q2", q), ("Q3", q),
            ("Q4", annual - 3 * q),
        ]

    # 1. CP→CM raw materials (intra-EU supply of goods)
    for qtr, amt in q_split(CP_TO_CM_RAW_MATERIALS):
        rows.append(ICSaleEU(
            seller="CP", buyer="CM",
            description=f"Raw materials — precision components {qtr} FY2025",
            amount_eur=amt,
            vat_treatment="Zero-rated (intra-EU supply, Art. 138)",
            invoice_vat_rate="0%",
            incoterms="EXW Munich",
            proof_of_dispatch="Y",
            quarter=qtr,
            is_error=False,
        ))

    # 2. CP→CD finished goods (export to UK, third country post-Brexit)
    for qtr, amt in q_split(CP_TO_CD_FINISHED_GOODS):
        pod = "Y"
        if qtr == "Q2":
            pod = "Partial"  # Missing data trap: 2/5 shipments lack docs
        rows.append(ICSaleEU(
            seller="CP", buyer="CD",
            description=f"Finished goods — industrial assemblies {qtr} FY2025",
            amount_eur=amt,
            vat_treatment="Zero-rated (export to third country, post-Brexit)",
            invoice_vat_rate="0%",
            incoterms="DDP London",
            proof_of_dispatch=pod,
            quarter=qtr,
            is_error=False,
        ))

    # 3. CE→CP management fees (reverse charge B2B services)
    for qtr, amt in q_split(CE_TO_CP_MGMT_FEE):
        rows.append(ICSaleEU(
            seller="CE", buyer="CP",
            description=f"Management services — strategic oversight, treasury, legal {qtr} FY2025",
            amount_eur=amt,
            vat_treatment="Reverse charge (Art. 196, B2B services)",
            invoice_vat_rate="0%",
            incoterms="N/A",
            proof_of_dispatch="N/A",
            quarter=qtr,
            is_error=False,
        ))

    # 4. CE→CM management fees — Q3 has ERR-EU-010
    for qtr, amt in q_split(CE_TO_CM_MGMT_FEE):
        if qtr == ERR_EU_010_QUARTER:
            # ERR-EU-010: CE charges 20% French VAT — wrong!
            rows.append(ICSaleEU(
                seller="CE", buyer="CM",
                description=f"Management services — strategic oversight, treasury, legal {qtr} FY2025",
                amount_eur=amt,
                vat_treatment="French VAT charged at 20%",
                invoice_vat_rate="20%",
                incoterms="N/A",
                proof_of_dispatch="N/A",
                quarter=qtr,
                is_error=True,
            ))
        else:
            rows.append(ICSaleEU(
                seller="CE", buyer="CM",
                description=f"Management services — strategic oversight, treasury, legal {qtr} FY2025",
                amount_eur=amt,
                vat_treatment="Reverse charge (Art. 196, B2B services)",
                invoice_vat_rate="0%",
                incoterms="N/A",
                proof_of_dispatch="N/A",
                quarter=qtr,
                is_error=False,
            ))

    # 5. CE→CD management fees (outside scope of EU VAT — UK)
    for qtr, amt in q_split(CE_TO_CD_MGMT_FEE):
        rows.append(ICSaleEU(
            seller="CE", buyer="CD",
            description=f"Management services — strategic oversight, treasury, legal {qtr} FY2025",
            amount_eur=amt,
            vat_treatment="Outside scope of EU VAT (UK reverse charge)",
            invoice_vat_rate="0%",
            incoterms="N/A",
            proof_of_dispatch="N/A",
            quarter=qtr,
            is_error=False,
        ))

    # 6. CM→CP R&D royalty (reverse charge B2B services)
    for qtr, amt in q_split(CM_TO_CP_ROYALTY):
        rows.append(ICSaleEU(
            seller="CM", buyer="CP",
            description=f"R&D royalty — technology license {qtr} FY2025",
            amount_eur=amt,
            vat_treatment="Reverse charge (Art. 196, B2B services)",
            invoice_vat_rate="0%",
            incoterms="N/A",
            proof_of_dispatch="N/A",
            quarter=qtr,
            is_error=False,
        ))

    # Sort deterministically
    return sorted(rows, key=lambda r: (r.quarter, r.seller, r.buyer, r.description))


# ── VAT registration data class ─────────────────────────────────────────────

@dataclass(frozen=True)
class VATRegistration:
    """One entity's VAT registration in a jurisdiction."""

    entity_code: str
    country: str
    vat_id: str
    registration_date: str  # DD.MM.YYYY
    vat_group: str  # "Y" / "N"
    fiscal_representative: str
    ecsl_filed: str  # "Y" / "N" / "N/A"
    intrastat_exceeded: str  # "Y" / "N" / "N/A"
    status: str  # "Active", "Pending", etc.


def generate_vat_registrations() -> list[VATRegistration]:
    """Generate VAT registration data for all entities + traps."""
    regs = [
        VATRegistration(
            entity_code="CE", country="Netherlands",
            vat_id="NL123456789B01", registration_date="01.01.2020",
            vat_group="N", fiscal_representative="",
            ecsl_filed="Y", intrastat_exceeded="N/A",
            status="Active",
        ),
        VATRegistration(
            entity_code="CP", country="Germany",
            vat_id="DE123456789", registration_date="01.01.2020",
            vat_group="N", fiscal_representative="",
            ecsl_filed="Y", intrastat_exceeded="Y",
            status="Active",
        ),
        # Missing data trap: CP has pending Polish registration
        VATRegistration(
            entity_code="CP", country="Poland",
            vat_id="PL9876543210", registration_date="15.03.2025",
            vat_group="N", fiscal_representative="",
            ecsl_filed="N", intrastat_exceeded="N",
            status="Pending — applied 2025-03-15",
        ),
        VATRegistration(
            entity_code="CM", country="France",
            vat_id="FR12345678901", registration_date="01.01.2020",
            vat_group="N", fiscal_representative="",
            ecsl_filed="Y", intrastat_exceeded="Y",
            status="Active",
        ),
        VATRegistration(
            entity_code="CD", country="United Kingdom",
            vat_id="GB123456789", registration_date="01.01.2020",
            vat_group="N", fiscal_representative="",
            ecsl_filed="N/A", intrastat_exceeded="N/A",
            status="Active",
        ),
    ]
    return regs


# ── Quarterly VAT return data class ─────────────────────────────────────────

@dataclass(frozen=True)
class VATReturnRow:
    """One quarterly VAT return summary row."""

    entity_code: str
    quarter: str
    output_vat_domestic: Decimal
    output_vat_intra_eu_export: Decimal
    input_vat_domestic: Decimal
    input_vat_reverse_charge: Decimal
    vat_payable: Decimal
    filing_status: str


def generate_vat_returns() -> list[VATReturnRow]:
    """Generate quarterly VAT return summaries for FY2025."""
    rows: list[VATReturnRow] = []

    # CP Germany: 19% on domestic, 0% on intra-EU, input VAT recovery
    cp_dom_rev_q = (CP_REVENUE * Decimal("0.65") / 4).quantize(Decimal("1"), ROUND_HALF_UP)
    cp_dom_output_q = (cp_dom_rev_q * VAT_RATES["DE"]).quantize(Decimal("1"), ROUND_HALF_UP)
    cp_input_q = (cp_dom_output_q * Decimal("0.85")).quantize(Decimal("1"), ROUND_HALF_UP)
    for qtr in ["Q1", "Q2", "Q3", "Q4"]:
        extra_input = Decimal("0")
        if qtr == "Q3":
            extra_input = Decimal("450000")  # Large capital equipment purchase
        input_total = cp_input_q + extra_input
        payable = cp_dom_output_q - input_total
        rows.append(VATReturnRow(
            entity_code="CP", quarter=qtr,
            output_vat_domestic=cp_dom_output_q,
            output_vat_intra_eu_export=Decimal("0"),
            input_vat_domestic=input_total,
            input_vat_reverse_charge=Decimal("0"),
            vat_payable=payable,
            filing_status="Filed — Assessed",
        ))

    # CM France: 20% on domestic, reverse charge on services received
    cm_dom_rev_q = (CM_REVENUE * Decimal("0.70") / 4).quantize(Decimal("1"), ROUND_HALF_UP)
    cm_dom_output_q = (cm_dom_rev_q * VAT_RATES["FR"]).quantize(Decimal("1"), ROUND_HALF_UP)
    cm_input_q = (cm_dom_output_q * Decimal("0.80")).quantize(Decimal("1"), ROUND_HALF_UP)
    cm_rc_q = (CE_TO_CM_MGMT_FEE / 4 * VAT_RATES["FR"]).quantize(Decimal("1"), ROUND_HALF_UP)
    for qtr in ["Q1", "Q2", "Q3", "Q4"]:
        rows.append(VATReturnRow(
            entity_code="CM", quarter=qtr,
            output_vat_domestic=cm_dom_output_q,
            output_vat_intra_eu_export=Decimal("0"),
            input_vat_domestic=cm_input_q,
            input_vat_reverse_charge=cm_rc_q,
            vat_payable=cm_dom_output_q - cm_input_q,
            filing_status="Filed — Assessed",
        ))

    # CD UK: 20% on domestic, import VAT via postponed accounting
    cd_dom_rev_q = (CD_REVENUE * Decimal("0.80") / 4).quantize(Decimal("1"), ROUND_HALF_UP)
    cd_dom_output_q = (cd_dom_rev_q * VAT_RATES["UK"]).quantize(Decimal("1"), ROUND_HALF_UP)
    cd_input_q = (cd_dom_output_q * Decimal("0.75")).quantize(Decimal("1"), ROUND_HALF_UP)
    cd_import_vat_q = (CP_TO_CD_FINISHED_GOODS / 4 * VAT_RATES["UK"]).quantize(
        Decimal("1"), ROUND_HALF_UP,
    )
    for qtr in ["Q1", "Q2", "Q3", "Q4"]:
        rows.append(VATReturnRow(
            entity_code="CD", quarter=qtr,
            output_vat_domestic=cd_dom_output_q,
            output_vat_intra_eu_export=Decimal("0"),
            input_vat_domestic=cd_input_q,
            input_vat_reverse_charge=cd_import_vat_q,
            vat_payable=cd_dom_output_q - cd_input_q,
            filing_status="Filed — Assessed",
        ))

    # CE Netherlands: minimal — mostly reverse charge
    ce_dom_q = Decimal("15000")  # Small NL domestic supply
    ce_output_q = (ce_dom_q * VAT_RATES["NL"]).quantize(Decimal("1"), ROUND_HALF_UP)
    ce_input_q = Decimal("2000")  # Minimal input
    for qtr in ["Q1", "Q2", "Q3", "Q4"]:
        status = "Filed — Assessed"
        if qtr == "Q4":
            status = "Filed — Pending Assessment"  # Unusual — trap
        rows.append(VATReturnRow(
            entity_code="CE", quarter=qtr,
            output_vat_domestic=ce_output_q,
            output_vat_intra_eu_export=Decimal("0"),
            input_vat_domestic=ce_input_q,
            input_vat_reverse_charge=Decimal("0"),
            vat_payable=ce_output_q - ce_input_q,
            filing_status=status,
        ))

    return sorted(rows, key=lambda r: (r.entity_code, r.quarter))


# ── Third-party sales summary (missing data: CM Italian sales) ──

CM_ITALIAN_SALES = Decimal("380000")  # Missing data trap


# ── Canary keys ──────────────────────────────────────────────────────────────

ALL_CANARY_KEYS_TC10EU: list[str] = sorted([
    "tc10eu_intercompany_sales_fy2025",
    "tc10eu_vat_registrations",
    "tc10eu_vat_returns_summary_fy2025",
    "tc10eu_eu_vat_rules_reference",
])
