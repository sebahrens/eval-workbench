"""Formatter: TC-20 — HR Diligence Exposure Summary.

Emits:
- test_cases/TC-20/input_files/agreements/
    7 docx employment agreements (one per EA-NNN)
- test_cases/TC-20/input_files/employee_census.xlsx
    Executive employee census
- test_cases/TC-20/input_files/severance_schedule.xlsx
    Severance exposure schedule
- test_cases/TC-20/input_files/retention_plan.xlsx
    Retention award schedule
- test_cases/TC-20/input_files/contractor_roster.xlsx
    Contractor classification roster

Judgment traps:
  - Executive severance exposure tied to salary and agreement multiplier
  - Retention/severance double-count avoidance (Dr. Patel: greater-of, not both)
  - Missing executed agreement (EA-006 Dr. Patel)
  - Contractor classification signal as follow-up, not definitive legal conclusion
  - Census and agreement source citations

Uses the canonical model — never hardcodes numbers.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from generator.canaries import CanaryRegistry
from generator.errors import ErrorRegistry
from generator.golds.framework import GoldStandard, register_gold
from generator.manifest import Manifest
from generator.model.build import CascadeModel
from generator.model.hr_diligence import (
    CONTRACTOR_CLASSIFICATION_SIGNALS,
    EMPLOYMENT_AGREEMENTS,
    RETENTION_AWARDS,
    SEVERANCE_EXPOSURES,
    high_risk_contractors,
    missing_executed_agreements,
    total_retention_awards,
    total_severance_exposure,
)
from generator.writers.hr import (
    write_all_employment_agreements,
    write_contractor_roster,
    write_employee_census,
    write_retention_schedule,
    write_severance_schedule,
)

# ── Constants ────────────────────────────────────────────────────────────────

_TC = "TC-20"
_INPUT_DIR = f"test_cases/{_TC}/input_files"

# Canary file keys — 7 agreements + 1 census + 1 severance + 1 retention + 1 contractor = 11
_CANARY_KEYS: list[str] = sorted([
    # Employment agreements (one per EA-NNN)
    "tc20_agreement_ea_001",
    "tc20_agreement_ea_002",
    "tc20_agreement_ea_003",
    "tc20_agreement_ea_004",
    "tc20_agreement_ea_005",
    "tc20_agreement_ea_006",
    "tc20_agreement_ea_007",
    # Employee census
    "tc20_employee_census",
    # Severance schedule
    "tc20_severance_schedule",
    # Retention plan
    "tc20_retention_plan",
    # Contractor roster
    "tc20_contractor_roster",
])


# ── File writers ─────────────────────────────────────────────────────────────


def _write_agreements(
    output_dir: Path,
    canaries: CanaryRegistry,
    manifest: Manifest,
) -> None:
    """Write one docx per EmploymentAgreement into TC-20/input_files/agreements/."""
    agreements_dir = output_dir / _INPUT_DIR / "agreements"
    agreements_dir.mkdir(parents=True, exist_ok=True)

    canary_keys = {
        ea.agreement_id: f"tc20_agreement_{ea.agreement_id.lower().replace('-', '_')}"
        for ea in EMPLOYMENT_AGREEMENTS
    }

    locations = write_all_employment_agreements(agreements_dir, canaries, canary_keys)

    for ea in EMPLOYMENT_AGREEMENTS:
        ckey = canary_keys[ea.agreement_id]
        fname = f"agreement_{ea.agreement_id.lower()}.docx"
        rel_path = f"{_INPUT_DIR}/agreements/{fname}"
        canaries.set_location(ckey, rel_path, locations[ea.agreement_id])
        manifest.register(rel_path, "docx", test_cases=[_TC])


def _write_census(
    output_dir: Path,
    canaries: CanaryRegistry,
    manifest: Manifest,
) -> None:
    """Write the executive employee census spreadsheet."""
    xlsx_path = output_dir / _INPUT_DIR / "employee_census.xlsx"
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)

    ckey = "tc20_employee_census"
    location = write_employee_census(xlsx_path, canaries, ckey)
    canaries.set_location(ckey, f"{_INPUT_DIR}/employee_census.xlsx", location)
    manifest.register(f"{_INPUT_DIR}/employee_census.xlsx", "xlsx", test_cases=[_TC])


def _write_severance(
    output_dir: Path,
    canaries: CanaryRegistry,
    manifest: Manifest,
) -> None:
    """Write the severance exposure schedule spreadsheet."""
    xlsx_path = output_dir / _INPUT_DIR / "severance_schedule.xlsx"
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)

    ckey = "tc20_severance_schedule"
    location = write_severance_schedule(xlsx_path, canaries, ckey)
    canaries.set_location(ckey, f"{_INPUT_DIR}/severance_schedule.xlsx", location)
    manifest.register(f"{_INPUT_DIR}/severance_schedule.xlsx", "xlsx", test_cases=[_TC])


def _write_retention(
    output_dir: Path,
    canaries: CanaryRegistry,
    manifest: Manifest,
) -> None:
    """Write the retention award schedule spreadsheet."""
    xlsx_path = output_dir / _INPUT_DIR / "retention_plan.xlsx"
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)

    ckey = "tc20_retention_plan"
    location = write_retention_schedule(xlsx_path, canaries, ckey)
    canaries.set_location(ckey, f"{_INPUT_DIR}/retention_plan.xlsx", location)
    manifest.register(f"{_INPUT_DIR}/retention_plan.xlsx", "xlsx", test_cases=[_TC])


def _write_contractors(
    output_dir: Path,
    canaries: CanaryRegistry,
    manifest: Manifest,
) -> None:
    """Write the contractor classification roster spreadsheet."""
    xlsx_path = output_dir / _INPUT_DIR / "contractor_roster.xlsx"
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)

    ckey = "tc20_contractor_roster"
    location = write_contractor_roster(xlsx_path, canaries, ckey)
    canaries.set_location(ckey, f"{_INPUT_DIR}/contractor_roster.xlsx", location)
    manifest.register(f"{_INPUT_DIR}/contractor_roster.xlsx", "xlsx", test_cases=[_TC])


# ── Prompt & Expected Behavior ───────────────────────────────────────────────


def _write_prompt(output_dir: Path) -> None:
    """Write test_cases/TC-20/prompt.md per spec."""
    text = """\
Review the HR diligence materials for Cascade Industries in connection
with a potential acquisition.

Using the employee census, executive employment agreements, severance
schedule, retention plan, and contractor roster provided:

1. Prepare an **exposure schedule** summarizing:
   - Total severance exposure by executive (base salary, multiplier,
     estimated payout, and trigger event)
   - Total retention award obligations (amount, vesting date, and
     forfeiture conditions)
   - Net combined people-cost exposure, avoiding double-counting
     where severance and retention provisions overlap

2. For each executive, verify that the severance payout in the schedule
   matches the agreement terms (base salary x multiplier). Flag any
   discrepancies.

3. Identify any **missing executed agreements** — cases where an
   employment agreement is referenced but no signed copy is on file.
   Note the implications for IP assignment and covenant enforceability.

4. Review the contractor roster for **classification risk signals**.
   For each contractor, assess whether the engagement characteristics
   (tenure, exclusivity, equipment, hours) suggest potential
   misclassification risk. Present findings as signals requiring
   further investigation, not as definitive legal conclusions.

5. Produce a **findings memo** covering:
   - Top severance exposures and change-of-control provisions
   - Retention/severance interaction (double-count risk)
   - Missing agreement and its impact
   - Contractor classification signals and recommended follow-up
   - Key assumptions and limitations

6. Cite specific source documents (agreement IDs, census data,
   schedule references) for every finding.
"""
    path = output_dir / f"test_cases/{_TC}/prompt.md"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text)


def _write_expected_behavior(output_dir: Path) -> None:
    """Write test_cases/TC-20/expected_behavior.md per spec."""
    # Pull numbers from canonical model for accuracy
    ceo_sev = SEVERANCE_EXPOSURES[0]  # SEV-001
    cfo_sev = SEVERANCE_EXPOSURES[1]  # SEV-002
    cto_sev = SEVERANCE_EXPOSURES[2]  # SEV-003
    patel_sev = SEVERANCE_EXPOSURES[5]  # SEV-006
    patel_ret = RETENTION_AWARDS[3]  # RET-004
    total_sev = total_severance_exposure()
    total_ret = total_retention_awards()

    text = f"""\
# TC-20: HR Diligence Exposure Summary — Expected Behavior

## Exposure Schedule Requirements
- The schedule must cover all 7 executives with employment agreements.
- Each entry must show base salary, multiplier, estimated payout, and trigger.
- Severance computations must match agreement terms exactly.
- Total severance exposure: ${int(total_sev):,}.
- Total retention awards: ${int(total_ret):,}.

## CEO Golden Parachute (SEV-001 / EA-001)
- Robert J. Cascade has a 3x change-of-control provision.
- Base salary ${ceo_sev.base_salary:,} x {ceo_sev.severance_multiplier}x = ${int(ceo_sev.estimated_payout):,}.
- This is the single largest severance exposure.
- The agent must cite EA-001 as the source.

## CFO and CTO Severance (SEV-002, SEV-003)
- Margaret L. Chen: ${cfo_sev.base_salary:,} x {cfo_sev.severance_multiplier}x = ${int(cfo_sev.estimated_payout):,}.
- David R. Nakamura: ${cto_sev.base_salary:,} x {cto_sev.severance_multiplier}x = ${int(cto_sev.estimated_payout):,}.
- Both triggered by change of control per EA-002 and EA-003.

## Retention/Severance Double-Count (SEV-006 / RET-004)
- Dr. Anika Patel has both severance (${int(patel_sev.estimated_payout):,}) and
  retention (${int(patel_ret.award_amount):,}).
- Per EA-006 terms: greater-of, not additive.
- Correct net exposure is ${int(patel_sev.estimated_payout):,} (the larger amount),
  NOT ${int(patel_sev.estimated_payout + patel_ret.award_amount):,}.
- The agent must identify this overlap and avoid double-counting.

## Missing Executed Agreement (EA-006)
- Dr. Anika Patel's agreement (EA-006) has no executed copy on file.
- Only a draft is available — management states it was signed but
  cannot locate the original.
- Implications: IP assignment clause may not be enforceable without
  an executed agreement; non-compete enforceability is uncertain.
- The agent must flag this as a diligence gap requiring follow-up.

## Contractor Classification Signals
- Martinez Technical Consulting (CCS-001): HIGH risk — 28 months,
  exclusive, company equipment, set hours. All major classification
  risk factors present.
- RDL Engineering Services (CCS-003): HIGH risk — 19 months,
  exclusive, company equipment, set hours. Reports to internal manager.
- Pinnacle IT Solutions (CCS-002): MEDIUM risk — 36 months but
  non-exclusive, own equipment, no set hours.
- GreenField Environmental (CCS-004): LOW risk — 6 months,
  project-based, independent.
- The agent must present these as signals for follow-up investigation,
  NOT as definitive legal conclusions about misclassification.

## Source Citation Requirements
- Every severance finding must reference the specific agreement ID.
- Retention findings must reference the retention award ID.
- Contractor findings must reference the signal ID and roster data.
- The census should be referenced for headcount and compensation verification.

## Scoring Focus
Grade on:
- Accuracy of severance computations (salary x multiplier = payout)
- Detection of retention/severance double-count risk (Patel)
- Identification of missing executed agreement (EA-006) and implications
- Contractor classification presented as signals, not conclusions
- Proper source citations (agreement IDs, not just "per management")
- Professional judgment and appropriate caveats
"""
    path = output_dir / f"test_cases/{_TC}/expected_behavior.md"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text)


# ── Gold Standard ────────────────────────────────────────────────────────────


@register_gold("TC-20")
def _tc20_gold(
    canaries: CanaryRegistry,
    errors: ErrorRegistry,
    **model_kwargs: Any,
) -> GoldStandard:
    """Build the TC-20 gold standard."""
    # Build canary verification from all TC-20 file keys
    canary_verification: dict[str, str] = {}
    for ea in EMPLOYMENT_AGREEMENTS:
        ckey = f"tc20_agreement_{ea.agreement_id.lower().replace('-', '_')}"
        label = f"read_agreement_{ea.agreement_id.lower().replace('-', '_')}"
        canary_verification[label] = canaries.canary_for(ckey)
    canary_verification["read_employee_census"] = canaries.canary_for(
        "tc20_employee_census"
    )
    canary_verification["read_severance_schedule"] = canaries.canary_for(
        "tc20_severance_schedule"
    )
    canary_verification["read_retention_plan"] = canaries.canary_for(
        "tc20_retention_plan"
    )
    canary_verification["read_contractor_roster"] = canaries.canary_for(
        "tc20_contractor_roster"
    )

    # Build severance exposure entries from canonical data
    severance_entries = []
    for sev in SEVERANCE_EXPOSURES:
        severance_entries.append({
            "exposure_id": sev.exposure_id,
            "employee_name": sev.employee_name,
            "employee_title": sev.employee_title,
            "entity_code": sev.entity_code,
            "base_salary": sev.base_salary,
            "severance_multiplier": str(sev.severance_multiplier),
            "estimated_payout": str(sev.estimated_payout),
            "trigger": sev.trigger,
        })

    # Build retention entries
    retention_entries = []
    for ret in RETENTION_AWARDS:
        retention_entries.append({
            "award_id": ret.award_id,
            "employee_name": ret.employee_name,
            "award_amount": str(ret.award_amount),
            "vesting_date": ret.vesting_date.isoformat(),
            "retention_period_months": ret.retention_period_months,
        })

    # Build contractor classification entries
    contractor_entries = []
    for ccs in CONTRACTOR_CLASSIFICATION_SIGNALS:
        contractor_entries.append({
            "signal_id": ccs.signal_id,
            "contractor_name": ccs.contractor_name,
            "risk_level": ccs.risk_level,
            "tenure_months": ccs.tenure_months,
            "exclusive_engagement": ccs.exclusive_engagement,
            "uses_company_equipment": ccs.uses_company_equipment,
            "has_set_hours": ccs.has_set_hours,
        })

    # Missing executed agreements
    missing = missing_executed_agreements()

    return GoldStandard(
        test_case=_TC,
        expected_outputs={
            "output_format": "exposure_schedule_and_findings_memo",
            "agreement_count": len(EMPLOYMENT_AGREEMENTS),
            "severance_exposure_count": len(SEVERANCE_EXPOSURES),
            "retention_award_count": len(RETENTION_AWARDS),
            "contractor_count": len(CONTRACTOR_CLASSIFICATION_SIGNALS),
            "total_severance_exposure": str(total_severance_exposure()),
            "total_retention_awards": str(total_retention_awards()),
            "severance_entries": severance_entries,
            "retention_entries": retention_entries,
            "contractor_entries": contractor_entries,
            "missing_executed_agreements": [ea.agreement_id for ea in missing],
            "high_risk_contractors": [c.signal_id for c in high_risk_contractors()],
            "double_count_risk": {
                "employee": "Dr. Anika Patel",
                "severance_id": "SEV-006",
                "retention_id": "RET-004",
                "severance_amount": str(SEVERANCE_EXPOSURES[5].estimated_payout),
                "retention_amount": str(RETENTION_AWARDS[3].award_amount),
                "interaction": "greater_of_not_additive",
            },
        },
        canary_verification=canary_verification,
        error_detection={},
        scoring_hints={
            "correctness": (
                "Severance computations must match agreement terms "
                "(base salary x multiplier = estimated payout)"
            ),
            "completeness": (
                "All 7 executives covered; retention/severance overlap "
                "identified; missing agreement flagged; all 4 contractors assessed"
            ),
            "double_count_avoidance": (
                "Dr. Patel's retention ($150,000) and severance ($255,000) "
                "must not be summed — agreement specifies greater-of"
            ),
            "professional_judgment": (
                "Contractor classification findings presented as signals "
                "for investigation, not definitive legal conclusions; "
                "missing agreement implications noted with appropriate caveats"
            ),
        },
        scenario_pack="ma_legal_hr_diligence",
        service_line="advisory",
        evidence_expectations={
            "exposure_ceo_golden_parachute": {
                "required_sources": ["tc20_agreement_ea_001"],
                "primary_source_required": True,
                "acceptable_terms": [
                    "change of control", "golden parachute", "3x base salary",
                ],
            },
            "exposure_double_count_risk": {
                "required_sources": [
                    "tc20_agreement_ea_006",
                    "tc20_severance_schedule",
                    "tc20_retention_plan",
                ],
                "primary_source_required": True,
                "acceptable_terms": [
                    "greater-of", "double-count", "not additive",
                    "retention", "severance",
                ],
            },
            "exposure_missing_agreement": {
                "required_sources": ["tc20_agreement_ea_006"],
                "primary_source_required": True,
                "acceptable_terms": [
                    "draft", "not executed", "missing", "IP assignment",
                ],
            },
            "exposure_contractor_classification": {
                "required_sources": ["tc20_contractor_roster"],
                "primary_source_required": True,
                "acceptable_terms": [
                    "classification", "misclassification", "signal",
                    "follow-up", "investigation",
                ],
            },
            "exposure_source_citations": {
                "required_sources": [
                    "tc20_employee_census",
                    "tc20_severance_schedule",
                ],
                "primary_source_required": True,
                "acceptable_terms": [
                    "census", "schedule", "agreement",
                ],
            },
        },
        source_requirements={
            "minimum_sources_per_finding": 1,
            "primary_source_required_for_high_risk": True,
            "management_summary_not_sufficient_alone": True,
        },
    )


# ── Public entry point ──────────────────────────────────────────────────────


def emit_tc20(
    model: CascadeModel,
    output_dir: Path,
    canaries: CanaryRegistry,
    errors: ErrorRegistry,
    manifest: Manifest,
) -> None:
    """Write all TC-20 files to *output_dir*."""
    _write_agreements(output_dir, canaries, manifest)
    _write_census(output_dir, canaries, manifest)
    _write_severance(output_dir, canaries, manifest)
    _write_retention(output_dir, canaries, manifest)
    _write_contractors(output_dir, canaries, manifest)
    _write_prompt(output_dir)
    _write_expected_behavior(output_dir)
