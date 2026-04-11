"""Formatter: TC-19 — Contract Risk Matrix (Legal Diligence).

Emits:
- test_cases/TC-19/input_files/contracts/
    10 docx contract summaries (one per LCTR-NNN)
- test_cases/TC-19/input_files/amendments/
    3 docx amendment / side letter documents
- test_cases/TC-19/input_files/management_summary_memo.docx
    Management summary (deliberately stale re AMD-002)
- test_cases/TC-19/input_files/diligence_request_list.xlsx
    Diligence request tracker
- test_cases/TC-19/prompt.md
- test_cases/TC-19/expected_behavior.md
- gold_standards/TC-19_gold.json

Judgment traps:
  - Change-of-control consent requirement (LDI-001)
  - MFN / exclusivity risk (LDI-002, LDI-003)
  - Management summary contradicted by primary contract (LDI-002)
  - Missing amendment or side letter (LDI-003)
  - Required source reference for high-risk findings (LDI-004)

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
from generator.model.legal import (
    CONTRACT_AMENDMENTS,
    CONTRACT_CLAUSES,
    LEGAL_CONTRACTS,
    LEGAL_DILIGENCE_ISSUES,
    amendments_for_contract,
    clauses_for_contract,
)
from generator.writers.legal import (
    write_all_amendments,
    write_all_contract_summaries,
    write_diligence_request_list,
    write_management_summary_memo,
)

# ── Constants ────────────────────────────────────────────────────────────────

_TC = "TC-19"
_INPUT_DIR = f"test_cases/{_TC}/input_files"

# Canary file keys — 10 contracts + 3 amendments + 1 memo + 1 request list = 15
_CANARY_KEYS: list[str] = sorted([
    # Contract summaries (one per LCTR-NNN)
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
    # Amendments / side letters
    "tc19_amendment_amd_001",
    "tc19_amendment_amd_002",
    "tc19_amendment_amd_003",
    # Management summary memo
    "tc19_management_summary_memo",
    # Diligence request list
    "tc19_diligence_request_list",
])


# ── File writers ─────────────────────────────────────────────────────────────


def _write_contracts(
    output_dir: Path,
    canaries: CanaryRegistry,
    manifest: Manifest,
) -> None:
    """Write one docx per LegalContract into TC-19/input_files/contracts/."""
    contracts_dir = output_dir / _INPUT_DIR / "contracts"
    contracts_dir.mkdir(parents=True, exist_ok=True)

    canary_keys = {
        c.contract_id: f"tc19_contract_{c.contract_id.lower().replace('-', '_')}"
        for c in LEGAL_CONTRACTS
    }

    locations = write_all_contract_summaries(contracts_dir, canaries, canary_keys)

    for contract in LEGAL_CONTRACTS:
        ckey = canary_keys[contract.contract_id]
        fname = f"contract_{contract.contract_id.lower()}.docx"
        rel_path = f"{_INPUT_DIR}/contracts/{fname}"
        canaries.set_location(ckey, rel_path, locations[contract.contract_id])
        manifest.register(rel_path, "docx", test_cases=[_TC])


def _write_amendments(
    output_dir: Path,
    canaries: CanaryRegistry,
    manifest: Manifest,
) -> None:
    """Write one docx per ContractAmendment into TC-19/input_files/amendments/."""
    amendments_dir = output_dir / _INPUT_DIR / "amendments"
    amendments_dir.mkdir(parents=True, exist_ok=True)

    canary_keys = {
        a.amendment_id: f"tc19_amendment_{a.amendment_id.lower().replace('-', '_')}"
        for a in CONTRACT_AMENDMENTS
    }

    locations = write_all_amendments(amendments_dir, canaries, canary_keys)

    for amd in CONTRACT_AMENDMENTS:
        ckey = canary_keys[amd.amendment_id]
        fname = f"amendment_{amd.amendment_id.lower()}.docx"
        rel_path = f"{_INPUT_DIR}/amendments/{fname}"
        canaries.set_location(ckey, rel_path, locations[amd.amendment_id])
        manifest.register(rel_path, "docx", test_cases=[_TC])


def _write_memo(
    output_dir: Path,
    canaries: CanaryRegistry,
    manifest: Manifest,
) -> None:
    """Write the management summary memo (deliberately stale)."""
    memo_path = output_dir / _INPUT_DIR / "management_summary_memo.docx"
    memo_path.parent.mkdir(parents=True, exist_ok=True)

    ckey = "tc19_management_summary_memo"
    location = write_management_summary_memo(memo_path, canaries, ckey)
    canaries.set_location(ckey, f"{_INPUT_DIR}/management_summary_memo.docx", location)
    manifest.register(f"{_INPUT_DIR}/management_summary_memo.docx", "docx", test_cases=[_TC])


def _write_request_list(
    output_dir: Path,
    canaries: CanaryRegistry,
    manifest: Manifest,
) -> None:
    """Write the diligence request tracker spreadsheet."""
    xlsx_path = output_dir / _INPUT_DIR / "diligence_request_list.xlsx"
    xlsx_path.parent.mkdir(parents=True, exist_ok=True)

    ckey = "tc19_diligence_request_list"
    location = write_diligence_request_list(xlsx_path, canaries, ckey)
    canaries.set_location(ckey, f"{_INPUT_DIR}/diligence_request_list.xlsx", location)
    manifest.register(f"{_INPUT_DIR}/diligence_request_list.xlsx", "xlsx", test_cases=[_TC])


# ── Prompt & Expected Behavior ───────────────────────────────────────────────


def _write_prompt(output_dir: Path) -> None:
    """Write test_cases/TC-19/prompt.md per spec."""
    text = """\
Review the contract portfolio for Cascade Industries in connection with
a potential acquisition.

Using the contracts, amendments, side letters, management summary memo,
and diligence request list provided:

1. Build a **contract risk matrix** with the following columns:
   - Contract ID
   - Counterparty
   - Key Clause (clause type and section reference)
   - Source Document (cite the specific contract, amendment, or memo)
   - Risk Level (high / medium / low)
   - Business Impact (revenue at risk, operational, compliance, retention)
   - Recommended Follow-Up (escalate, investigate, request document, flag)

2. For each high-risk finding, cite the **primary source document**
   (contract or amendment), not just the management summary.

3. Identify any cases where the management summary memo **contradicts
   or omits** information found in the primary contracts or amendments.

4. Flag any contracts requiring **consent or novation** for a change
   of ownership.

5. Note any **missing amendments or side letters** referenced in the
   contracts but not included in the diligence binder.

6. Optionally, produce a brief **summary memo** highlighting the top
   3 risks and recommended next steps.
"""
    path = output_dir / f"test_cases/{_TC}/prompt.md"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text)


def _write_expected_behavior(output_dir: Path) -> None:
    """Write test_cases/TC-19/expected_behavior.md per spec."""
    text = """\
# TC-19: Contract Risk Matrix — Expected Behavior

## Risk Matrix Requirements
- The risk matrix must include all 10 contracts (LCTR-001 through LCTR-010).
- Each row must cite a specific clause (CLS-NNN) and source section.
- Risk levels must be justified by the clause content, not assumed.
- High-risk items must cite the primary contract or amendment, not
  only the management summary memo.

## Change-of-Control (CLS-001 / LDI-001)
- Acme Manufacturing (LCTR-001) has a change-of-control termination
  clause in Section 14.2 — rated high risk.
- No waiver or consent has been obtained from Acme.
- Revenue at risk: approximately $36M/year (~18% of consolidated revenue).
- The agent must flag this as requiring consent or novation pre-close.

## MFN Pricing Contradiction (CLS-003 / LDI-002)
- TechAlloy Systems (LCTR-003) has an MFN clause (Section 8.1).
- The management summary memo describes the MFN as triggering on "any
  pricing difference."
- Side letter AMD-002 (effective 2025-01-15) raised the threshold to
  differences exceeding 5% per kg.
- The agent must identify that the management summary contradicts the
  current contractual terms as amended.
- Source reliance: the agent must cite AMD-002, not just the memo.

## Exclusivity Scope Expansion (CLS-005 / LDI-003)
- NextGen Composites (LCTR-004) has an exclusivity clause (Section 6.4).
- Amendment AMD-003 (effective 2025-03-01) expanded exclusivity to
  thermal barrier coatings, not just alloy specs.
- The management summary memo does not mention AMD-003.
- The agent should flag the summary as stale regarding exclusivity scope.

## Government Subcontract Assignment (CLS-006 / LDI-004)
- NextGen government subcontract requires prior written consent for
  assignment (Section 22.1).
- Change of ownership may trigger FAR 42.12 novation.
- The agent must flag this as requiring consent pre-close.

## IP License Dependency (CLS-008 / LDI-005)
- ChemSource (LCTR-010) IP license survives termination for 36 months.
- If the supply relationship ends post-acquisition, the advanced
  materials product line loses access to proprietary formulations.
- The agent should flag this as a strategic dependency, not a defect.

## Source Citation Requirements
- Every high-risk finding must reference the primary contract/amendment.
- Findings that rely solely on the management summary should be
  flagged as potentially unreliable.
- The diligence request list should be referenced for open items.

## Scoring Focus
Grade on:
- Completeness of risk matrix (all 10 contracts covered)
- Accuracy of risk levels and clause identification
- Detection of management summary contradictions (LDI-002, LDI-003)
- Proper source citation (primary documents over summary)
- Identification of consent/novation requirements (LDI-001, LDI-004)
- Professional judgment on IP license dependency (LDI-005)
"""
    path = output_dir / f"test_cases/{_TC}/expected_behavior.md"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text)


# ── Gold Standard ────────────────────────────────────────────────────────────


@register_gold("TC-19")
def _tc19_gold(
    canaries: CanaryRegistry,
    errors: ErrorRegistry,
    **model_kwargs: Any,
) -> GoldStandard:
    """Build the TC-19 gold standard."""
    # Build canary verification from all TC-19 file keys
    canary_verification: dict[str, str] = {}
    for contract in LEGAL_CONTRACTS:
        ckey = f"tc19_contract_{contract.contract_id.lower().replace('-', '_')}"
        label = f"read_contract_{contract.contract_id.lower().replace('-', '_')}"
        canary_verification[label] = canaries.canary_for(ckey)
    for amd in CONTRACT_AMENDMENTS:
        ckey = f"tc19_amendment_{amd.amendment_id.lower().replace('-', '_')}"
        label = f"read_amendment_{amd.amendment_id.lower().replace('-', '_')}"
        canary_verification[label] = canaries.canary_for(ckey)
    canary_verification["read_management_summary_memo"] = canaries.canary_for(
        "tc19_management_summary_memo"
    )
    canary_verification["read_diligence_request_list"] = canaries.canary_for(
        "tc19_diligence_request_list"
    )

    # Build expected risk matrix entries from canonical data
    risk_matrix_entries = []
    for contract in LEGAL_CONTRACTS:
        clauses = clauses_for_contract(contract.contract_id)
        amendments = amendments_for_contract(contract.contract_id)
        for clause in clauses:
            entry = {
                "contract_id": contract.contract_id,
                "counterparty": contract.counterparty_name,
                "clause_id": clause.clause_id,
                "clause_type": clause.clause_type,
                "source_section": clause.source_section,
                "risk_level": clause.risk_level,
                "business_impact": clause.business_impact,
            }
            # Add amendment references if any modify this clause
            related_amendments = [
                a.amendment_id for a in amendments
                if clause.clause_id in a.changes_clause_ids
            ]
            if related_amendments:
                entry["related_amendments"] = related_amendments
            risk_matrix_entries.append(entry)

    # Build diligence issues summary
    diligence_findings = []
    for issue in LEGAL_DILIGENCE_ISSUES:
        diligence_findings.append({
            "issue_id": issue.issue_id,
            "contract_id": issue.contract_id,
            "issue_type": issue.issue_type,
            "severity": issue.severity,
            "recommended_action": issue.recommended_action,
        })

    return GoldStandard(
        test_case=_TC,
        expected_outputs={
            "output_format": "risk_matrix",
            "contract_count": len(LEGAL_CONTRACTS),
            "clause_count": len(CONTRACT_CLAUSES),
            "amendment_count": len(CONTRACT_AMENDMENTS),
            "risk_matrix_entries": risk_matrix_entries,
            "diligence_findings": diligence_findings,
            "high_risk_contracts": [
                c.contract_id for c in LEGAL_CONTRACTS
                if any(
                    cl.risk_level == "high"
                    for cl in clauses_for_contract(c.contract_id)
                )
            ],
            "consent_required_contracts": ["LCTR-001", "LCTR-004"],
            "summary_contradictions": ["LDI-002", "LDI-003"],
        },
        canary_verification=canary_verification,
        error_detection={},
        scoring_hints={
            "correctness": (
                "Risk matrix must cover all 10 contracts with accurate "
                "clause identification and risk levels"
            ),
            "completeness": (
                "All 5 diligence issues (LDI-001 through LDI-005) identified; "
                "all consent/novation requirements flagged"
            ),
            "source_reliance": (
                "High-risk findings cite primary contracts/amendments, "
                "not just management summary; summary contradictions detected"
            ),
            "professional_judgment": (
                "IP license dependency flagged as strategic risk; "
                "exclusivity scope expansion noted; appropriate caveats used"
            ),
        },
        scenario_pack="ma_legal_hr_diligence",
        service_line="advisory",
        evidence_expectations={
            "risk_change_of_control": {
                "required_sources": ["tc19_contract_lctr_001"],
                "primary_source_required": True,
                "acceptable_terms": [
                    "change of control", "assignment", "counterparty consent",
                ],
            },
            "risk_mfn_contradiction": {
                "required_sources": [
                    "tc19_contract_lctr_003",
                    "tc19_amendment_amd_002",
                ],
                "primary_source_required": True,
                "acceptable_terms": [
                    "most-favored-nation", "MFN", "pricing threshold",
                ],
            },
            "risk_exclusivity_expansion": {
                "required_sources": [
                    "tc19_contract_lctr_004",
                    "tc19_amendment_amd_003",
                ],
                "primary_source_required": True,
                "acceptable_terms": [
                    "exclusivity", "thermal barrier coatings", "scope expansion",
                ],
            },
            "risk_govt_assignment": {
                "required_sources": ["tc19_contract_lctr_004"],
                "primary_source_required": True,
                "acceptable_terms": [
                    "assignment", "novation", "FAR 42.12", "consent",
                ],
            },
            "risk_ip_dependency": {
                "required_sources": ["tc19_contract_lctr_010"],
                "primary_source_required": True,
                "acceptable_terms": [
                    "IP license", "survival period", "proprietary formulations",
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


def emit_tc19(
    model: CascadeModel,
    output_dir: Path,
    canaries: CanaryRegistry,
    errors: ErrorRegistry,
    manifest: Manifest,
) -> None:
    """Write all TC-19 files to *output_dir*."""
    _write_contracts(output_dir, canaries, manifest)
    _write_amendments(output_dir, canaries, manifest)
    _write_memo(output_dir, canaries, manifest)
    _write_request_list(output_dir, canaries, manifest)
    _write_prompt(output_dir)
    _write_expected_behavior(output_dir)
