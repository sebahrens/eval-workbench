# Universal Professional Services Scenario Packs

Status: Draft
Last updated: 2026-04-11
Primary tracker epic: `synth-data-ups`

## Summary

Expand the Cascade Industries generator from a single accounting-focused Big 4 suite into a reusable professional-services scenario-pack system. The first new pack is M&A legal/HR diligence, adding TC-19 through TC-21 while preserving the current TC-01 through TC-18 accounting core.

The extension must preserve the current generator's core invariants:

- deterministic output from a seeded canonical model
- formatters as pure views over canonical facts
- gold standards derived from the same canonical facts as inputs
- canaries for source-provenance checks
- planted errors and judgment traps that are recorded and testable
- backward-compatible default generation for the current suite

## Goals

- Add a v1 static scenario-pack registry.
- Move the existing accounting-core registration out of `generate_test_suite.py`.
- Add legal/HR canonical facts tied to the existing Cascade model.
- Add judgment traps for professional-services reasoning that is not purely numeric.
- Add evidence expectations to gold standards and scoring.
- Add TC-19, TC-20, and TC-21 for M&A legal/HR diligence.
- Document pack behavior, scoring changes, and supported file formats.

## Non-Goals

- Do not build a dynamic plugin system in v1.
- Do not add `.eml`, `.msg`, `.html`, or `.pptx` support in the first legal/HR pack.
- Do not make legal/HR diligence part of default generation until its self-test is stable.
- Do not make definitive legal advice claims in prompts, golds, or rubrics.
- Do not duplicate facts in formatters that belong in the canonical model.

## Scenario Pack Contract

V1 scenario packs are statically registered. A pack has:

- `pack_id`: stable identifier, for example `cascade_accounting_core` or `ma_legal_hr_diligence`
- `display_name`: human-readable pack name
- `test_cases`: ordered test case IDs
- `canary_file_keys`: deterministic file-key list used to build the canary registry
- `emitters`: ordered formatter callables
- `dependencies`: optional pack IDs required by this pack

The initial implementation should add:

- `generator/packs/__init__.py`
- `generator/packs/accounting_core.py`
- `generator/packs/legal_hr_diligence.py`

`generate_test_suite.py` should stop owning the full formatter and canary list directly. Default generation remains accounting core only.

## Output Layout

Keep the existing generated output shape:

- `test_cases/TC-01` through `test_cases/TC-18`: accounting core
- `test_cases/TC-19` through `test_cases/TC-21`: legal/HR diligence when enabled
- `gold_standards/TC-*_gold.json`: all generated golds

Do not nest test cases by pack in v1. The current grader discovers `TC-*` golds and should keep working.

## Manifest Extension

Manifest entries may include:

- `scenario_pack`: optional pack ID that emitted the file

The field must be deterministic and backward-compatible. Existing consumers should tolerate its absence or presence.

## Gold Schema Extension

Existing gold standards remain valid. New gold standards may include:

- `scenario_pack`
- `service_line`
- `evidence_expectations`
- `judgment_traps`
- `source_requirements`

Example evidence expectation:

```json
{
  "evidence_expectations": {
    "risk_change_of_control": {
      "required_sources": ["tc19_contract_acme", "tc19_acme_amendment_2025"],
      "primary_source_required": true,
      "acceptable_terms": ["change of control", "assignment", "counterparty consent"]
    }
  }
}
```

## Legal/HR Canonical Model

Add canonical objects for M&A legal/HR diligence. Suggested dataclasses:

- `LegalContract`
- `ContractClause`
- `ContractAmendment`
- `LegalDiligenceIssue`
- `EmploymentAgreement`
- `RetentionAward`
- `SeveranceExposure`
- `ContractorClassificationSignal`
- `DiligenceRequest`

All objects should use stable IDs and reference existing Cascade facts where possible:

- entity code
- customer/vendor ID
- employee ID
- contract ID
- document ID
- clause ID

Initial modules:

- `generator/model/legal.py`
- `generator/model/hr_diligence.py`

`CascadeModel` in `generator/model/build.py` should include legal and HR diligence fields after the model modules are implemented.

## Judgment Trap Registry

Use a separate registry from `ErrorRegistry` for professional judgment issues.

Fields:

- `trap_id`: for example `JDG-001`
- `type`: one of `summary_contradiction`, `missing_evidence`, `stale_document`, `scope_boundary`, `overconfident_conclusion`, `immaterial_issue`
- `severity`: `high`, `medium`, `low`, or `caveat`
- `source_refs`: source files, document IDs, clause IDs, or sections
- `expected_response`: `flag`, `caveat`, `deprioritize`, or `do_not_assert`
- `which_test_cases_should_catch`: TC IDs
- `description`: human-readable expected response

The registry must serialize deterministically and be filterable by test case.

## File Format Policy

V1 legal/HR diligence may use:

- `.xlsx`
- `.docx`
- `.pdf`
- `.csv`
- `.md`
- `.txt`

V1 must not introduce:

- `.eml`
- `.msg`
- `.html`
- `.pptx`

Those formats require separate canary extraction, file-open, and grader support.

## Test Cases

### TC-19: Contract Risk Matrix

Purpose: Test legal diligence extraction, source reliance, and risk classification.

Inputs:

- 8 to 12 customer/vendor contracts
- 2 amendments or side letters
- 1 management summary memo
- 1 diligence request list

Deliverable:

- risk matrix with contract, counterparty, clause, source, risk level, business impact, and follow-up
- optional summary memo

Required traps:

- change-of-control consent requirement
- MFN or exclusivity risk
- management summary contradicted by primary contract
- missing amendment or side letter
- required source reference for high-risk findings

### TC-20: HR Diligence Exposure Summary

Purpose: Test people-risk diligence, severance/retention computation, missing evidence, and calibrated follow-up.

Inputs:

- employee census
- executive employment agreements
- severance schedule
- retention plan
- contractor roster

Deliverable:

- exposure schedule
- findings memo

Required traps:

- executive severance exposure tied to salary and agreement multiplier
- retention/severance double-count avoidance
- missing executed agreement
- contractor classification signal as follow-up, not definitive legal conclusion
- census and agreement source citations

### TC-21: Combined Diligence Findings Memo

Purpose: Test synthesis across legal and HR artifacts, evidence citations, unresolved request tracking, and calibrated professional language.

Inputs:

- curated mini data room from TC-19 and TC-20
- management Q&A summary
- diligence request tracker

Deliverable:

- client-ready findings memo with executive summary, legal findings, HR findings, unresolved requests, and assumptions/limitations

Required traps:

- top legal risk
- top HR exposure
- unresolved diligence request
- management-statement versus source-backed finding distinction
- no definitive legal advice language

## Atomic Beads

The implementation backlog is filed under `synth-data-ups`:

- `synth-data-ups.1`: Design v1 scenario pack registry contract
- `synth-data-ups.2`: Move accounting-core formatter and canary registration into a pack module
- `synth-data-ups.3`: Add scenario pack registry tests and unknown-pack validation
- `synth-data-ups.4`: Add optional scenario_pack metadata to manifest entries
- `synth-data-ups.5`: Add CLI support for selecting scenario packs
- `synth-data-ups.6`: Design legal and HR diligence canonical dataclasses
- `synth-data-ups.7`: Implement legal contract and clause canonical model
- `synth-data-ups.8`: Implement HR diligence canonical model
- `synth-data-ups.9`: Attach legal and HR diligence facts to CascadeModel
- `synth-data-ups.10`: Implement judgment trap registry
- `synth-data-ups.11`: Extend gold schema for scenario packs and evidence expectations
- `synth-data-ups.12`: Add evidence expectation checks to auto grader
- `synth-data-ups.13`: Update rubrics for evidence-backed professional judgment
- `synth-data-ups.14`: Add legal/HR document writer helpers
- `synth-data-ups.15`: Implement TC-19 contract risk matrix formatter and prompt
- `synth-data-ups.16`: Implement TC-19 gold and rubric
- `synth-data-ups.17`: Implement TC-20 HR diligence exposure formatter and prompt
- `synth-data-ups.18`: Implement TC-20 gold and rubric
- `synth-data-ups.19`: Implement TC-21 combined diligence findings formatter and prompt
- `synth-data-ups.20`: Implement TC-21 gold and rubric
- `synth-data-ups.21`: Add legal/HR pack end-to-end generation and self-test gate
- `synth-data-ups.22`: Document scenario packs and legal/HR diligence extension

## Validation Gates

Before closing the epic:

- `bd lint` has no template warnings for the new beads.
- `bd dep cycles` has no cycles.
- Default generation remains deterministic for accounting core.
- Legal/HR pack generation and self-test pass.
- Docs link from `SPEC.md`, `README.md`, `docs/architecture.md`, and `docs/scoring.md`.
