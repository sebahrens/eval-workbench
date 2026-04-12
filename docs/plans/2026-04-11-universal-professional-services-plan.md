---
title: Universal Professional Services Scenario Packs
date: 2026-04-11
status: draft
origin: user request to plan features for a more universal professional services synthetic document generator
---

# Universal Professional Services Scenario Packs

## Problem Frame

The current generator is strong for one deterministic Big 4 accounting universe: Cascade Industries with 18 Audit, Tax, Advisory, and Cross-service test cases. The next step is not to make the existing model arbitrarily broad. It is to add a scenario-pack layer that can support new professional-services domains while preserving the properties that make the current suite valuable:

- canonical facts as the single source of truth
- pure formatters as document views
- gold standards derived from the same facts
- deterministic canaries, planted errors, and self-testable grading
- realistic professional deliverables rather than isolated extraction tasks

The first expansion pack should be M&A legal/HR diligence because it reuses the existing Advisory/data-room foundation while adding non-accounting risk analysis, contract review, employment facts, evidence citations, and professional judgment.

## Scope

In scope:

- Add the feature architecture needed for reusable professional-services scenario packs.
- Implement one first pack: M&A legal/HR diligence around Cascade Industries.
- Add document types and golds needed to evaluate contract and HR diligence.
- Extend scoring to evaluate evidence citation, risk classification, and professional judgment.
- Keep default Cascade accounting generation deterministic and backward-compatible.

Out of scope for the first pack:

- Current legal advice or jurisdiction-specific legal compliance determinations.
- A general-purpose law engine.
- Provider-dependent LLM document augmentation as a default generation path.
- Every professional-services domain at once.
- Replacing the existing TC-01 through TC-18 structure.

## Recommended Feature Workstreams

### 1. Scenario Pack Abstraction

Introduce a scenario-pack concept above individual test cases. A pack should describe a coherent professional-services engagement family, including its canonical facts, document corpus, prompts, golds, rubrics, and capability axes.

Initial pack types:

- `cascade_accounting_core` for the existing 18 cases
- `ma_legal_hr_diligence` for the new first expansion

Key requirements:

- Existing generation path still emits TC-01 through TC-18 unchanged by default.
- A pack can register new test cases without hardcoding every TC in the orchestrator.
- A pack can share existing Cascade model facts while adding pack-specific facts.
- Pack metadata appears in `manifest.json` and scoring summaries.

Likely files:

- `generate_test_suite.py`
- `generator/formatters/__init__.py`
- `generator/golds/framework.py`
- `generator/manifest.py`
- `docs/architecture.md`
- `docs/adding-a-test-case.md`

Tests:

- `tests/test_scenario_packs.py`
- `tests/test_determinism.py`

Test scenarios:

- Default generation includes the existing 18 accounting cases with byte-identical output.
- Enabling the legal/HR diligence pack adds only its new cases and metadata.
- Pack registration order is deterministic.
- Unknown pack names fail with a clear error.

### 2. Legal/HR Diligence Canonical Facts

Extend the canonical model with diligence facts that trace back to Cascade entities and people. This should not be a separate invented universe. It should model legal and HR data around the existing acquisition/diligence context already used by TC-12.

Facts to model:

- material customer and vendor contracts
- amendments and side letters
- change-of-control and assignment restrictions
- termination, exclusivity, MFN, indemnity, limitation-of-liability, and renewal terms
- executive employment agreements
- severance and change-in-control payouts
- contractor and employee classification facts
- equity/bonus/retention arrangements
- pending HR/legal matters
- diligence request-list status and missing-document markers

Likely files:

- `generator/model/customers.py`
- new `generator/model/legal.py`
- new `generator/model/hr_diligence.py`
- `docs/canonical-model.md`

Tests:

- `tests/test_legal_model.py`
- `tests/test_hr_diligence_model.py`

Test scenarios:

- Every modeled contract clause references a known customer, vendor, employee, or entity.
- Change-of-control exposure ties to existing revenue concentration where relevant.
- Severance exposure ties to employee roster and compensation facts.
- Missing-document facts are explicit, not inferred from absent files.
- Deterministic seed produces stable contract IDs, clause IDs, and employee agreement IDs.

### 3. Professional Document Corpus Extensions

Add document emitters for legal/HR diligence evidence. These should remain formatter views over canonical facts, with deterministic presentation noise and canaries.

Initial document types:

- contract PDFs or DOCX files
- amendment and side-letter PDFs
- employment agreement DOCX/PDF files
- HR census CSV/XLSX
- equity/bonus/retention schedule XLSX
- diligence request list XLSX
- board or management summary memo PDF/DOCX
- email-like `.eml` or plain-text correspondence, if supported by the grader

Key requirements:

- Every material clause, date, amount, person, entity, and counterparty traces to the model.
- Each document has a stable source ID and canary.
- Some documents intentionally conflict with summaries to test primary-evidence use.
- Missing documents are represented in request lists and gold standards.

Likely files:

- new `generator/formatters/tc19.py`
- new `generator/formatters/tc20.py`
- new `generator/formatters/tc21.py`
- `generator/canaries.py`
- `generator/manifest.py`
- `templates/`

Tests:

- `tests/test_legal_hr_formatters.py`
- `tests/test_canaries.py`
- `tests/test_generated_files_open.py`

Test scenarios:

- Contract and HR files open cleanly in supported readers.
- Canaries remain extractable from new file types or the file type is not admitted.
- Diligence request list ties to emitted and intentionally missing documents.
- Management summaries can contain planted overstatements without changing canonical facts.

### 4. Judgment and Contradiction Registry

The existing error registry is good for numeric and document-quality errors. Legal/HR diligence also needs judgment traps and contradiction tracking.

Add a registry layer for:

- unsupported management claim
- contradiction between summary and primary contract
- missing evidence
- stale agreement version
- immaterial issue correctly deprioritized
- scope-excluded issue
- high-risk clause requiring escalation
- classification uncertainty requiring caveat rather than definitive conclusion

Key requirements:

- Judgment traps should identify the source documents that create the issue.
- Gold standards should distinguish "must flag" from "should caveat" and "do not overstate."
- The registry should be deterministic and self-testable like the current error registry.

Likely files:

- `generator/errors.py`
- new `generator/judgment_traps.py`
- `docs/canaries-and-errors.md`
- `scoring/auto_grader.py`

Tests:

- `tests/test_judgment_traps.py`
- `tests/test_auto_grader.py`

Test scenarios:

- A management summary says a contract is freely assignable, while the contract requires counterparty consent.
- A termination clause exists but is immaterial under the scenario's stated scope.
- A missing employee agreement should be listed as a diligence gap, not invented from the HR census.
- A stale draft agreement conflicts with an executed amendment; the agent should rely on the executed/latest version.

### 5. Evidence Citation and Source Traceability Scoring

Add scoring support for source-backed conclusions. Current scoring covers correctness, completeness, format, robustness, and communication. Universal professional-services cases need a more explicit evidence dimension or sub-rubric.

Capabilities to score:

- cites source file and page/sheet/row/section where practical
- distinguishes primary evidence from management summaries
- flags missing evidence instead of hallucinating
- ties risk findings to source clauses or facts
- marks assumptions separately from facts

Implementation can start as rubric and gold extensions before making a new top-level scoring dimension.

Likely files:

- `scoring/rubrics.yaml`
- `scoring/auto_grader.py`
- `docs/scoring.md`
- `generator/golds/framework.py`

Tests:

- `tests/test_auto_grader.py`
- `tests/test_scoring_template.py`

Test scenarios:

- Full-credit output cites primary contract and HR census evidence for each key finding.
- Partial-credit output gives correct risk labels but weak citations.
- Failing output cites only the management summary when primary documents contradict it.
- Missing-source output does not receive evidence credit even if the conclusion happens to be right.

### 6. First New Test Cases

Add a small, coherent set instead of a large domain expansion.

Recommended first cases:

- `TC-19`: Contract risk matrix for M&A legal diligence.
  - Inputs: 8 to 12 contracts, amendments, management summary, request list.
  - Deliverable: risk matrix with clause, source, severity, business impact, and recommended follow-up.
  - Core traps: change-of-control consent, exclusivity, MFN, stale summary, missing amendment.

- `TC-20`: HR diligence exposure summary.
  - Inputs: employee census, executive agreements, severance schedule, retention plan, contractor roster.
  - Deliverable: memo quantifying executive severance/retention exposure and flagging classification or missing-document issues.
  - Core traps: severance multiplier, missing executed agreement, contractor misclassification signal, retention double-count.

- `TC-21`: Diligence findings memo with evidence citations.
  - Inputs: curated mini data room combining legal and HR artifacts.
  - Deliverable: client-ready findings memo with evidence citations and unresolved diligence requests.
  - Core traps: summary vs source conflict, missing evidence, scope boundary, overconfident legal conclusion.

Likely files:

- `prompt.md`
- `README.md`
- `generator/formatters/tc19.py`
- `generator/formatters/tc20.py`
- `generator/formatters/tc21.py`
- `generator/golds/framework.py`
- `scoring/rubrics.yaml`

Tests:

- `tests/test_tc19_formatter.py`
- `tests/test_tc20_formatter.py`
- `tests/test_tc21_formatter.py`

Test scenarios:

- Self-test passes for all three new cases.
- Each case has at least one primary-evidence contradiction and one missing-evidence requirement.
- Each case requires citations to receive full credit.
- Each case has a professional deliverable shape that a human reviewer can evaluate.

### 7. Documentation and Capability Matrix

Update docs so this is framed as a professional-services generator with scenario packs, not just a Cascade accounting suite.

Likely files:

- `README.md`
- `docs/architecture.md`
- `docs/adding-a-test-case.md`
- `docs/scoring.md`
- `docs/glossary.md`
- new `docs/scenario-packs.md`

Tests:

- no dedicated test needed unless docs links are checked elsewhere

Test scenarios:

- README lists accounting core and legal/HR diligence as separate scenario packs.
- Adding-a-test-case docs explain when to add a new pack vs a new TC in an existing pack.
- Scoring docs explain evidence citation expectations.

## Deeper Design

### V1 Pack Interface

Use a small registry object for v1 rather than a broad plugin framework. The goal is to remove the hardcoded TC count, formatter list, and canary-key list from `generate_test_suite.py` while preserving the existing behavior.

Proposed pack metadata:

- `pack_id`: stable ID such as `cascade_accounting_core` or `ma_legal_hr_diligence`
- `display_name`: human-readable name for manifests and scoring summaries
- `test_cases`: ordered TC IDs owned by the pack
- `canary_file_keys`: deterministic list of file keys owned by the pack
- `emitters`: ordered formatter callables
- `gold_registrations`: existing `register_gold()` functions remain the gold source of truth
- `dependencies`: optional pack IDs required by this pack

V1 should avoid dynamic imports. Static registration is enough and easier to test deterministically.

Initial implementation shape:

- Move the existing 18 TC emitters and canary keys into `generator/packs/accounting_core.py`.
- Add a small registry in `generator/packs/__init__.py`.
- Add `generator/packs/legal_hr_diligence.py` after the registry exists.
- Keep `generate(config, output)` defaulting to the accounting core pack.
- Add CLI support later as `--pack` or `--packs`; do not make the new pack default until the first legal/HR pack self-test is stable.

### Directory Layout

Generated output can preserve the current layout for compatibility:

- `test_cases/TC-01` through `test_cases/TC-18` for the accounting core
- `test_cases/TC-19` through `test_cases/TC-21` for the legal/HR diligence pack

Add pack metadata rather than nesting directories by pack. This keeps the grader's `TC-*` discovery path working.

Manifest entries should grow one optional field:

- `scenario_pack`: pack ID that emitted the file

Gold standards should grow optional fields:

- `scenario_pack`
- `service_line`
- `evidence_expectations`
- `judgment_traps`
- `source_requirements`

The current gold schema should remain valid when those fields are absent.

### File Format Policy

For v1 legal/HR diligence, use only currently supported formats:

- `.xlsx`
- `.docx`
- `.pdf`
- `.csv`
- `.md`

Do not add `.eml`, `.msg`, `.html`, or `.pptx` in the first pack. Those require new file-open tests, canary extraction support, and grader behavior. Simulate email correspondence as deterministic `.md` or `.txt` first if needed.

### Legal/HR Canonical Model

Add model facts, not formatter-local text. The legal/HR pack should use deterministic dataclasses with stable IDs:

- `LegalContract`
- `ContractClause`
- `ContractAmendment`
- `LegalDiligenceIssue`
- `EmploymentAgreement`
- `RetentionAward`
- `SeveranceExposure`
- `ContractorClassificationSignal`
- `DiligenceRequest`

All references should point to existing or newly modeled IDs:

- entity code
- customer/vendor ID
- employee ID
- contract ID
- document ID
- clause ID

The first implementation can derive several facts from existing `generator/model/customers.py` and `generator/model/employees.py` while adding a compact fixed set of legal/HR objects.

### Judgment Trap Schema

Use a registry separate from `ErrorRegistry` because the shape is different from numeric/document transformations.

Suggested fields:

- `trap_id`: `JDG-001`
- `type`: `summary_contradiction`, `missing_evidence`, `stale_document`, `scope_boundary`, `overconfident_conclusion`, `immaterial_issue`
- `severity`: `high`, `medium`, `low`, or `caveat`
- `source_refs`: list of `{file_key, document_id, clause_id_or_section}`
- `expected_response`: `flag`, `caveat`, `deprioritize`, or `do_not_assert`
- `which_test_cases_should_catch`: list of TC IDs
- `description`: human-readable expectation for golds and reports

This lets scoring distinguish a missed high-risk clause from a correct decision not to overstate an immaterial issue.

### Evidence Expectations Schema

Golds for TC-19 through TC-21 should include machine-readable evidence expectations. Start simple:

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

The first grader can search agent outputs for required source file stems and key terms. That is imperfect, but better than making evidence quality entirely manual. Human scoring can still own nuanced citation quality.

### TC-19 Design: Contract Risk Matrix

Purpose:

- Test legal diligence extraction, primary-source reliance, and risk classification.

Inputs:

- 8 to 12 customer/vendor contract PDFs or DOCX files
- 2 amendments/side letters
- 1 management summary memo
- 1 diligence request list XLSX

Deliverable:

- `.xlsx` risk matrix with columns: contract, counterparty, clause, source, risk level, business impact, follow-up
- optional `.md` or `.docx` summary memo

Gold expectations:

- identify Acme change-of-control consent requirement
- identify MFN or exclusivity risk where planted
- reject or caveat management summary claim when contradicted by primary contract
- flag missing amendment as diligence gap
- include source references for each high-risk finding

### TC-20 Design: HR Diligence Exposure Summary

Purpose:

- Test people-risk diligence, compensation/severance computation, missing evidence, and caveated conclusions.

Inputs:

- employee census XLSX
- executive employment agreements DOCX/PDF
- severance schedule XLSX
- retention plan DOCX
- contractor roster CSV

Deliverable:

- `.xlsx` exposure schedule
- `.docx` or `.md` findings memo

Gold expectations:

- compute executive severance exposure from modeled salary and agreement multiplier
- avoid double-counting retention and severance where terms exclude overlap
- flag missing executed agreement for one executive
- flag contractor classification signal as follow-up, not definitive legal conclusion
- cite census and agreement sources

### TC-21 Design: Combined Diligence Findings Memo

Purpose:

- Test synthesis across legal and HR artifacts, evidence citations, unresolved request tracking, and calibrated professional language.

Inputs:

- curated mini data room from TC-19 and TC-20
- management Q&A summary
- diligence request tracker

Deliverable:

- client-ready `.docx` or `.md` findings memo with sections: executive summary, legal findings, HR findings, unresolved requests, assumptions/limitations

Gold expectations:

- include the top legal risk and top HR exposure
- identify at least one unresolved request
- distinguish management statements from source-backed findings
- avoid making definitive legal advice claims
- include evidence references for key findings

## Atomic Bead Drafts

These are written so they can be filed directly as beads after review. IDs are intentionally omitted so `bd create` can assign them.

### Pack Infrastructure

1. `Design v1 scenario pack registry contract`
   - Type: decision
   - Priority: P1
   - Parent: universal professional services expansion epic
   - Description: Define the v1 static pack registry contract for pack ID, ordered test cases, canary keys, formatter emitters, dependencies, and manifest metadata. Explicitly reject dynamic plugin loading for v1.
   - Acceptance: Decision records fields, default pack behavior, registration order, unknown-pack behavior, and why existing TC-01 through TC-18 output remains backward-compatible.

2. `Move accounting-core formatter and canary registration into a pack module`
   - Type: task
   - Priority: P1
   - Depends on: `Design v1 scenario pack registry contract`
   - Description: Move hardcoded accounting core emitter ordering and canary file keys out of `generate_test_suite.py` into `generator/packs/accounting_core.py`.
   - Acceptance: Default generation still emits TC-01 through TC-18; deterministic test passes; no output paths or canary keys change.

3. `Add scenario pack registry tests and unknown-pack validation`
   - Type: task
   - Priority: P1
   - Depends on: `Move accounting-core formatter and canary registration into a pack module`
   - Description: Add tests for pack lookup, deterministic order, dependency validation, and unknown pack failure.
   - Acceptance: `tests/test_scenario_packs.py` covers default accounting core, explicit pack lookup, duplicate IDs, and unknown IDs.

4. `Add optional scenario_pack metadata to manifest entries`
   - Type: task
   - Priority: P2
   - Depends on: `Design v1 scenario pack registry contract`
   - Description: Extend `ManifestEntry` and manifest JSON output with optional pack metadata while preserving compatibility for existing entries.
   - Acceptance: Existing manifest tests pass after updating expected schema; new tests assert pack metadata is stable and sorted deterministically.

5. `Add CLI support for selecting scenario packs`
   - Type: feature
   - Priority: P2
   - Depends on: `Add scenario pack registry tests and unknown-pack validation`
   - Description: Add `--pack` or `--packs` to `generate_test_suite.py` so the legal/HR pack can be generated intentionally without changing default output.
   - Acceptance: Default command emits accounting core only; explicit legal/HR pack command emits TC-19 through TC-21 after pack implementation; unknown pack exits with a clear error.

### Legal/HR Canonical Facts

6. `Design legal and HR diligence canonical dataclasses`
   - Type: decision
   - Priority: P1
   - Parent: universal professional services expansion epic
   - Description: Define stable dataclasses and ID conventions for contracts, clauses, amendments, employment agreements, retention awards, severance exposures, contractor signals, and diligence requests.
   - Acceptance: Decision records dataclasses, required references to existing model IDs, ID naming conventions, and non-goals for jurisdiction-specific legal advice.

7. `Implement legal contract and clause canonical model`
   - Type: feature
   - Priority: P1
   - Depends on: `Design legal and HR diligence canonical dataclasses`
   - Description: Add `generator/model/legal.py` with deterministic contracts, clauses, amendments, source IDs, and modeled legal diligence issues tied to Cascade customers/vendors/entities.
   - Acceptance: Tests prove stable IDs, known counterparties, at least one change-of-control risk, one MFN/exclusivity-style risk, one stale-summary contradiction, and one missing amendment gap.

8. `Implement HR diligence canonical model`
   - Type: feature
   - Priority: P1
   - Depends on: `Design legal and HR diligence canonical dataclasses`
   - Description: Add `generator/model/hr_diligence.py` with executive agreements, severance exposure, retention awards, contractor classification signals, and missing-document facts tied to employees.
   - Acceptance: Tests prove severance exposure ties to salary/employee IDs, retention does not double-count where excluded, one missing executed agreement exists, and contractor signals are caveated follow-up items.

9. `Attach legal and HR diligence facts to CascadeModel`
   - Type: task
   - Priority: P1
   - Depends on: `Implement legal contract and clause canonical model`, `Implement HR diligence canonical model`
   - Description: Extend `CascadeModel` and `build_model()` to include legal and HR diligence facts without perturbing existing default accounting facts.
   - Acceptance: Existing model tests pass; new tests verify fields exist; determinism test catches no drift in existing TC-01 through TC-18 outputs.

### Judgment, Gold, and Scoring

10. `Implement judgment trap registry`
    - Type: feature
    - Priority: P1
    - Depends on: `Design legal and HR diligence canonical dataclasses`
    - Description: Add a deterministic judgment trap registry for non-numeric professional judgment issues such as summary contradiction, missing evidence, stale document, scope boundary, and overconfident conclusion.
    - Acceptance: Tests cover JSON serialization, deterministic ordering, required fields, and filtering traps by test case.

11. `Extend gold schema for scenario packs and evidence expectations`
    - Type: feature
    - Priority: P1
    - Depends on: `Implement judgment trap registry`
    - Description: Extend `GoldStandard` with optional scenario pack metadata, evidence expectations, source requirements, and judgment trap expectations.
    - Acceptance: Existing gold JSON remains compatible; round-trip tests cover new optional fields; self-test can read old and new golds.

12. `Add evidence expectation checks to auto grader`
    - Type: feature
    - Priority: P2
    - Depends on: `Extend gold schema for scenario packs and evidence expectations`
    - Description: Add a lightweight evidence check that searches agent outputs for required source references and key terms from gold evidence expectations.
    - Acceptance: Tests cover full credit for primary-source citation, partial/failed evidence checks, and no regression for old golds without evidence expectations.

13. `Update rubrics for evidence-backed professional judgment`
    - Type: task
    - Priority: P2
    - Depends on: `Add evidence expectation checks to auto grader`
    - Description: Update `scoring/rubrics.yaml` and scoring docs to explain evidence expectations and professional judgment scoring for TC-19 through TC-21.
    - Acceptance: Rubric contains anchors for legal/HR diligence cases; docs explain primary-source reliance, missing evidence, assumptions, and caveated conclusions.

### Legal/HR Documents and Test Cases

14. `Add legal/HR document writer helpers`
    - Type: task
    - Priority: P1
    - Depends on: `Implement legal contract and clause canonical model`, `Implement HR diligence canonical model`
    - Description: Add deterministic helper functions for contract PDFs/DOCX, amendments, employment agreements, HR census schedules, retention schedules, and diligence request lists using only supported file formats.
    - Acceptance: File integrity tests open generated files; canaries are supported; document properties are deterministic.

15. `Implement TC-19 contract risk matrix formatter and prompt`
    - Type: feature
    - Priority: P1
    - Depends on: `Add legal/HR document writer helpers`, `Extend gold schema for scenario packs and evidence expectations`
    - Description: Add TC-19 input corpus and prompt for contract risk matrix legal diligence.
    - Acceptance: Formatter emits contracts/amendments/summary/request list; prompt requests risk matrix; canaries embedded; self-test can find inputs.

16. `Implement TC-19 gold and rubric`
    - Type: task
    - Priority: P1
    - Depends on: `Implement TC-19 contract risk matrix formatter and prompt`
    - Description: Add TC-19 gold expectations for change-of-control, MFN/exclusivity, stale summary, missing amendment, and source evidence.
    - Acceptance: TC-19 self-test passes; gold has evidence expectations and judgment traps; rubric anchors are concrete.

17. `Implement TC-20 HR diligence exposure formatter and prompt`
    - Type: feature
    - Priority: P1
    - Depends on: `Add legal/HR document writer helpers`, `Extend gold schema for scenario packs and evidence expectations`
    - Description: Add TC-20 input corpus and prompt for HR diligence exposure summary.
    - Acceptance: Formatter emits census, agreements, severance schedule, retention plan, and contractor roster; prompt requests exposure schedule and memo; canaries embedded.

18. `Implement TC-20 gold and rubric`
    - Type: task
    - Priority: P1
    - Depends on: `Implement TC-20 HR diligence exposure formatter and prompt`
    - Description: Add TC-20 gold expectations for severance exposure, retention double-count avoidance, missing executed agreement, contractor follow-up, and source evidence.
    - Acceptance: TC-20 self-test passes; gold has evidence expectations and judgment traps; rubric anchors are concrete.

19. `Implement TC-21 combined diligence findings formatter and prompt`
    - Type: feature
    - Priority: P2
    - Depends on: `Implement TC-19 gold and rubric`, `Implement TC-20 gold and rubric`
    - Description: Add TC-21 combined mini data room and prompt for a client-ready legal/HR diligence findings memo.
    - Acceptance: Formatter emits curated combined inputs; prompt requires findings, unresolved requests, assumptions/limitations, and citations.

20. `Implement TC-21 gold and rubric`
    - Type: task
    - Priority: P2
    - Depends on: `Implement TC-21 combined diligence findings formatter and prompt`
    - Description: Add TC-21 gold expectations for synthesis across legal and HR artifacts, unresolved request tracking, source-backed findings, and calibrated professional language.
    - Acceptance: TC-21 self-test passes; gold has evidence expectations and judgment traps; rubric anchors are concrete.

### End-to-End Gates and Docs

21. `Add legal/HR pack end-to-end generation and self-test gate`
    - Type: task
    - Priority: P1
    - Depends on: `Implement TC-19 gold and rubric`, `Implement TC-20 gold and rubric`, `Implement TC-21 gold and rubric`
    - Description: Add an end-to-end test that generates the legal/HR pack, runs self-test, and verifies determinism without altering accounting-core output.
    - Acceptance: Test proves TC-19 through TC-21 are generated when enabled, self-test passes, and default accounting-core output remains byte-identical.

22. `Document scenario packs and legal/HR diligence extension`
    - Type: task
    - Priority: P2
    - Depends on: `Add legal/HR pack end-to-end generation and self-test gate`
    - Description: Update README and docs for scenario packs, TC-19 through TC-21, evidence expectations, and when to add a new pack vs a new test case.
    - Acceptance: Docs list pack IDs, generation commands, new test cases, scoring changes, and file format policy.

## Sequencing

1. Define scenario-pack metadata and backward-compatible registration.
2. Add legal/HR canonical facts with deterministic invariants.
3. Add document emitters and canary support for the first legal/HR file set.
4. Add judgment-trap registry and gold-standard fields.
5. Extend scoring/rubrics for evidence citation and primary-source reliance.
6. Implement TC-19, then TC-20, then TC-21.
7. Update docs and run full determinism/self-test gates.

## Risks

- Over-generalizing too early: keep the first abstraction only as broad as needed for accounting core plus legal/HR diligence.
- Turning legal diligence into legal advice: prompts and rubrics should ask for diligence risk identification and follow-up questions, not definitive legal opinions.
- Scoring citations too rigidly: start with source-file and section/row-level expectations before trying to parse every page citation perfectly.
- New file formats can break determinism or canary extraction: admit them only after canary and open-file tests exist.
- Duplicating facts outside the model: legal/HR facts must reference existing entities, customers, vendors, and employees where possible.

## Open Questions

- Should TC-19 through TC-21 ship in the default generated suite, or behind a `--pack legal_hr_diligence` option first?
- Should evidence citation become a sixth scoring dimension, or remain a sub-requirement under correctness/completeness/communication for v1?
- Which new file formats are worth supporting first: `.eml`, `.html`, `.pptx`, or only existing `.xlsx`, `.docx`, `.pdf`, and `.csv`?
- How much jurisdiction specificity should be included in prompts without pretending to provide legal advice?
