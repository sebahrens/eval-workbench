# Scoring

This document explains how the test suite grades agent outputs.

## The five dimensions

Every test case is scored on **five dimensions**, each on a **1–3 scale**:

| Dimension | Score 3 | Score 2 | Score 1 |
|---|---|---|---|
| **Correctness** | All values match the gold standard within ±0.5% tolerance | Minor errors that don't change conclusions | Material errors in key outputs |
| **Completeness** | All requested deliverables produced, all components present | Missing minor elements (one sheet, one section) | Missing major deliverables or sections |
| **Format Compliance** | Valid files, professional formatting, correct file types | Functional but poorly formatted, or minor type issues | Broken files, wrong file type, unusable layout |
| **Robustness** | Handled all edge cases: messy data, format variations, ambiguity, planted errors flagged | Handled most edge cases, stumbled on some | Failed on edge cases or didn't attempt to handle them |
| **Communication** | Proactively explained approach, flagged uncertainties, identified errors | Adequate communication, some gaps | Silent about approach, missed obvious items to flag |

A single test case produces a score tuple like `3/3/2/3/3` (Correctness=3, Completeness=3, Format=2, Robustness=3, Communication=3).

## Pass / fail thresholds

From spec §7.4:

| Threshold | Rule |
|---|---|
| **Pass** | Average score ≥ 2.4 across all 5 dimensions |
| **Conditional Pass** | Average score ≥ 2.0 AND no dimension scored 1 |
| **Fail** | Average score < 2.0 OR any dimension scored 1 |

The **any-dimension-1-is-automatic-fail** rule matters. A test case scoring `3/3/3/1/3` has an average of 2.6 (which looks like a pass) but fails because Robustness is 1. This is deliberate: a TC that produces correct numbers but missed a planted error isn't acceptable — the whole point of the adversarial inputs is that the agent must catch them.

## Mechanical vs human grading

The suite grades test cases in two passes:

### 1. Auto-grader (mechanical dimensions)

`scoring/auto_grader.py` handles everything a machine can evaluate reliably:

```bash
# Grade a single test case
uv run python -m scoring.auto_grader \
    --tc TC-01 \
    --gold /tmp/test_suite/gold_standards/TC-01_gold.json \
    --agent-output /path/to/agent/output/TC-01

# Grade every test case
uv run python -m scoring.auto_grader \
    --suite-dir /tmp/test_suite \
    --agent-output-dir /path/to/agent/outputs

# Self-test (grade the gold standards against themselves; must score 3/3/3/3/3)
uv run python -m scoring.auto_grader \
    --self-test --suite-dir /tmp/test_suite
```

The auto-grader evaluates:

- **Correctness**: compares numerical values to gold with ±0.5% tolerance
- **Completeness**: checks presence of required files, sheets, sections
- **Format**: validates file types can be opened without errors
- **Canary verification**: checks if the agent read the correct files
- **Error detection**: checks if the agent flagged planted errors

Output: one JSON report per test case + an aggregate summary.

#### Normal grading vs self-test: evidence semantics

The auto-grader has two distinct modes that answer different questions, and they look for evidence in different places. Confusing the two leads to false passes.

**Normal grading** (`--agent-output`) answers: *Did the agent produce correct results from the test inputs?*

Every check searches **only the agent's output directory** for evidence:

| Check | What it searches | What it proves |
|---|---|---|
| Canary verification | Agent output files | The agent actually read the correct input files (not other files or cached data) |
| Error detection | Agent output files | The agent flagged planted errors in its response |
| Evidence (TC-19+) | Agent output files | The agent cited required sources and used calibrated terms |
| Correctness | Agent output JSON/xlsx | The agent's computed values match gold within tolerance |

If the agent produces no structured output (no JSON, no extractable data), correctness scores **1** with a clear "no agent data" message. The grader does not fall back to gold standard data — that would compare the gold standard to itself and always pass.

Similarly, if a canary appears only in the generated input files but not in the agent's output, the canary check **fails**. The agent must demonstrate it read the file by referencing the canary in its response.

**Self-test** (`--self-test`) answers: *Is the benchmark itself well-formed?*

Self-test validates the generated suite, not an agent. Each check searches **the generated input files and gold standards** instead of agent output:

| Check | What it searches | What it validates |
|---|---|---|
| Canary verification | Test case `input_files/` | Canary codes are actually embedded in the generated files |
| Error detection | Gold standard descriptions | Each planted error has a well-formed description (>5 chars) |
| Evidence (TC-19+) | Gold standard evidence specs | Each finding has non-empty `required_sources` and `acceptable_terms` |
| Correctness | Gold standard data | The gold defines non-empty expected values |
| Completeness | Gold standard structure | The gold defines `required_sheets`, `file_type`, and sections |

Self-test uses `SelfTestGrader`, a subclass that overrides the base `TestCaseGrader` methods. The base class never falls back to input files or gold data — that behavior is isolated in the subclass.

This separation is load-bearing. Before the fix (see `synth-data-bpt.1`, `synth-data-bpt.2`), the base grader could over-credit canaries by finding them in input files and over-credit correctness by falling back to gold data. Both paths now enforce strict boundaries.

### 2. Human rater (qualitative dimensions)

Two human raters independently score each TC on all 5 dimensions using `scoring/scoring_template.xlsx`. The template has four sheets:

| Sheet | Purpose |
|---|---|
| **Scorecard** | One row per test case, columns for each dimension + notes |
| **Grader Instructions** | Per-dimension per-TC guidance for anchoring the 1/2/3 judgments |
| **Inter-Rater Agreement** | Tracks both raters and computes Cohen's kappa |
| **Aggregate Dashboard** | Pivot tables: service line × capability × difficulty |

The rubric for each TC is auto-generated from `scoring/rubrics.yaml`, which declares the 3/2/1 anchors per dimension per TC with concrete value references to the gold standards.

Reconciliation: when two raters disagree by more than 1 point on any dimension, they discuss and produce a consensus score.

## Evidence expectations (legal/HR diligence pack)

Gold standards for TC-19 through TC-21 may include an `evidence_expectations` field. This declares which primary-source documents and key terms the agent must cite when arriving at a finding. Evidence expectations support scoring for professional-services reasoning that goes beyond numerical accuracy — the agent must demonstrate it relied on the correct sources and used calibrated language.

Example from a TC-19 gold standard:

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

Evidence expectations are optional and backward-compatible. Older gold standards without this field are scored normally. When present, the auto-grader checks that agent outputs reference the required sources and contain acceptable terms.

### The evidence dimension

When a gold standard defines `evidence_expectations`, the auto-grader scores an additional **evidence** dimension (1–3) alongside the standard five dimensions. This dimension is omitted for test cases without evidence expectations, so it does not affect TC-01 through TC-18 scoring.

The evidence dimension evaluates four concepts:

**Primary-source reliance.** High-risk findings must cite primary sources (contracts, agreements, amendments) rather than secondary summaries. For example, a change-of-control risk finding should reference the specific contract clause (e.g., `LCTR-001 §12.3`) and any relevant amendment (`AMD-002`), not just a management summary that paraphrases the clause. The auto-grader checks `required_sources` and enforces `primary_source_required` when set.

**Missing evidence.** When the data room is incomplete — a contract is unsigned, an executed copy is absent, a diligence request is unanswered — the agent must flag the gap rather than silently proceeding as if the evidence exists. Missing evidence is distinct from a wrong answer: the correct response is to note the gap and caveat any conclusions that depend on it.

**Assumptions and limitations.** Professional diligence memos include an explicit assumptions/limitations section. The agent should disclose what it assumed (e.g., that an unsigned draft reflects the final terms) and what it could not verify (e.g., whether a non-compete is enforceable in a specific jurisdiction). Findings that omit these qualifications score lower on evidence.

**Caveated conclusions.** Professional advisors use calibrated language: "appears to," "may indicate," "subject to further review." The agent must avoid definitive legal advice (e.g., "this contract is unenforceable") and instead present findings as signals requiring follow-up. The `acceptable_terms` list in each evidence expectation captures the key phrases the agent should use.

### Scoring mechanics

The auto-grader checks each finding in `evidence_expectations` independently:

1. **Source check**: Are all `required_sources` referenced in the agent output?
2. **Term check**: Does the agent output contain at least one of the `acceptable_terms`?

A finding passes if both checks pass. The overall evidence score:

| Score | Criteria |
|---|---|
| 3 | All findings pass both source and term checks |
| 2 | At least 50% of findings pass |
| 1 | Fewer than 50% of findings pass |

Per-TC evidence anchors in `rubrics.yaml` provide concrete guidance for human raters beyond what the auto-grader checks mechanically.

## Judgment traps (legal/HR diligence pack)

Judgment traps are professional-judgment challenges distinct from planted numerical errors. They test whether the agent can identify contradictions, missing evidence, scope boundaries, and overconfident conclusions.

Each trap is recorded in a judgment trap registry with a `JDG-NNN` ID. Unlike `ERR-NNN` planted errors (which have a definitive right answer), judgment traps have expected responses like `flag`, `caveat`, `deprioritize`, or `do_not_assert` — reflecting the calibrated language expected of a professional advisor.

### How judgment traps affect scoring

Judgment traps influence the standard five dimensions rather than creating a separate score:

- **Correctness**: Did the agent identify the trap? A missed contradiction (e.g., management summary says $5M but the contract says $3.6M) is a correctness failure.
- **Robustness**: Did the agent cross-reference sources to detect the trap, or did it accept the first source uncritically?
- **Communication**: Did the agent use appropriate language when reporting the trap? A `flag` trap expects the agent to raise the issue prominently; a `caveat` trap expects a qualified statement; a `do_not_assert` trap expects the agent to refrain from making a definitive claim.

The distinction between judgment traps and planted errors matters for grading calibration. A planted error (`ERR-NNN`) has one correct answer — the agent either catches the $9,000 mismatch or it doesn't. A judgment trap (`JDG-NNN`) tests the quality of professional reasoning — there is no single right number, but there is a right approach (cite sources, acknowledge uncertainty, use calibrated language).

See [`specs/universal-professional-services-scenario-packs.md`](../specs/universal-professional-services-scenario-packs.md) for the full judgment trap schema.

## The capability matrix

Spec §7.5 defines a 3D matrix for tracking results:

- **Axis 1 — Service Line:** Audit, Tax, Advisory, Legal, HR Diligence, Cross-Service
- **Axis 2 — Capability:** File Reading, File Writing, Data Analysis (SQL/DataFrame), RAG, Multi-Step Workflow
- **Axis 3 — Difficulty:** Routine, Complex, Adversarial

Each test case maps to one or more cells in this matrix. The aggregate dashboard pivots results so you can see things like "the agent is strong at routine file reading in Audit but weak at adversarial RAG in Tax" — which is actionable for prioritizing improvements.

## Regression testing

The generator is deterministic, so the same seed produces the same inputs every time. This makes regression testing cheap:

1. **Run the agent** on all test cases for the selected packs.
2. **Record** the scores in the scoring template.
3. **After an improvement cycle** (model change, prompt change, tool change), rerun all tests.
4. **Compare** score tuples to the prior run.
5. **Flag regressions** — any dimension that scored worse than before.
6. **Track trends** over time.

Because the inputs never change, any score change is attributable to the agent, not data drift. This is the single biggest reason the suite is worth the construction cost.

## Test execution protocol

Per spec §8, for each test case:

### Environment setup

1. Create a clean working directory.
2. Copy only the files listed in the test case's `input_files/` directory.
3. **Do not** provide any files from other test cases or the gold standards.
4. Provide the prompt exactly as written in the test case's `prompt.md`. No additional context.
5. Record the agent's full interaction log.

### Recording requirements

For each test run, capture:

- Start and end timestamps
- Full interaction log (every message, tool call, output)
- All files produced by the agent (copied to a results directory)
- Any errors or exceptions encountered
- Token count / cost if applicable

### Evaluation sequence

1. Run `auto_grader.py` on the agent outputs → JSON scores for mechanical dimensions.
2. Human Rater 1 reviews the interaction log and outputs → scores all 5 dimensions.
3. Human Rater 2 independently scores all 5 dimensions.
4. Reconcile any disagreements (>1 point on any dimension).
5. Record final scores in the scoring template.
6. Generate the capability matrix dashboard.

## Scoring the suite against the suite

`auto_grader --self-test` grades the **generated benchmark against itself** using `SelfTestGrader` (see [evidence semantics](#normal-grading-vs-self-test-evidence-semantics) above for how it differs from normal grading). Because the gold standards are derived from the same canonical model as the inputs, this should always produce `3/3/3/3/3 PASS` for every test case. When it doesn't, there's a shape mismatch somewhere in the pipeline:

| Symptom | Likely cause |
|---|---|
| `TC-NN: ?/?/?/1/?` with `error_detection/error_ERR-NNN: fail` | Formatter didn't plant the error at the location the gold standard expects, or the grader detection logic doesn't match the plant location |
| `TC-NN: ?/?/2/?/?` on Format | A file in the test case directory is corrupt, has the wrong extension, or can't be opened by the standard library |
| `TC-NN: 2/?/?/?/?` on Correctness | The gold standard and the input file computed different values from the same model — bug in either the formatter's view logic or the gold emitter |
| TC-NN missing entirely from grader output | Silent skip — usually a None-handling bug in the grader. (This has bitten us; see the `synth-data-06y`/`synth-data-ibr` fix history.) |

The self-test is the single most valuable regression check in the suite. Run it before every commit that touches the model, formatters, or gold standards.
