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

## The capability matrix

Spec §7.5 defines a 3D matrix for tracking results:

- **Axis 1 — Service Line:** Audit, Tax, Advisory, Cross-Service
- **Axis 2 — Capability:** File Reading, File Writing, Data Analysis (SQL/DataFrame), RAG, Multi-Step Workflow
- **Axis 3 — Difficulty:** Routine, Complex, Adversarial

Each test case maps to one or more cells in this matrix. The aggregate dashboard pivots results so you can see things like "the agent is strong at routine file reading in Audit but weak at adversarial RAG in Tax" — which is actionable for prioritizing improvements.

## Regression testing

The generator is deterministic, so the same seed produces the same inputs every time. This makes regression testing cheap:

1. **Run the agent** on all 18 test cases.
2. **Record** the scores in the scoring template.
3. **After an improvement cycle** (model change, prompt change, tool change), rerun all 18 tests.
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

`auto_grader --self-test` grades the **gold standards against themselves**. Because the gold standards are derived from the same canonical model as the inputs, this should always produce `3/3/3/3/3 PASS` for every test case. When it doesn't, there's a shape mismatch somewhere in the pipeline:

| Symptom | Likely cause |
|---|---|
| `TC-NN: ?/?/?/1/?` with `error_detection/error_ERR-NNN: fail` | Formatter didn't plant the error at the location the gold standard expects, or the grader detection logic doesn't match the plant location |
| `TC-NN: ?/?/2/?/?` on Format | A file in the test case directory is corrupt, has the wrong extension, or can't be opened by the standard library |
| `TC-NN: 2/?/?/?/?` on Correctness | The gold standard and the input file computed different values from the same model — bug in either the formatter's view logic or the gold emitter |
| TC-NN missing entirely from grader output | Silent skip — usually a None-handling bug in the grader. (This has bitten us; see the `synth-data-06y`/`synth-data-ibr` fix history.) |

The self-test is the single most valuable regression check in the suite. Run it before every commit that touches the model, formatters, or gold standards.
