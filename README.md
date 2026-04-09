# Cascade Industries Test Suite

A deterministic test suite for evaluating AI agents on Big 4 professional services workloads (Audit, Tax, Advisory). The suite generates 18 test cases built around a single fictional company — Cascade Industries, Inc. — ensuring cross-referential integrity across all service lines.

## Documentation

| Doc | For |
|---|---|
| [`docs/architecture.md`](./docs/architecture.md) | How the three-phase generator works and why it exists |
| [`docs/canonical-model.md`](./docs/canonical-model.md) | Reference for the Cascade Industries data model |
| [`docs/canaries-and-errors.md`](./docs/canaries-and-errors.md) | How provenance canaries and planted errors work |
| [`docs/scoring.md`](./docs/scoring.md) | Grading methodology, rubric, and test execution protocol |
| [`docs/adding-a-test-case.md`](./docs/adding-a-test-case.md) | How to extend the suite with TC-19+ |
| [`docs/troubleshooting.md`](./docs/troubleshooting.md) | Common failure modes and fixes |
| [`docs/glossary.md`](./docs/glossary.md) | Big 4 terminology for engineers |
| [`prompt.md`](./prompt.md) | Authoritative specification (the original design doc) |

## Prerequisites

- Python 3.11+
- [uv](https://docs.astral.sh/uv/) package manager

## Setup

```bash
git clone <repo-url> && cd synth-data
uv sync --all-extras
```

## Generate the Test Suite

```bash
uv run python generate_test_suite.py --output /tmp/test_suite
```

This produces the full suite under `/tmp/test_suite/`:

```
test_suite/
├── manifest.json            # Every generated file with metadata
├── canary_registry.json     # 8-char canary per file (for provenance checks)
├── error_registry.json      # 25 planted errors across test cases
├── shared_data/             # Cross-test reference files (COA, roster, financials, …)
├── test_cases/TC-01..TC-18/ # Per-test prompt, input files, expected behavior
├── gold_standards/          # Expected outputs (JSON + reference files)
├── scoring/                 # Rubrics, auto-grader, scoring template
└── templates/               # Word/cover templates
```

The generator is seeded (`SEED=42`) and produces byte-identical output on every run.

## Test Cases

| ID | Service Line | Difficulty | Description |
|----|-------------|------------|-------------|
| TC-01 | Audit | Complex | Trial balance reconciliation |
| TC-02 | Audit | Complex | Bank reconciliation & confirmation matching |
| TC-03 | Audit | Complex | Substantive analytical procedures on revenue |
| TC-04 | Audit | Adversarial | Lease extraction & ASC 842 schedule |
| TC-05 | Audit | Routine | AR workpaper memo |
| TC-06 | Tax | Complex | Tax provision under ASC 740 |
| TC-07 | Tax | Adversarial | K-1 extraction & consolidation |
| TC-08 | Tax | Complex | R&D tax credit study — Section 41 |
| TC-09 | Tax | Complex | Transfer pricing documentation |
| TC-10 | Tax | Routine | Multi-state apportionment |
| TC-11 | Advisory | Complex | Quality of earnings |
| TC-12 | Advisory | Adversarial | Data room triage & document index |
| TC-13 | Advisory | Complex | Forensic AP transaction analysis |
| TC-14 | Advisory | Routine | 13-week cash flow forecast |
| TC-15 | Advisory | Complex | DCF valuation |
| TC-16 | Cross-service | Routine | Engagement letter generation |
| TC-17 | Cross-service | Complex | Multi-file deliverable assembly |
| TC-18 | Cross-service | Adversarial | Prior year workpaper rollforward |

## Running Tests Against an Agent

### 1. Set up the test environment

For each test case:

1. Create a clean working directory.
2. Copy only the files from that test case's `input_files/` directory.
3. Do not provide files from other test cases or gold standards.
4. Deliver the prompt from the test case's `prompt.md` exactly as written.
5. Record the agent's full interaction (all tool calls, outputs, messages).

### 2. Record results

For each run, capture:

- Start and end timestamps
- Full agent interaction log (every message, tool call, output)
- All files produced by the agent
- Any errors or exceptions
- Token count / cost if applicable

### 3. Grade with the auto-grader

Grade a single test case:

```bash
uv run python -m scoring.auto_grader \
    --tc TC-01 \
    --gold /tmp/test_suite/gold_standards/TC-01_gold.json \
    --agent-output /path/to/agent/output/TC-01
```

Grade all test cases at once:

```bash
uv run python -m scoring.auto_grader \
    --suite-dir /tmp/test_suite \
    --agent-output-dir /path/to/agent/outputs
```

Self-test (gold standards must score 3/3/3/3/3):

```bash
uv run python -m scoring.auto_grader --self-test --suite-dir /tmp/test_suite
```

### 4. Human evaluation

After auto-grading:

1. Human Rater 1 reviews the interaction log and outputs, scores all 5 dimensions.
2. Human Rater 2 independently scores all 5 dimensions.
3. Reconcile disagreements (>1 point difference on any dimension).
4. Record final scores in the scoring template.

### 5. Scoring dimensions

Each test case is scored 1–3 on five dimensions:

| Dimension | What it measures |
|-----------|-----------------|
| Correctness | Numerical accuracy, proper calculations |
| Completeness | All required files, sheets, sections present |
| Format | File validity, structure, professional presentation |
| Robustness | Handling of edge cases, planted errors, adversarial inputs |
| Communication | Clarity of memos, notes, and explanations |

## Regression Testing

After each agent improvement cycle:

1. Re-run all 18 test cases (the fixed seed ensures identical inputs).
2. Compare scores to the prior run.
3. Flag any regressions (score decreased on any dimension).
4. Track improvement trends over time.

## Determinism Verification

Confirm the generator produces identical output:

```bash
uv run python generate_test_suite.py --output /tmp/run1
uv run python generate_test_suite.py --output /tmp/run2
diff -r /tmp/run1 /tmp/run2  # must produce no output
```

## Development

Run the test suite for the generator itself:

```bash
uv run python -m pytest tests/ -x
```

Lint:

```bash
uv run ruff check .
```

## License

Proprietary. All rights reserved.
