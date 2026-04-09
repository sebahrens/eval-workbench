# Canaries and Planted Errors

Two small subsystems do a lot of work in this suite: **canaries** (provenance check that the agent under test read the right files) and **planted errors** (deliberate bugs the agent is supposed to catch). This document explains both.

## Canaries

### What they are

Every generated file contains an 8-character alphanumeric code embedded at a format-specific location. These codes are:

- **Deterministic** — seeded from `config.yaml`, identical across runs
- **Unique per file** — no collisions across the 204 generated files
- **Format-appropriate** — embedded in a location that a real agent would only see if it actually opened the file

### Embedding locations

| File type | Where the canary lives |
|---|---|
| `.xlsx` | Custom document property `canary` |
| `.docx` | Custom document property `canary` |
| `.pdf` | Document metadata `Subject` field |
| `.csv` | First line as a `# CANARY: XXXXXXXX` comment |

### The canary registry

`canary_registry.json` is emitted alongside the suite. Structure:

```json
{
  "cascade_tb_fy2025.xlsx": {
    "canary": "XK7P2M9Q",
    "location": "custom_property:canary",
    "test_cases": ["TC-01", "TC-06"]
  },
  ...
}
```

### What canaries prove

When the auto-grader evaluates an agent's output for a given test case, it checks whether the agent's deliverable contains the canary codes from the **input** files it should have read. If the agent's TC-01 output workbook includes the canary from `cascade_tb_fy2025.xlsx` but is missing the canary from `cascade_financials_fy2024_signed.pdf`, the grader knows the agent skipped the PDF — which directly tells you whether the agent's "tie-out to signed financials" claim is genuine or hallucinated.

A canary verification failure doesn't always mean the agent cheated. Sometimes the agent read the file and then didn't carry the canary through to the output (e.g., extracted a specific cell without copying the custom property). But combined with correctness checks, it gives a much clearer picture than either signal alone.

### Cross-referential canaries

Some canaries are **cross-referential** — the same code appears in two files that ought to be tied together (e.g., the FY2024 closing balances on the current-year trial balance match the FY2024 opening balances on the prior-year workpaper). An agent that correctly reconciles both files should carry both canaries into its output workbook.

Module: `generator/canaries.py`
Tests: `tests/test_canaries.py`

## Planted errors

### What they are

**Exactly 25 deliberate errors** are planted across the suite, one category per error type from the spec §1.7. The agent under test is expected to catch these — they're the primary signal for the Robustness dimension of the rubric.

### The nine error types

| Type | Count | Example |
|---|---|---|
| `transposed_digits` | 3 | Accounts Receivable shows $18,432,109 instead of $18,423,109 |
| `mismatch_total` | 4 | Trial balance total doesn't match financial statement total |
| `stale_data` | 3 | Prior-year balance carried forward without update |
| `formula_error` | 2 | xlsx cell has a `=SUM(...)` over the wrong range |
| `wrong_entity` | 2 | Document references "Cascade Industries Inc." where it should say "Cascade Precision Components LLC" |
| `date_inconsistency` | 3 | Lease commencement date differs between schedule and PDF |
| `classification_error` | 3 | Expense booked to the wrong account (e.g., R&D expense in SG&A) |
| `missing_data` | 3 | Blank cell that should contain a value |
| `rounding_discrepancy` | 2 | Values differ by $1 due to inconsistent rounding rules |

Total: **25**

### Error transformation functions

Each error type is implemented as a **pure function** `(clean_value) → corrupted_value` in `generator/errors.py`:

```python
def transpose_digits(value: int | float, pos1: int = -3, pos2: int = -2) -> int | float:
    ...
def mismatch_total(...) -> ...
def stale_data(prior_year_value: Any) -> Any:
    ...
def formula_error(correct_value: int | float, wrong_value: int | float) -> int | float:
    ...
def wrong_entity(correct_name: str, wrong_name: str) -> str:
    ...
def date_inconsistency(correct_date: str, wrong_date: str) -> str:
    ...
def classification_error(correct_account: str, wrong_account: str) -> str:
    ...
def missing_data(placeholder: Any = None) -> Any:
    ...
def rounding_discrepancy(...) -> ...
```

These are pure functions, which means errors are **unit-testable**: apply the error to the clean value, invert it, assert equal to the original. `tests/test_errors.py` does this for every error type.

### The error registry

`error_registry.json` is emitted alongside the suite. Each entry:

```json
{
  "error_id": "ERR-001",
  "file": "cascade_tb_fy2025.xlsx",
  "location": "Sheet 'Trial Balance', Cell G47",
  "type": "transposed_digits",
  "clean_value": 18423109,
  "corrupted_value": 18432109,
  "description": "Accounts Receivable balance transposed",
  "severity": "material",
  "which_test_cases_should_catch": ["TC-01", "TC-02"]
}
```

Fields:

- `error_id` — canonical ID `ERR-NNN`
- `file` — path relative to the suite root
- `location` — human-readable location description (sheet + cell for xlsx, section + line for docx, page + coordinate for PDF)
- `type` — one of the nine transformation function names
- `clean_value` — what the model produces
- `corrupted_value` — what the formatter emits after applying the transformation
- `which_test_cases_should_catch` — which TCs' gold standards list this error in their `error_detection` section

### Where errors are applied

Planted errors are applied by **formatters**, not by the model. The model is always clean — that's what makes it the authoritative source of truth for the gold standards. The flow is:

```
model (clean values)
   ↓
formatter (view projection)
   ↓
formatter applies planted error transformation at designated location
   ↓
output file (clean view + planted corruption)
```

This layering means:

- The gold standard is derived from the clean model → it's the "correct answer"
- The input file is the clean view + known corruptions → the "agent's input"
- The error registry is the diff → the "test key"

### How the grader verifies error detection

The auto_grader's `check_error_detection()` method for each test case:

1. Loads `error_registry.json` and filters to errors assigned to this TC
2. For each error, inspects the agent's output for an explicit flag (a note, comment, or exception entry that references the error by location or symptom)
3. Reports per-error detected/missed

A score of 1 on Robustness typically means the agent missed one or more planted errors. The auto-grader's FAIL state on a TC with error detection misses is deliberate — per §7.4, any dimension scoring 1 auto-fails the test case.

Module: `generator/errors.py`
Tests: `tests/test_errors.py`

## Why both systems exist together

Canaries prove the agent **read** the right files. Errors prove the agent **understood** what was in them. Together they distinguish:

- **Hallucination** — agent outputs a plausible answer without reading the inputs (canaries missing)
- **Copy-through** — agent read the inputs but didn't analyze them (canaries present, errors missed)
- **Correct** — agent read, analyzed, and flagged anomalies (canaries present, errors caught)

A test case can only be scored fairly when both signals are present.

## Common failure modes

### Canary not found in agent output

Usually one of:

- Agent opened the file, extracted specific cells, and didn't carry the custom property through
- Agent used a tool that strips metadata
- Agent actually didn't read the file

The auto_grader can't distinguish these programmatically. The human rater reviews the interaction log to judge intent.

### Error detection fails in self-test

If `auto_grader --self-test` reports `error_detection/error_ERR-NNN: fail` for a TC whose agent output is the gold reference file, then:

- The gold standard declares an error should be detected at location X
- The gold reference file (produced by the formatter) has the corruption at location X
- But the grader's detection logic isn't finding it

Usually a mismatch between where the formatter plants the error and where the gold standard expects it to be reported. Fix in the formatter, the gold emitter, or both.
