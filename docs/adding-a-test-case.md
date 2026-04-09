# Adding a Test Case

This is a how-to for extending the suite with a new test case (TC-19 and beyond). It assumes you've read [`architecture.md`](./architecture.md) and understand the three-phase design.

## When to add a test case

Add a test case when you want to measure a capability that the existing 18 don't cover. Good candidates:

- A service line or regulatory area not currently represented (e.g., consolidated SEC reporting, state sales tax, IFRS conversion)
- A specific adversarial pattern the existing cases don't exercise (e.g., contradictory data across multiple PDFs, OCR on genuinely low-quality scans, conflicting management representations across meetings)
- A new file format the suite doesn't currently generate (e.g., `.pptx`, `.msg` email, `.eml`, `.html`)

Don't add a test case just to have more coverage. Each TC costs a formatter, a gold emitter, and maintenance; thin variants of existing TCs add cost without proportional signal.

## The nine-step checklist

### 1. Sketch the TC on paper

Before writing any code, write the test case design as a plain Markdown section. Cover:

- **Service line and difficulty tier** (Audit/Tax/Advisory/Cross-Service, Routine/Complex/Adversarial)
- **Input files** and their formats
- **The prompt** as you'd write it for a real engagement manager delegating the work
- **Expected behavior** — what a competent agent should do
- **Gold standard values** — at least the top-line numbers; the rest will be derived
- **Planted errors** (if any) — which error types, which locations
- **Capability matrix mapping** — which cells in the service × capability × difficulty matrix does this exercise

Append the section to `prompt.md` at the end of the relevant service line block.

### 2. Extend the canonical model if needed

If the TC requires data the model doesn't already have (e.g., a new entity, a new ledger, a new schedule), **extend the model first**. Model changes are never TC-local:

- Find or create the appropriate module under `generator/model/`.
- Add dataclasses, generators, and query functions.
- Seed everything deterministically.
- Add unit tests under `tests/test_<module>.py` asserting invariants.
- Run `uv run pytest tests/ -x` and the determinism smoke test.

If the TC only consumes existing data (most do), skip this step.

### 3. Write the formatter

Create `generator/formatters/tc19.py`. Required structure:

```python
from pathlib import Path
from generator.model import views
from generator.canaries import embed_canary
from generator.errors import ErrorRegistry, transpose_digits

def build_tc19(model, output_dir: Path, canary_registry, error_registry: ErrorRegistry):
    """Emit TC-19 input files to output_dir/test_cases/TC-19/input_files/."""
    tc_dir = output_dir / "test_cases" / "TC-19"
    inputs_dir = tc_dir / "input_files"
    inputs_dir.mkdir(parents=True, exist_ok=True)

    # 1. Project model data into this TC's specific shape.
    trial_balance = views.trial_balance(model, period="FY2025")

    # 2. Emit the file with canary embedded.
    wb_path = inputs_dir / "cascade_something.xlsx"
    write_xlsx(wb_path, trial_balance)
    embed_canary(wb_path, canary_registry["cascade_something.xlsx"])

    # 3. Apply any planted errors.
    error_registry.apply(
        error_id="ERR-026",
        file=wb_path,
        location="Sheet 'Data', Cell B47",
        transformation=transpose_digits,
    )

    # 4. Write the prompt.md and expected_behavior.md (verbatim from prompt.md).
    (tc_dir / "prompt.md").write_text(TC19_PROMPT)
    (tc_dir / "expected_behavior.md").write_text(TC19_EXPECTED_BEHAVIOR)
```

Rules:

- **Never hardcode numbers that appear in the model.** Pull them from `views` or model modules.
- **Never invent numbers.** If the model doesn't have it, go back to step 2.
- **Always embed the canary.**
- **Plant errors through the registry**, not inline. This guarantees the error is recorded and verifiable.

Register the formatter in `generator/formatters/__init__.py` and wire it into `generate_test_suite.py`.

### 4. Write the gold standard emitter

Create `generator/golds/tc19.py` (or add to an existing module):

```python
from generator.model import views

def emit_tc19_gold(model) -> dict:
    """Build the TC-19 gold standard from the canonical model."""
    tb = views.trial_balance(model, period="FY2025")

    return {
        "test_case": "TC-19",
        "expected_outputs": {
            "file_type": "xlsx",
            "required_sheets": ["Summary", "Detail"],
            "summary": {
                "total_revenue": tb.account_balance("4000"),
                "total_expenses": tb.account_balance("5000") + tb.account_balance("6000"),
            },
        },
        "canary_verification": {
            "read_correct_input": "<canary from input file>",
        },
        "error_detection": {
            "ERR-026": {
                "location": "Trial Balance cell B47",
                "expected_flag": "transposed digit variance > material threshold",
            },
        },
    }
```

Register it in `generator/golds/framework.py` so the orchestrator emits it.

### 5. Write the rubric

Add a TC-19 block to `scoring/rubrics.yaml`:

```yaml
TC-19:
  dimensions:
    correctness:
      3: "All numbers match gold standard within ±0.5%"
      2: "Minor variances that don't change the conclusion"
      1: "Material errors"
    completeness:
      3: "Both Summary and Detail sheets present with all required columns"
      2: "One sheet present, the other missing or incomplete"
      1: "No deliverable or missing Summary"
    format:
      3: "Valid xlsx, professional formatting, no broken formulas"
      2: "Valid xlsx but sloppy formatting"
      1: "Broken file or wrong extension"
    robustness:
      3: "Caught ERR-026 and flagged the variance"
      2: "Noticed the variance but didn't investigate"
      1: "Missed the planted error entirely"
    communication:
      3: "Explained approach, flagged uncertainty on X, called out the error"
      2: "Adequate notes"
      1: "Silent or misleading"
  canary_inputs:
    - cascade_something.xlsx
  errors:
    - ERR-026
```

### 6. Update the auto-grader if needed

The auto-grader discovers test cases from `gold_standards/TC-*_gold.json`, so adding TC-19 should work automatically for structural checks. If TC-19 has novel grading logic (e.g., a new file format that the grader can't open), extend `scoring/auto_grader.py` with the new check.

### 7. Regenerate and self-test

```bash
uv run python generate_test_suite.py --output /tmp/tc19_check
uv run python -m scoring.auto_grader --self-test --suite-dir /tmp/tc19_check
```

The new TC-19 should appear in the output and score `3/3/3/3/3 PASS`. If it doesn't:

- **TC-19 missing from output:** gold emitter not registered in `framework.py`, or the grader isn't discovering the new gold file. Check the glob pattern and file naming.
- **Correctness fail:** the formatter and gold emitter are computing different values. Both must pull from the same model view.
- **Error detection fail:** the planted error isn't at the location the gold standard expects. Reconcile.

### 8. Update tests

Add a regression test at `tests/test_tc19_formatter.py`:

```python
def test_tc19_formatter_produces_required_files(suite_dir):
    tc_dir = suite_dir / "test_cases" / "TC-19"
    assert (tc_dir / "prompt.md").exists()
    assert (tc_dir / "input_files" / "cascade_something.xlsx").exists()

def test_tc19_gold_matches_formatter(suite_dir):
    gold = json.loads((suite_dir / "gold_standards" / "TC-19_gold.json").read_text())
    assert gold["expected_outputs"]["summary"]["total_revenue"] > 0
```

Run `uv run pytest tests/test_tc19_formatter.py -x`.

### 9. Update documentation

- Add the TC-19 row to the test case table in `README.md`
- If TC-19 introduces a new capability axis, update `scoring/rubrics.yaml` and the capability matrix in `docs/scoring.md`

## Things that usually go wrong

### Non-deterministic formatter

Symptoms: `tests/test_determinism.py` fails on the new TC.

Common causes:

- Iterating over a `set` or unordered `dict`
- Using `datetime.now()` for "current date" (use a fixed value from config instead)
- Calling Faker without seeding the per-instance state
- PDF generation without pinning timestamps (reportlab needs `invariant=True`, fpdf2 needs explicit `creation_date`)

Fix in the formatter; never in the determinism test.

### Gold standard contradicts the formatter

Symptoms: `auto_grader --self-test` shows the new TC with Correctness < 3.

Root cause: the formatter and the gold emitter derived the same number differently (e.g., one uses `sum()` over raw GL entries, the other uses `views.trial_balance().balance_for("4000")` which rounds). The fix is always "pull both from the same view function." Never reconcile numbers by editing the gold standard — that defeats the architecture.

### Canary not verifiable in output

Symptoms: `tests/test_canaries.py` fails on the new TC's input file.

Causes:

- Formatter forgot to call `embed_canary()` after writing the file
- `embed_canary()` was called but the file format doesn't support the embedding method (e.g., plain `.txt`, which isn't in the supported list)
- The canary registry doesn't have an entry for the new file — make sure `canary_generator` assigns one

### Planted error can't be round-tripped

Symptoms: `tests/test_errors.py` fails on ERR-026.

Causes:

- The error transformation isn't inverseable for the specific input (e.g., `transpose_digits` can produce the same output as a different clean input — pick a clean value where the transformation is uniquely reversible)
- The formatter applied the error before embedding the canary, and canary embedding re-wrote the file and clobbered the planted error

Apply errors after all other formatter operations, and unit-test the transformation function in isolation before using it in a formatter.

## Keep the spec in sync

`prompt.md` is the authoritative specification. Every new test case must be documented there first, not just in code. The build loop's prompt reads from `prompt.md` to understand what to implement — if the spec lags, build agents will either skip the new work or invent conflicting details.
