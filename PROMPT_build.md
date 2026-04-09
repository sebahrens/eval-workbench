# Cascade Industries Test Suite — Build Mode

You are an AI agent implementing one task at a time for the **synth-data** project: a deterministic Python generator that produces a qualitative test suite for an AI agent operating on Big 4 professional services workloads (Audit, Tax, Advisory).

The full specification lives in `prompt.md` at the project root. It is the source of truth for every deliverable, test case, gold standard, and quality gate. Do not invent requirements that aren't in `prompt.md` or a bead.

## Project Structure (target)

```
synth-data/
├── prompt.md              # The specification — read before touching a task
├── AGENTS.md              # Beads landing-the-plane instructions
├── .beads/                # Beads issue tracker (source of truth for work)
├── config.yaml            # Generator configuration (company params, errors, canaries)
├── generate_test_suite.py # Deterministic generator (SEED=42)
├── manifest.json          # Emitted: every generated file + canary + test cases
├── canary_registry.json   # Emitted: 8-char canary per file
├── error_registry.json    # Emitted: 25 planted errors
├── shared_data/           # Cross-test reference files (COA, roster, org chart, financials, intercompany)
├── test_cases/TC-01..TC-18/  # Per-test prompt.md + input_files/ + expected_behavior.md
├── gold_standards/        # Expected outputs (JSON + reference xlsx/docx)
├── scoring/               # rubrics.yaml, auto_grader.py, scoring_template.xlsx
├── templates/             # Word/cover templates used by test cases
├── tests/                 # pytest unit + integration tests for the generator and grader
└── pyproject.toml         # uv-managed lockfile (Python 3.11+)
```

## Tech Stack (locked in bead `synth-data-1l0.26`)

- Python 3.11+
- openpyxl, python-docx, reportlab, fpdf2, pypdf
- pandas, numpy, faker
- pyyaml, pytest
- uv for dependency management

## Determinism Rules (non-negotiable)

The §9 quality gate requires `generate_test_suite.py` to produce byte-identical outputs across reruns. When writing any generator code:

1. Seed everything: `random.seed(42)`, `numpy.random.default_rng(42)`, `Faker.seed(42)`.
2. Never iterate an unordered `set`. Sort `dict` keys before writing to disk.
3. PDFs: pin `invariant=True` on reportlab Canvas and set a fixed creation date in document metadata. fpdf2 needs `creation_date` and `modification_date` set explicitly.
4. Avoid anything non-deterministic (`datetime.now()`, `uuid.uuid4()`, network calls).
5. Pin all dependency versions via the lockfile.

If you cannot satisfy determinism for something, file a bead rather than silently breaking the quality gate.

## Workflow

### 1. Find your task

```bash
bd list --status=in_progress
```

If any exist, resume the first one (`bd show <id>` for context). Otherwise:

```bash
bd ready
```

Pick the first ready task. If nothing is ready, run `bd list` to inspect blocked work and stop — do not invent tasks.

### 2. Claim and understand the task

```bash
bd update <id> --status=in_progress
bd show <id>
```

Read the bead description carefully. It references sections of `prompt.md`. Open those sections and read them before writing code.

### 3. Implement

- Stay strictly within the scope of the bead. Do not "improve" adjacent code.
- Follow existing patterns in the file you're modifying.
- Cross-referential integrity: if a number appears in multiple test cases, it must come from the master data model. Never hardcode a number that the spec ties across files.
- Canary values: every generated file must carry its 8-char canary from `canary_registry.json`.
- Planted errors: when a bead is responsible for a file that contains a planted error from `error_registry.json`, the error must match the registry exactly (location, type, value).

### 4. Verify

```bash
# Must pass
uv run python -m pytest tests/ -x

# Must pass — static checks
uv run ruff check .

# For any task that touches the generator: determinism smoke test
uv run python generate_test_suite.py --output /tmp/run1
uv run python generate_test_suite.py --output /tmp/run2
diff -r /tmp/run1 /tmp/run2  # must be empty
```

If you added a test case input file, also verify its canary is findable and any planted error is locatable.

If the generator or grader doesn't exist yet (you're implementing foundation work), run whatever subset of checks applies and note it in the close reason.

### 5. Complete the task

Close the bead and commit the work in a single logical commit. The bd 1.0 git hooks auto-sync `.beads/issues.jsonl` on commit, so you don't need a separate flush step.

```bash
bd close <id> --reason="Implemented: brief description of what was done"
git add -A
git commit -m "[<id>] <short description of change>"
```

**Never skip the commit.** If you close a bead without committing, the work is stranded on disk and the next iteration has no record of what changed. If the commit fails (e.g. pre-commit hook rejects it), investigate and fix the underlying issue — do not use `--no-verify`.

### Closing a bug bead as "not reproducible"

Before closing any bug bead with a "cannot reproduce" reason, you **must** run the exact same command path that `loop.sh` uses to reproduce it — not a manual variant against a cached suite in the repo. The `loop.sh` Phase 2 checks regenerate the suite into a fresh tempdir and run the grader against that. A bug that only appears under a fresh generation will look fine if you grade against `test_suite/` that's been sitting in the repo.

Minimum reproduction protocol:

```bash
rm -rf /tmp/reprocheck
uv run python generate_test_suite.py --output /tmp/reprocheck
uv run python scoring/auto_grader.py --self-test --suite-dir /tmp/reprocheck
```

Rules:

- If the bug is described in terms of grader output (e.g. "TC-XX missing," "ERR-NNN fails"), you must paste the actual output of the above commands into the close reason. Do not paraphrase.
- If the output matches what you expected (all TCs present, all errors detected), and the bug description claims otherwise, **do not close the bead**. Instead, update the description with the discrepancy you observed and leave it open for a human to reconcile. A false close on a self-test failure is worse than leaving the bead open — it hides a real regression behind an all-green loop.
- If you cannot reproduce the bug at all (clean output matching the close claim), still paste the output into the close reason so the next reader can verify.

## Rules

- **One task at a time.** Finish fully before starting another.
- **Stay in scope.** Only touch code relevant to the task.
- **Read `prompt.md` for ground truth.** Beads summarize; the spec is authoritative.
- **Respect determinism.** If your change could introduce nondeterminism, add a regression test that reruns the generator twice and diffs the outputs.
- **Don't skip the stack.** The stack is locked by bead `synth-data-1l0.26`. Don't introduce new libraries without updating that bead first.
- **Discover work.** If you find something new that's needed:
  ```bash
  bd create --title="Description" --type=bug --priority=2 --description="Details including the prompt.md section that motivates it"
  ```

## Now begin

Find the next task, implement it, verify it, close the bead, and **stop**. Do exactly ONE task per invocation.
