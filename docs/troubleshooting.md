# Troubleshooting

Common failure modes in the generator, the grader, and the Ralph build loop, with known fixes.

## Determinism failures

### Symptom

`tests/test_determinism.py` fails, or `loop.sh include-tests` reports:

```
⚠ Determinism FAILED — generator emits non-identical output across runs
```

The loop auto-files a P0 bead when this happens.

### Diagnosis

```bash
rm -rf /tmp/run1 /tmp/run2
uv run python generate_test_suite.py --output /tmp/run1
uv run python generate_test_suite.py --output /tmp/run2
diff -r /tmp/run1 /tmp/run2
```

The diff output identifies which files differ. Common patterns:

### Root causes and fixes

| Symptom in diff | Cause | Fix |
|---|---|---|
| PDF differs in every run, others don't | reportlab timestamp in document info | `canvas.setProducer(None)` + `canvas.setCreationDate(datetime(2025,1,1))` + pass `invariant=True` to `canvas.Canvas` |
| PDF differs in every run (fpdf2) | fpdf2 creation_date and modification_date | `pdf.set_creation_date(datetime(2025,1,1,0,0,0))` and `pdf.set_modification_date(...)` |
| xlsx differs in every run | openpyxl writes the modification time into `core.xml` | Set `wb.properties.modified = datetime(2025,1,1)` before saving |
| docx differs in every run | python-docx writes document properties with timestamps | Set `doc.core_properties.modified = datetime(...)` |
| Content order differs between runs | Iterating a `set` or an unordered `dict` | Sort explicitly: `for k in sorted(d.keys()):` |
| JSON file differs | `json.dump()` without `sort_keys=True` | Always use `json.dump(obj, f, sort_keys=True, indent=2)` |
| New random names each run | Faker not seeded | `Faker.seed(42)` at top of module plus `fake = Faker(); fake.seed_instance(42)` per instance |
| Numbers differ slightly each run | `numpy.random` without seeding | Use `rng = numpy.random.default_rng(42)` exclusively |

**Never paper over determinism with a retry loop or fuzzy comparison.** Fix the root cause.

## `bd sync` and embedded Dolt lock errors

### Symptom

```
Error: failed to open database: embeddeddolt: another process holds the exclusive lock
```

### Cause

bd 1.0 uses an embedded Dolt backend that supports only one writer at a time. If two `bd` commands run concurrently (e.g., a background loop plus a manual `bd list`), the second one fails.

### Fix

- Wait for the first command to finish: `pgrep -f "bd " | xargs ps -p` shows running processes
- In loop scripts, avoid parallelizing `bd dep add` calls — run them sequentially
- If a lock is orphaned (bd process crashed), delete `.beads/embeddeddolt/.lock` manually

## Grader output truncation

### Symptom

`loop.sh include-tests` Phase 2 output shows TC-05..TC-18 but TC-01..TC-04 appear to be missing.

### Cause

`loop.sh` pipes the grader output through `| tail -20`. When the grader reports 2+ TCs as FAIL, each failing TC adds expansion lines (one per failed subfield), which pushes the total output past 20 lines and the first few TCs are chopped off.

Manual runs without `tail` show all 18 TCs correctly, which misleads investigators into closing the bug as "not reproducible" — see the synth-data-06y history.

### Fix

In `loop.sh`, replace `| tail -20` with a larger limit (`| tail -60`) or drop the tail entirely. The grader's output is already bounded and structured; truncation serves no useful purpose.

## Self-test shows a TC failing on error detection

### Symptom

```
TC-08: 3/3/3/1/3 [FAIL]
    error_detection/error_ERR-005: fail —
    error_detection/error_ERR-019: fail —
```

### Cause

The formatter and the gold standard disagree about where the planted error lives, OR the grader's detection logic doesn't match what either of them did.

### Diagnosis

1. Load `error_registry.json` from the generated suite and find the entry for the failing error ID.
2. Open the file at that location and verify the corrupted value is actually present.
3. Open the gold standard JSON and find the `error_detection` section — does the expected flag match what the error registry says?
4. Check `scoring/auto_grader.py`'s `check_error_detection()` — does it look for the error at the right path?

### Fix

Reconcile the three sources. The error registry is the source of truth for *where* the error is, the formatter is responsible for *planting* it, the gold standard declares *what a good response looks like*, and the grader evaluates *whether the agent flagged it*.

## Agent closes a bug bead as "not reproducible" but the bug is still live

### Symptom

Loop iteration closes a bug bead with "not reproducible" reason. The next Phase 2 check in the same iteration reproduces the bug.

### Cause

The build agent ran a manual reproduction command that differed from what `loop.sh` actually runs. Often:

- Agent ran the grader against a cached `test_suite/` in the repo instead of against a fresh temp-dir generation
- Agent ran the grader without the `| tail -20` that `loop.sh` applies
- Agent used a different suite directory than `$RUN1`

### Fix

`PROMPT_build.md` now requires any "not reproducible" close to follow a specific reproduction protocol:

```bash
rm -rf /tmp/reprocheck
uv run python generate_test_suite.py --output /tmp/reprocheck
uv run python scoring/auto_grader.py --self-test --suite-dir /tmp/reprocheck
```

And paste the actual output into the close reason. If the agent violates this, reopen the bead with a note citing the contradictory Phase 2 output.

## Ralph loop hangs indefinitely

### Symptom

Claude subprocess finishes its work but the loop doesn't advance. Top shows the `claude -p` process at 0% CPU with stdout still attached.

### Cause

Known Claude Code bug (GitHub #19060, #25629, #31050): the process completes work but never calls `process.exit()`. The stdout pipe stays open, so the parent shell never sees EOF.

### Fix already in place

`loop.sh` uses `--output-format stream-json` and watches for a `{"type":"result"}` event. When it sees one, it gives the process 3 seconds to exit cleanly and then kills it. A 45-minute hard timeout is the fallback.

If you see a hang that lasts longer than the iteration's expected runtime, check:

```bash
ps aux | grep claude
```

If there's a stuck `claude` process, kill it manually:

```bash
pkill -f "claude.*--dangerously-skip-permissions"
```

The loop will create a fallback tracking bead for the failed iteration on the next pass.

## Loop iteration completes but nothing was committed

### Symptom

`git log` shows no new commits after a loop run that closed multiple beads.

### Cause

An earlier version of `PROMPT_build.md` didn't include a git commit step. Code changes accumulated on disk without being committed. Fixed in commit `6ddea77`.

### Check

Current `PROMPT_build.md` includes an explicit Step 5 with `bd close` + `git add -A` + `git commit -m "[<id>] ..."`. Verify with:

```bash
grep -A2 "git commit" PROMPT_build.md
```

### Recovery

If you have stranded work on disk from an old run:

```bash
git status   # inspect
git add -A
git commit -m "Recover stranded work from loop iterations NNN..MMM"
```

## Test case 01..04 missing from grader output

See the `tail -20` truncation entry above. This is the same bug.

## Model number changes after a refactor but a TC formatter still has the old value hardcoded

### Symptom

`auto_grader --self-test` reports Correctness < 3 for the affected TC. The model view returns value A, but the formatter emits value B, or the gold standard expects C.

### Diagnosis

```bash
grep -rn "18423109" generator/  # or whatever the old hardcoded number is
```

### Cause

A formatter or gold emitter hardcoded a number instead of pulling it from the model. This is explicitly forbidden in the formatter contract.

### Fix

Replace the literal with a model view call:

```python
# wrong
ar_balance = 18423109

# right
ar_balance = views.trial_balance(model, period="FY2025").account_balance("1200")
```

## Grader scores the right TC on the wrong file

### Symptom

An agent passes a test case that should fail, or vice versa.

### Cause

The grader's `--agent-output-dir` layout doesn't match what the grader expects. It looks for `agent_output/TC-01/`, not `agent_output/trial_balance.xlsx`.

### Fix

Make sure each test case's agent output lives in its own `TC-NN` subdirectory:

```
agent_output/
├── TC-01/
│   └── reconciliation_workpaper.xlsx
├── TC-02/
│   └── bank_recon.xlsx
└── ...
```

## When in doubt

1. Run `tests/test_determinism.py`. If it passes, the generator is healthy.
2. Run `auto_grader --self-test --suite-dir <fresh suite>`. If it passes 18×`3/3/3/3/3`, the internal consistency is intact.
3. Run `pytest tests/ -x`. If it passes, the model invariants hold.

All three green = the suite is in a known-good state and any failure you see is in code you just changed or in the agent under test, not in the suite itself.
