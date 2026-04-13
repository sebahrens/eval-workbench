# Architecture

The synth-data generator is a **three-phase deterministic pipeline** that turns a seeded configuration into a 204-file test corpus for evaluating AI agents on Big 4 workloads. This document explains how the phases fit together and why the design exists.

## The problem this architecture solves

The suite covers 21 test cases across Audit, Tax, Advisory, and Legal/HR Diligence service lines, and the cases are deliberately cross-referential: TC-01's trial balance must tie to TC-06's tax provision pre-tax book income, which must tie to TC-11's QofE EBITDA, which must tie to TC-15's DCF base. If these numbers ever drift, the suite is useless — an agent that correctly reconciles TC-01 would look like it's hallucinating when TC-06 disagrees.

Two naive approaches fail:

1. **Per-test-case owned generation.** Each test case owns generating its own input files. Simple, but guarantees drift the moment anyone touches one of them.
2. **Hand-maintained gold standards.** Gold files declare expected numbers and test cases must match. Works for one snapshot, breaks on the next regeneration.

The three-phase design eliminates both failure modes by construction.

## Phase 1 — Canonical model

`generator/model/` builds the entire Cascade Industries universe as in-memory Python dataclasses in a single seeded pass. This is the **single source of truth** for every number, name, date, and entity relationship in the entire suite.

Key modules:

| Module | Owns |
|---|---|
| `entities.py` | Parent + 3 subsidiaries with revenue targets, margins, locations |
| `coa.py` | ~120-account unified chart of accounts |
| `employees.py` | 850-employee roster (Faker-seeded, per-entity distribution, R&D eligibility, churn) |
| `gl.py` | Journal entry engine with balance-check enforcement |
| `revenue.py` | Monthly revenue by product × entity for FY2023–FY2025 with seasonal patterns |
| `opex.py`, `ppe.py` | SG&A, R&D, payroll; fixed assets with book vs tax depreciation |
| `leases.py` | 15 leases with terms, escalations, amendments; ASC 842 ROU / liability |
| `ar.py`, `ap.py` | Customer/vendor subledgers with aging buckets |
| `bank.py` | Bank transactions with outstanding checks and deposits in transit |
| `intercompany.py` | IC transactions recorded on both sides; nets to zero on consolidation |
| `tax.py` | Permanent/temporary differences, DTAs/DTLs, ETR reconciliation |
| `k1.py` | 8 partnership K-1 investments including amended |
| `rd.py` | 12 R&D projects, 2,340 time records, QRE computation |
| `ap_ledger.py` | 52,000 AP transactions with 7 planted anomaly categories |
| `customers.py` | Top customers, contracts, key people, pending litigation |
| `consolidation.py` | Eliminations, consolidated financials |
| `views.py` | Pure functions producing trial balance, balance sheet, income statement, cash flow, monthly P&L |
| `legal.py` | Legal contracts, clauses, amendments, diligence issues (M&A legal diligence pack) |
| `hr_diligence.py` | Employment agreements, retention awards, severance, contractor signals (HR diligence pack) |

Rules the model obeys:

- **Every number is posted through the GL engine.** The trial balance is a `GROUP BY account` on the GL. Nothing is hand-entered.
- **Intercompany eliminations happen by construction.** Each IC transaction is recorded on both sides, so consolidation sums to zero without special-case logic.
- **Derived views are pure.** `views.trial_balance()` takes a ledger and period, never mutates state, and returns the same result on every call.
- **Seeding is total.** `random`, `numpy.random`, and `Faker` are all seeded from `config.yaml`. Anything reachable from the model is byte-identical across runs.

## Phase 2 — Formatters (views into the model)

`generator/formatters/tcNN.py` modules are **pure functions that consume the canonical model and emit the specific files required by a test case.**

A formatter's job is presentation, not computation. It can:

- Project and aggregate model data into the TC's required shape
- Apply presentation chaos (merged xlsx cells, messy headers, abbreviations, inconsistent date formats, CSV column reordering)
- Embed a canary value at the designated location
- Apply planted errors from the error registry

A formatter **cannot**:

- Invent dollar amounts, dates, names, or counts
- Hardcode expected values that must also appear in the gold standard
- Read from one model module and declare contradicting values in another TC's output

Because formatters are pure views, changing a TC's file format (from xlsx to CSV, say) never changes the underlying numbers — they still trace back through the model.

## Phase 3 — Gold standards (derived from the model)

`generator/golds/framework.py` plus per-TC gold emitters produce `gold_standards/TC-NN_gold.json` files. Each gold is computed from the same canonical model as the formatter inputs, so they cannot drift from each other — they can only agree.

A gold standard declares:

- Required deliverable file types, sheets, sections
- Expected numerical values with tolerances
- Canary verifications (which canaries must appear in the agent's output to prove it read the right files)
- Error detection expectations (which `ERR-NNN` the agent is expected to flag)
- Communication and format acceptance hints

The gold standards are **not hand-maintained.** When the model changes, regenerating the suite automatically regenerates consistent golds.

## The error layer

`generator/errors.py` defines 9 deterministic transformation functions (one per error type in the spec §1.7):

`transpose_digits`, `mismatch_total`, `stale_data`, `formula_error`, `wrong_entity`, `date_inconsistency`, `classification_error`, `missing_data`, `rounding_discrepancy`.

Each is a pure function `(clean_value) → corrupted_value`. Planted errors are applied by formatters **on top of** clean formatter output, with every error recorded in `error_registry.json`:

```json
{
  "error_id": "ERR-001",
  "file": "cascade_tb_fy2025.xlsx",
  "location": "Sheet 'Trial Balance', Cell G47",
  "type": "transposed_digits",
  "clean_value": 18423109,
  "corrupted_value": 18432109,
  "severity": "material",
  "which_test_cases_should_catch": ["TC-01", "TC-02"]
}
```

Because errors are transformations, they are unit-testable: apply the error, un-apply, assert equal to the clean model output. `tests/test_errors.py` does exactly this for every planted error.

## The canary layer

`generator/canaries.py` embeds an 8-character alphanumeric code in every generated file at a format-specific location:

- **xlsx** — custom document property
- **docx** — custom document property
- **pdf** — document metadata subject field
- **csv** — leading comment line

`canary_registry.json` maps `file → canary → location`. The auto-grader uses this to verify the agent under test actually read the correct files — if an agent's TC-01 output contains the canary from `cascade_tb_fy2025.xlsx` but not the canary from `cascade_financials_fy2024_signed.pdf`, you know it skipped the PDF.

## Scenario packs

Test cases are organized into **scenario packs** — self-contained bundles registered in `generator/packs/`. Each pack declares its test case IDs, canary file keys, ordered emitters, and dependencies on other packs.

| Pack | Module | Test Cases | Depends On |
|---|---|---|---|
| `cascade_accounting_core` | `generator/packs/accounting_core.py` | TC-01 – TC-18 | — |
| `cascade_europe_ifrs` | `generator/packs/cascade_europe_ifrs.py` | TC-04-EU – TC-18-EU (9 cases) | `cascade_accounting_core` |
| `cascade_legal_hr_diligence` | `generator/packs/legal_hr_diligence.py` | TC-19 – TC-21 | `cascade_accounting_core` |

Packs are statically registered at import time. The orchestrator resolves selected packs, validates dependencies, and runs emitters in topological order.

Default generation runs `cascade_accounting_core` only. Packs beyond the default must be explicitly selected via the `pack_ids` parameter to `generate()`.

For the full pack contract and extension rules, see [`specs/universal-professional-services-scenario-packs.md`](../specs/universal-professional-services-scenario-packs.md).

## The orchestrator

`generate_test_suite.py` is the top-level entry point. It:

1. Parses `config.yaml` (seed, output dir, growth rates, margins, canary and error assignments)
2. Seeds all RNGs
3. Resolves selected scenario packs (validates dependencies, topological sort)
4. Builds the canonical model (Phase 1)
5. Invokes each pack's emitters (Phase 2) with the model, registering every emitted file in the manifest
6. Runs the gold standard emitters for active test cases (Phase 3)
7. Writes `manifest.json`, `canary_registry.json`, `error_registry.json`
8. Exits

The full pipeline runs in a few minutes and produces byte-identical output on every invocation.

## Why this design wins

- **Cross-referential drift is impossible.** Every number traces to the same model.
- **Gold standards are free.** They're derived, not maintained.
- **Adding a test case is local.** Write a new formatter and a new gold emitter; the model already has the data.
- **Determinism is cheap to enforce.** One seeded model build + pure formatters. `tests/test_determinism.py` double-runs the generator and `diff -r`s the outputs; anything nondeterministic shows up immediately.
- **Errors round-trip.** Planted errors are transformations, so they're unit-testable without needing the full suite to exist.

## What this design does not handle

- **PDF byte-level determinism** requires explicit handling. reportlab embeds timestamps and a producer string by default; fpdf2 embeds creation/modification dates. Both need to be pinned (`invariant=True` on reportlab Canvas, explicit `creation_date` and `modification_date` on fpdf2). See `docs/troubleshooting.md`.
- **Running the grader against an actual agent's output** is a separate concern from the self-test. The self-test grades the gold standards against themselves to ensure internal consistency; real grading compares agent output against a fresh suite + gold.
