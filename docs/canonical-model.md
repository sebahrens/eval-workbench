# The Canonical Model

The canonical model is Phase 1 of the generator — the in-memory representation of the entire Cascade Industries universe that every test case draws from. This document is a reference for what's in the model and how to query it.

For the **why** behind this design, see [`architecture.md`](./architecture.md).

## Cascade Industries, Inc. — the parent

```
Entity:        Cascade Industries, Inc.
Type:          US C-Corporation, mid-market manufacturer
Revenue:       ~$200M consolidated
Fiscal Year:   Calendar year (Jan 1 – Dec 31)
Years:         FY2023, FY2024, FY2025
HQ:            Portland, Oregon
```

FY2025 is the current year under audit/provision/analysis for every test case.

## The three wholly-owned subsidiaries

| Entity | Location | Revenue | Type | Characteristics |
|---|---|---|---|---|
| Cascade Precision Components LLC | Portland, OR | ~$95M | Industrial parts manufacturing | Stable, mature. Long-term leases, high fixed assets, low returns. Gross margin ~35%. |
| Cascade Advanced Materials, Inc. | Austin, TX | ~$65M | Specialty materials R&D + mfg | Growing fast. R&D ~12% of revenue. Government contracts. Gross margin ~52% but volatile. |
| Cascade Distribution Services LLC | Chicago, IL | ~$40M | Warehousing + logistics | Asset-light, high headcount, intercompany revenue. Gross margin ~18%. |

All three roll up to the parent. Consolidated revenue is approximately $200M with growth rates of 6% (FY23→24) and 9% (FY24→25).

Module: `generator/model/entities.py`

## Chart of accounts

Unified across all entities, ~120 accounts with 4-digit codes:

| Range | Type |
|---|---|
| 1xxx | Assets |
| 2xxx | Liabilities |
| 3xxx | Equity |
| 4xxx | Revenue |
| 5xxx | Cost of goods sold |
| 6xxx | Operating expenses (SG&A) |
| 7xxx | Other income/expense |
| 8xxx | Tax accounts |
| 9xxx | Intercompany |

Special subsets required by downstream test cases:

- **Granular R&D expense accounts** (TC-08 R&D credit study needs to isolate qualified research expenses)
- **Separate lease liability accounts** (TC-04 ASC 842 lease testing needs ROU asset and lease liability distinct from other long-term debt)

Module: `generator/model/coa.py`

## Employee roster

850 synthetic employees generated deterministically via Faker (seeded). Each row has:

| Field | Notes |
|---|---|
| `employee_id` | Sequential, prefixed by entity code (`PC-`, `AM-`, `DS-`, `CI-`) |
| `name` | Faker-seeded |
| `entity` | Which subsidiary |
| `department` | Engineering, Manufacturing, Sales, G&A, R&D, Warehouse, Finance |
| `title` | Realistic titles matched to department |
| `hire_date` | Distributed across 3 years with realistic churn |
| `annual_salary` | Realistic ranges by title and location |
| `state` | Work state (OR, TX, IL, plus remote in CA, WA, NY) |
| `cost_center` | 4-digit codes |
| `is_r_and_d_eligible` | Boolean, only true for Advanced Materials R&D and Engineering staff |
| `termination_date` | ~8% annual turnover; null for active employees |

**Guaranteed invariants** (asserted in `tests/test_employees.py`):

- Exactly 850 rows
- Exactly 45 R&D-eligible employees at Advanced Materials (consumed by TC-08)
- Termination rate ≈8% per year
- Byte-identical across reruns

Module: `generator/model/employees.py`

## General ledger

Every journal entry in the suite is posted through the GL engine:

```python
from generator.model.gl import Ledger, JournalEntry

ledger = Ledger()
ledger.post(JournalEntry(
    date="2025-03-15",
    entity="PC",
    description="March payroll accrual",
    lines=[
        ("6100", 120000, 0, "salaries"),
        ("2100", 0, 120000, "accrued payroll"),
    ],
))
```

`post()` validates that debits equal credits and raises on unbalanced entries.

Query API:

- `ledger.balance_by_account(entity, account, as_of_date)`
- `ledger.filter_by_date_range(start, end)`
- `ledger.filter_by_entity(entity)`
- `ledger.consolidated()` — returns a view of all entries across entities with intercompany eliminations applied

Module: `generator/model/gl.py`

## Revenue and COGS

`generator/model/revenue.py` generates monthly revenue by product × entity × month for all three fiscal years. Rules:

- **Seasonal patterns**, not uniform distribution. Q4 is heavy (manufacturing pull), Q1 dips.
- **Entity-specific gross margins:** Precision 35%, Advanced 52%, Distribution 18%. COGS is computed from revenue with these margins and posted automatically.
- **Growth rates:** FY23→FY24 +6%, FY24→FY25 +9%, with Advanced Materials driving most of the growth.
- **One product line at Advanced Materials declines 4%** (planted for TC-03 analytical procedures detection).
- **Consolidated FY25 YoY growth is 9.2%**, which is the "correct" answer TC-03 expects — management's representation of "approximately 8%" is a deliberate understatement.

## Operating expenses, fixed assets, depreciation

`generator/model/opex.py` and `generator/model/ppe.py`:

- SG&A by department
- R&D spend at approximately 12% of Advanced Materials revenue (TC-08 base)
- Payroll derived from the employee roster (no double-counting with the headcount data)
- Fixed asset register with acquisition dates
- **Book depreciation** (straight-line) and **tax depreciation** (MACRS) computed separately — the difference is the depreciation temporary difference feeding TC-06's tax provision

## Leases

15 leases with full term data:

- Lessee, lessor, commencement date, term in months
- Base rent, escalation type (fixed %, CPI, or stepped)
- Renewal options (term and rent)
- Purchase options
- Termination provisions

Special cases:

- **2 leases qualify for the short-term exemption** (≤12 months remaining with no purchase option reasonably certain) — TC-04 must flag these
- **3 leases have amendments** that materially change terms (rent, term, or escalation); the amended terms are authoritative, not the original

The model computes ROU asset and lease liability per ASC 842, and the difference between book and tax treatment is another temporary difference feeding TC-06.

Module: `generator/model/leases.py`

## AR and AP subledgers

`generator/model/ar.py` and `generator/model/ap.py`:

- AR subledger derived from revenue with DSO distribution producing realistic aging (current / 30 / 60 / 90 / 120+)
- Bad debt allowance and historical write-offs
- AP subledger with vendor-level balances by aging bucket
- Aging totals sum to the trial balance AR/AP balances (asserted in `tests/test_ar_ap.py`)

Consumed by TC-05 (AR workpaper memo), TC-11 (QofE), TC-14 (13-week cash flow), TC-13 (forensic AP).

## Bank transactions

340 bank transactions for December 2025 plus matching GL cash detail, with deliberately planted reconciling items:

- 4 outstanding checks (cleared in January)
- 2 deposits in transit
- Bank confirmation letter balance derivable from the bank statement ending balance + DIT − OS checks

TC-02 gold values:

- Adjusted bank balance and adjusted book balance both agree at **$4,287,331**
- Bank confirmation balance: **$4,312,117**

Fuzzy-match attributes: dates within ±2 business days, exact amounts.

Module: `generator/model/bank.py`

## Intercompany ledger

Per the spec §1.3, the IC relationships are:

| Relationship | Pricing |
|---|---|
| Precision → Advanced (raw materials) | Cost plus 8% |
| Distribution → Precision, Advanced (warehousing fees) | Market rate benchmarked to third-party contracts |
| Parent → all 3 subs (management fees) | 1.5% of sub revenue |
| Parent → Advanced (intercompany loan) | $5M at 5% interest |

Every IC transaction is recorded on both sides, so `consolidation.eliminate()` sums all 9xxx IC accounts to zero by construction.

**Deliberate outlier:** the services transaction (Distribution → others) is priced at an operating margin of **11.2%**, which falls outside the 4.2%–8.7% interquartile range of the comparable set used in TC-09. The agent under test should flag this.

Module: `generator/model/intercompany.py`

## Tax

`generator/model/tax.py` computes the full FY2025 tax provision per ASC 740:

- Pre-tax book income pulled from the consolidated income statement
- **Permanent differences:** M&E 50% disallowed, tax-exempt interest, stock comp excess, fines/penalties
- **Temporary differences:** depreciation (MACRS vs straight-line), ASC 842 lease adjustments, warranty reserve, inventory reserve, accrued bonuses, bad debt reserve
- Current tax: federal 21% + blended state 6.2%
- Deferred tax: rollforward of DTA/DTL from FY2024 balance sheet
- **Effective tax rate ≈ 24.8%** (higher than statutory due to state + permanent differences, partially offset by R&D credit)

TC-06 gold values are computed from this module.

## K-1 investments

8 partnership K-1s with box data (boxes 1-13 plus Box 20 codes). Amounts range $5K to $2.3M. Special cases:

- **1 amended K-1** (the agent must catch this): ordinary income changed from $340K to $285K and a $55K guaranteed payment was added
- **Section 199A amounts are present** on qualifying K-1s but flagged N/A because the filer is a C-corporation

Module: `generator/model/k1.py`

## R&D projects and time records

- **12 R&D projects** with descriptions, objectives, technical uncertainty, methodology:
  - 8 clearly qualify under the 4-part test (permitted purpose, technological in nature, technological uncertainty, process of experimentation)
  - 2 are borderline (should be flagged for manager review)
  - 2 clearly don't qualify (e.g., market research labeled as "R&D")
- **~2,340 time record rows** (45 R&D-eligible employees × 52 weeks with multiple project allocations per week)
- **Supply and materials expenses** coded to R&D cost centers
- **Qualified research expenses (QREs)** computed from time × W-2 wages + supplies
- **Prior year QREs:** FY2023 $3.1M, FY2024 $3.4M
- **Alternative Simplified Credit ≈ $185,000**

Module: `generator/model/rd.py`

## AP transaction ledger

52,000 AP transactions for FY2025 — the base for TC-13 forensic analysis. Every anomaly category in the spec is planted at known transaction IDs:

| Category | Count | Notes |
|---|---|---|
| Exact duplicate payments | 4 | Same vendor, amount, invoice |
| Near-duplicate payments | 4 | Same vendor, similar amount/invoice differing by 1 digit |
| Benford's law cluster | 35 | Transactions between $9,900–$9,999 (below $10K approval threshold) |
| Round-number shell vendor | 12 | Payments of $5K/$10K/$25K to "Pacific Consulting Group" |
| Weekend/holiday approvals | 15 | Approved outside business hours |
| Invoice-date-after-payment-date | 8 | Temporal anomaly |
| Similar-name vendors | 2 | "JKL Services LLC" / "JKL Services Inc" |
| Vendor at employee home address | 1 | Match to employee PC-0342 |
| Split transactions | 3 sets | Split to stay below approval thresholds |
| Self-approved transaction | 1 | Employee = approver |
| Single-approver cost center | 1 | 95% approval concentration |

**Total duplicate exposure:** $127,340 (TC-13 gold value).

Module: `generator/model/ap_ledger.py`

## Customers, contracts, and data room content

`generator/model/customers.py` seeds the TC-11 (QofE) and TC-12 (data room) test cases:

- **Top 10 customers** with revenue concentration. Top customer = 18% of revenue (TC-11 concentration risk flag).
- **8 major customer contracts** with terms, volumes, renewal dates. One (Acme) has a **change-of-control termination clause** planted for TC-12.
- **Key people:** CEO with 3× salary golden parachute, CFO, CTO employment agreements. Planted for TC-12 red flag detection.
- **Pending litigation:** product liability suit with $2.5M exposure (TC-12).

## Consolidation and derived views

Phase 1 ends with `generator/model/consolidation.py` (entity rollups + IC eliminations) and `generator/model/views.py` (pure-function derived artifacts):

- `trial_balance(period, entity="consolidated")`
- `balance_sheet(period)`
- `income_statement(period)`
- `cash_flow_indirect(period)`
- `monthly_pnl(period)`

These are the **inputs that Phase 2 formatters consume**. Nothing in the model below this layer is hand-entered; everything derives from GL journal entries posted in upstream modules.

## Adding a field to the model

1. Find the right module under `generator/model/`.
2. Add the field to the dataclass / dict / DataFrame as appropriate.
3. Ensure it's seeded deterministically — no `datetime.now()`, no unsorted set iteration, no unseeded randomness.
4. Add a unit test under `tests/test_<module>.py` asserting the field's shape and any invariants.
5. Run `uv run pytest tests/ -x` to confirm nothing broke.
6. Run `uv run python generate_test_suite.py --output /tmp/a && uv run python generate_test_suite.py --output /tmp/b && diff -r /tmp/a /tmp/b` to confirm determinism.

If the field is consumed by a test case, also update the relevant formatter in `generator/formatters/tcNN.py` and the corresponding gold standard emitter.
