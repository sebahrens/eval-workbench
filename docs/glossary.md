# Glossary

Big 4 professional services terminology used throughout the test suite, annotated for engineers who didn't spend four years reviewing journal entries in a windowless room.

## Audit

### ASC 842 (Leases)

US accounting standard requiring lessees to recognize most leases on the balance sheet as a **right-of-use (ROU) asset** and a corresponding **lease liability**. Previously, most operating leases were off-balance-sheet. TC-04 tests the agent's ability to extract lease terms and compute ROU/liability per ASC 842.

**Key judgment call:** whether a renewal option is "reasonably certain" to be exercised — it changes the lease term and therefore the balance-sheet numbers. The adversarial version of TC-04 plants ambiguous renewal clauses.

### Bank confirmation

A letter from the client's bank, sent directly to the auditor, stating the account balance and any loans/guarantees as of the year-end date. Auditors use it to independently verify the cash balance on the financial statements. TC-02 includes a confirmation PDF that the agent must reconcile against the bank statement and GL.

### Canary

Not a real audit term — it's project-specific. See `docs/canaries-and-errors.md`.

### Materiality

A threshold below which a misstatement wouldn't affect the decisions of a reasonable user of the financial statements. The audit rubric treats planted errors as "material" if they exceed this threshold; otherwise "immaterial." Auditors are expected to catch and document material items, and may pass on immaterial ones.

### Outstanding check

A check written by the company but not yet cleared by the bank. Shows up in the GL cash balance but not in the bank statement balance. TC-02 plants 4 outstanding checks the agent must identify.

### Planted error

Not an audit term. See `docs/canaries-and-errors.md`.

### Representation letter (rep letter)

A letter signed by client management at the close of the audit in which they assert things like "all known liabilities are recorded" and "there have been no material subsequent events." TC-03 plants a deliberate misrepresentation ("revenue grew approximately 8%" when actual growth is 9.2%) in the rep letter for the agent to challenge.

### Substantive analytical procedures (SAP)

Audit procedures that test financial statement assertions by comparing recorded amounts to the auditor's expectations developed from other data. Example: "we expect Q4 revenue to be X based on Y historical relationship; is the recorded revenue consistent?" TC-03 is a SAP on revenue.

### Tie-out

Tracing a number on one document back to its source on another document. "Tie the trial balance back to the signed financial statements." If the numbers don't match, the agent must flag it.

### Trial balance (TB)

A list of every account in the general ledger with its debit and credit balance as of a point in time. Debits must equal credits. This is the audit workpaper starting point for most procedures. TC-01 is a trial balance reconciliation.

### Workpaper

Documentation of audit procedures performed, evidence obtained, and conclusions reached. Organized by audit area (e.g., "Cash" workpaper, "AR" workpaper). Each workpaper has a standardized format: objective, scope, procedures, findings, conclusion. TC-05 tests workpaper memo drafting.

## Tax

### ASC 740 (Income Taxes)

US accounting standard for income tax provision. Requires companies to compute:

- **Current tax expense** — what they owe this year based on taxable income
- **Deferred tax expense** — the future tax consequence of timing differences between book and tax

The sum is the **total tax provision** on the income statement. TC-06 tests this end-to-end.

### Apportionment

The rule that determines how much of a multi-state company's income each state can tax. Each state has its own formula — some use a single sales factor, some use a three-factor formula (sales + payroll + property), some (like Texas) tax gross receipts instead of income. TC-10 tests multi-state apportionment.

### Deferred tax asset (DTA) / deferred tax liability (DTL)

Future tax consequences of differences between the book basis and the tax basis of assets/liabilities. Example: a warranty reserve is booked as an expense when recorded but isn't deductible until paid, so it creates a DTA. The **deferred rollforward** shows how these balances change year over year.

### Effective tax rate (ETR)

Total tax provision divided by pre-tax book income. Typically higher than the statutory rate (21% federal in the US) because of state taxes and permanent differences, or lower because of credits. The **rate reconciliation** shows the walk from statutory to effective rate. TC-06 expects an ETR ≈ 24.8%.

### K-1

A tax form (Schedule K-1) that a partnership or S-corporation files to report each partner/shareholder's share of income, deductions, and credits. A corporation that owns partnership interests receives K-1s from those partnerships and must incorporate the data into its own return. TC-07 tests K-1 extraction.

### Permanent difference

A book/tax difference that will never reverse. Example: 50% of meals and entertainment expenses is disallowed for tax purposes forever. Permanent differences affect the effective tax rate but don't create DTAs/DTLs.

### QRE (Qualified Research Expenses)

The inputs to the R&D tax credit under Section 41 of the Internal Revenue Code. Consists of:

- **Wages** paid to employees performing qualified research
- **Supplies** used in qualified research
- **Contract research** (limited to 65%)

NOT qualified: general overhead, routine testing, market research, reverse engineering. TC-08 tests QRE computation.

### Section 41 (R&D tax credit)

A federal tax credit for companies performing qualified research. Requires activities to pass a **4-part test**:

1. **Permitted purpose** — creating new or improved function, performance, reliability, or quality
2. **Technological in nature** — relies on physical, biological, computer, or engineering sciences
3. **Technological uncertainty** — uncertainty about capability, method, or design
4. **Process of experimentation** — systematic evaluation of alternatives

The **Alternative Simplified Credit (ASC)** method computes the credit as 14% × (current year QREs − 50% × average of prior 3 years QREs). TC-08 uses ASC.

### Section 199A

The Qualified Business Income (QBI) deduction for pass-through entities. **Not applicable to C-corporations**, which is the gotcha in TC-07: the K-1s report 199A amounts but the C-corp filer can't claim the deduction.

### Temporary difference

A book/tax difference that will reverse in a future period. Example: depreciation is computed differently for book (straight-line) and tax (MACRS); the cumulative depreciation eventually equals the cost basis under both methods, but the timing differs. Temporary differences create DTAs/DTLs.

### Transfer pricing

The pricing of goods, services, loans, and intangibles between related entities (e.g., a parent and its subsidiary, or two subsidiaries of the same parent). Must be "arm's length" — what unrelated parties would charge each other. TC-09 tests transfer pricing benchmarking.

## Advisory / Financial Due Diligence

### Adjusted EBITDA

EBITDA (Earnings Before Interest, Taxes, Depreciation, Amortization) plus management-proposed adjustments for items considered one-time or non-recurring. The **QofE bridge** walks from reported EBITDA to adjusted EBITDA. Buyers pay multiples of adjusted EBITDA, so every $1 of adjustment can translate to 5–15× that in enterprise value — making it the most contested number in any deal.

TC-11 plants aggressive adjustments (run-rate for a customer won in Q4, recurring consulting fees labeled "non-recurring") that the agent should challenge.

### Benford's Law

A statistical observation that in many naturally-occurring datasets, the first digit of values is distributed as log10(1 + 1/d). In forensic accounting, deviations from Benford can flag manipulated data — e.g., a cluster of invoices just below an approval threshold. TC-13 plants a Benford anomaly cluster at $9,900–$9,999 (just below $10K approval).

### Change of control

A contract clause that gives the counterparty the right to terminate or renegotiate if the company is acquired. In M&A due diligence, change-of-control clauses are red flags because they can destroy deal value (a major customer leaving post-close). TC-12 plants a change-of-control clause in the Acme customer agreement.

### Data room

A secure document repository where a seller gathers everything a buyer needs for due diligence: financials, contracts, HR records, tax returns, IP, litigation, etc. Modern data rooms are cloud-based (e.g., Intralinks, Datasite). TC-12 simulates a 32-file data room.

### DCF (Discounted Cash Flow)

A valuation method that estimates the present value of a business as the sum of its projected future free cash flows, discounted at a weighted average cost of capital (WACC). The terminal value (beyond the explicit projection period) typically dominates, making it sensitive to WACC and terminal growth assumptions. TC-15 is a DCF.

### EBITDA

Earnings Before Interest, Taxes, Depreciation, and Amortization. A proxy for operating cash flow that normalizes for capital structure and tax strategy. Used as a valuation metric (EV/EBITDA multiples) because it's roughly comparable across companies with different debt levels and jurisdictions.

### Golden parachute

A clause in a senior executive's employment agreement guaranteeing large severance payments if the company is acquired (typically 2–3× salary + bonus + accelerated equity). Buyers flag these because they directly cost money post-close. TC-12 plants a 3× golden parachute in the CEO employment agreement.

### Quality of Earnings (QofE)

A buyside due diligence exercise that validates the seller's reported earnings and challenges management's proposed adjustments. The deliverable is usually a report section titled "Key Findings" that says, "here's the adjusted EBITDA we can defend, here's what we accepted, here's what we challenged, here are the risks." TC-11 tests QofE.

### Run-rate adjustment

An EBITDA adjustment that annualizes a partial-period event. Example: a new customer signed in Q4 is run-rated as if it had been in place all year. **Aggressive** when based on short history (one quarter). TC-11 plants an aggressive run-rate adjustment that the agent should haircut.

## Cross-service

### Engagement letter

The contract between an audit/tax/advisory firm and its client, defining the scope, fees, deliverables, and payment terms of the engagement. Every new engagement starts with one. TC-16 tests engagement letter generation from a template and fee schedule.

### Rollforward

The process of updating a prior-year workpaper or model with current-year data. "Roll forward the audit workpapers from FY2024 to FY2025." Adversarial because format changes and new items can trip up an agent that assumes the structure matches last year. TC-18 tests rollforward.

### SG&A (Selling, General and Administrative Expenses)

The catch-all line on the income statement for operating costs that aren't part of cost of goods sold: sales salaries, marketing, office rent, admin headcount, legal fees, etc. Common source of classification errors (e.g., R&D expense misclassified to G&A).

### Workpaper memo

A written memo documenting audit procedures and conclusions for a specific audit area. Follows a template: Objective, Scope, Procedures, Findings, Conclusion. Language is stilted and conservative ("we obtained sufficient appropriate audit evidence..."). TC-05 tests workpaper memo drafting.

## Generic engineering vocabulary that shows up in the spec

### Merge fields

Placeholders in a Word template like `<<client_name>>` that get replaced with values when the document is generated from the template. TC-16's engagement letter template uses merge fields.

### OCR (Optical Character Recognition)

Converting images of text (scanned documents, photographs) into machine-readable text. Quality varies with scan resolution, font, and layout. The adversarial K-1 test case (TC-07) originally included OCR challenges but was simplified — now the adversarial signal is layout/format variation, not OCR quality.

### Trial balance reconciliation

Checking that two trial balances (e.g., prior year workpaper and current year client export) tie to each other or to supporting documentation. Variances are investigated and documented. TC-01 is a TB reconciliation.
