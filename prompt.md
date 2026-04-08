# Agent Test Suite: Big 4 Professional Services Workload Simulator

## Instructions for the Producing Agent

You are tasked with building a complete qualitative test suite for an AI agent ("the Agent Under Test") that operates on a user's local machine with access to folders and tools for reading/writing office files, PDFs, text files, performing data analysis (SQL, DataFrames), and RAG over document collections. The test suite must simulate realistic day-to-day work across the three core Big 4 service lines: Audit, Tax, and Advisory.

This document is your complete specification. Follow it section by section. Do not skip steps. Every deliverable must be machine-reproducible from the master data model.

---

## 1. MASTER DATA MODEL — "CASCADE INDUSTRIES"

### 1.1 Company Profile

Create a single fictional company ecosystem that feeds every test case across all three service lines. This ensures cross-referential integrity (the trial balance in Audit ties to the provision in Tax ties to the QofE in Advisory).

Generate the following master entity:

```
Company:        Cascade Industries, Inc.
Type:           US C-Corporation, mid-market manufacturer
Revenue:        ~$200M consolidated
Fiscal Year:    Calendar year (Jan 1 – Dec 31)
Years:          Generate 3 years of history: FY2023, FY2024, FY2025
                FY2025 is the "current year" under audit/provision/analysis
Headquarters:   Portland, Oregon
```

### 1.2 Subsidiary Structure

Generate three wholly-owned subsidiaries, each with distinct financial characteristics:

```
1. Cascade Precision Components LLC
   - Location: Portland, OR
   - Revenue: ~$95M
   - Type: Core manufacturing (industrial parts)
   - Characteristics: Stable, mature business. Several long-term leases.
     High fixed assets. Low return rate.

2. Cascade Advanced Materials, Inc.
   - Location: Austin, TX
   - Revenue: ~$65M
   - Type: Specialty materials R&D and manufacturing
   - Characteristics: High R&D spend (~12% of revenue). Growing rapidly.
     Several government contracts. Higher margin but volatile.

3. Cascade Distribution Services LLC
   - Location: Chicago, IL
   - Revenue: ~$40M
   - Type: Warehousing and logistics for parent + third parties
   - Characteristics: Asset-light, high employee count. Some intercompany
     revenue from parent entities. Lower margin.
```

### 1.3 Intercompany Relationships

Define intercompany transactions that create transfer pricing and elimination complexity:

- Precision Components sells raw materials to Advanced Materials at cost-plus-8%
- Distribution Services charges warehousing fees to both entities (market rate benchmarked to third-party contracts)
- Management fees flow from parent to all three subs (1.5% of sub revenue)
- An intercompany loan: parent lent $5M to Advanced Materials at 5% interest

### 1.4 Employee Roster

Generate a synthetic employee database (CSV) with 850 employees across all entities:

| Field | Notes |
|-------|-------|
| employee_id | Sequential, prefixed by entity code (PC-, AM-, DS-, CI-) |
| name | Synthetic names (use a name generator with fixed seed) |
| entity | Which subsidiary |
| department | Engineering, Manufacturing, Sales, G&A, R&D, Warehouse, Finance |
| title | Realistic titles matching department |
| hire_date | Distributed across 3 years, with realistic churn |
| annual_salary | Realistic ranges by title and location |
| state | Work state (OR, TX, IL, plus some remote in CA, WA, NY) |
| cost_center | 4-digit codes |
| is_r&d_eligible | Boolean — only for Advanced Materials R&D and Engineering staff |
| termination_date | ~8% annual turnover, nulls for active employees |

### 1.5 Chart of Accounts

Create a unified chart of accounts (4-digit codes) that maps consistently across all entities:

```
1xxx - Assets
2xxx - Liabilities
3xxx - Equity
4xxx - Revenue
5xxx - Cost of Goods Sold
6xxx - Operating Expenses (SG&A)
7xxx - Other Income/Expense
8xxx - Tax accounts
9xxx - Intercompany
```

Generate ~120 accounts with realistic descriptions. Ensure the account structure supports all test cases (e.g., include granular R&D expense accounts for the R&D credit study, separate lease liability accounts for ASC 842 testing).

### 1.6 Canary Values

Embed unique, verifiable "canary" strings and numbers throughout the data to confirm the Agent Under Test actually read the correct file:

- Every file must contain a unique 8-character alphanumeric code in a comment, metadata field, or designated cell (e.g., `CANARY: XK7P2M9Q`)
- Create a master canary registry (JSON) mapping each canary to its file and location
- Some canaries should be cross-referential: a canary in the audit trial balance should match a canary in the tax provision workpaper, proving the agent connected the right files

### 1.7 Deliberate Errors

Plant exactly 25 deliberate errors across the full data set. Document each in a master error registry (JSON) with:

```json
{
  "error_id": "ERR-001",
  "file": "cascade_tb_fy2025.xlsx",
  "location": "Sheet 'Trial Balance', Cell G47",
  "type": "transposed_digits",
  "description": "Accounts Receivable balance shows $18,432,109 instead of $18,423,109",
  "which_test_cases_should_catch": ["TC-01", "TC-02"],
  "severity": "material"
}
```

Error types to include:
- Transposed digits (3 instances)
- Mismatched totals between files (4 instances — e.g., TB total doesn't match financial statement)
- Stale data carried from prior year without update (3 instances)
- Incorrect formulas in Excel (2 instances)
- Wrong entity name in a document (2 instances)
- Inconsistent dates (3 instances — e.g., lease commencement date differs between schedule and PDF)
- Classification errors (3 instances — expense in wrong account)
- Missing data that should be present (3 instances — blank cells that should have values)
- Rounding discrepancies (2 instances — values that should tie but differ by rounding)

---

## 2. SYNTHETIC DATA GENERATION

### 2.1 Generator Script Requirements

Write a Python script (`generate_test_suite.py`) that:

1. Uses a fixed random seed (`SEED = 42`) for full reproducibility
2. Reads configuration from a `config.yaml` file (company parameters, error injection points, canary assignments)
3. Outputs all files to a structured directory (see 2.3)
4. Generates a manifest file (`manifest.json`) listing every generated file with its path, type, size, canary code, and which test cases use it
5. Generates the error registry and canary registry as JSON
6. Can be re-run to regenerate the full corpus identically

### 2.2 File Generation Libraries

Use these Python libraries for file generation:

| File Type | Library | Notes |
|-----------|---------|-------|
| .xlsx | openpyxl | Include formatting, merged cells, multiple sheets, formulas |
| .csv | stdlib csv | Use for raw data exports, deliberately messy where specified |
| .docx | python-docx | Include headers, footers, styles, tables |
| .pdf | reportlab (native) + fpdf2 (scanned-style) | Mix of text-native and image-based PDFs |
| .txt | stdlib | For meeting notes, emails, plain-text memos |
| .sqlite | sqlite3 | For SQL analysis test cases |

### 2.3 Output Directory Structure

```
test_suite/
├── config.yaml
├── generate_test_suite.py
├── manifest.json
├── canary_registry.json
├── error_registry.json
├── gold_standards/
│   ├── TC-01_gold.json
│   ├── TC-01_gold_output.xlsx
│   ├── ...
│   └── TC-30_gold.json
├── scoring/
│   ├── rubrics.yaml
│   ├── auto_grader.py
│   └── scoring_template.xlsx
├── test_cases/
│   ├── TC-01/
│   │   ├── prompt.md
│   │   ├── input_files/
│   │   │   ├── cascade_tb_fy2025.xlsx
│   │   │   └── ...
│   │   └── expected_behavior.md
│   ├── TC-02/
│   │   └── ...
│   └── ...
├── shared_data/
│   ├── master_coa.xlsx
│   ├── employee_roster.csv
│   ├── entity_org_chart.pdf
│   ├── cascade_consolidated_financials_fy2023.xlsx
│   ├── cascade_consolidated_financials_fy2024.xlsx
│   ├── cascade_consolidated_financials_fy2025.xlsx
│   └── intercompany_transactions.xlsx
└── templates/
    ├── workpaper_memo_template.docx
    ├── engagement_letter_template.docx
    ├── deliverable_cover_page.docx
    └── formatting_guide.pdf
```

### 2.4 Financial Data Generation Rules

When generating financial data, follow these rules for realism:

1. **Revenue recognition**: Monthly revenue should follow seasonal patterns (Q4 heavy for manufacturing, Q1 dip). Do not use uniform distribution.
2. **Expense correlation**: COGS should correlate with revenue at realistic gross margins (Precision: 35%, Advanced: 52%, Distribution: 18%).
3. **Balance sheet**: Assets = Liabilities + Equity at all times. Generate balance sheets that actually balance.
4. **Cash flow**: Derive cash flow from the balance sheet changes (indirect method). Do not generate it independently.
5. **Intercompany**: All intercompany transactions must net to zero on consolidation.
6. **Growth rates**: FY2023→FY2024 +6%, FY2024→FY2025 +9% (Advanced Materials driving growth).
7. **Round to whole dollars** in financial statements, keep cents in transaction-level data.

---

## 3. TEST CASES — AUDIT SERVICE LINE

### TC-01: Trial Balance Reconciliation

**Difficulty**: Complex

**Input Files**:
- `cascade_tb_fy2025.xlsx`: Current year trial balance from client. Deliberately messy: merged cells in header row, inconsistent account name formatting (some have leading spaces, some use abbreviations), two accounts renamed from prior year without mapping notes. Contains 1 transposed digit error (ERR-001).
- `cascade_tb_fy2024_workpaper.xlsx`: Prior year audit workpaper with clean, standardized account names and groupings. Contains prior year balances and lead schedule mappings.
- `cascade_financials_fy2024_signed.pdf`: Signed prior year financial statements (text-native PDF, 8 pages). Includes balance sheet, income statement, cash flow, and selected notes.

**Prompt** (store in `prompt.md`):
```
You have received the client's FY2025 trial balance and the prior year audit workpaper.

1. Map each account in the FY2025 trial balance to the prior year chart of accounts.
   Flag any new accounts that don't have a prior year equivalent, and any prior
   year accounts that are missing from the current year.
2. Compute the year-over-year variance ($ and %) for each account.
3. Flag any account with a variance greater than 10% AND greater than $100,000.
4. Verify that the FY2024 closing balances in this year's TB match the prior year
   signed financial statements.
5. Export the completed reconciliation as an Excel workpaper with the following sheets:
   - "Mapping": account-by-account mapping with flags
   - "Variance Analysis": all accounts with YoY variance
   - "Exceptions": flagged items requiring follow-up
   - "Tie-Out": comparison of TB opening balances to signed financials
```

**Gold Standard** (`TC-01_gold.json`):
```json
{
  "test_case": "TC-01",
  "expected_outputs": {
    "file_type": "xlsx",
    "required_sheets": ["Mapping", "Variance Analysis", "Exceptions", "Tie-Out"],
    "mapping": {
      "total_accounts_mapped": 118,
      "new_accounts_flagged": 3,
      "missing_accounts_flagged": 1,
      "renamed_accounts_correctly_identified": 2
    },
    "variance_analysis": {
      "flagged_accounts_count": 7,
      "flagged_accounts": ["4110", "5220", "6310", "6450", "1340", "2110", "7010"],
      "largest_variance_account": "4110",
      "largest_variance_pct": 14.2
    },
    "tie_out": {
      "discrepancies_found": 1,
      "discrepancy_details": "ERR-001: Accounts Receivable balance mismatch of $9,000"
    }
  },
  "canary_verification": {
    "read_correct_tb": "XK7P2M9Q",
    "read_correct_prior_wp": "LM3N8R2T",
    "read_correct_pdf": "QW5E9Y1A"
  }
}
```

**Expected Behavior Notes** (store in `expected_behavior.md`):
- The agent must handle merged cells and clean the header row before processing
- The agent should recognize that "Accts Recv" in the client TB maps to "Accounts Receivable" in the prior year
- The agent should catch the transposed digit error (ERR-001) when tying out to signed financials
- The variance analysis should use absolute values for percentage calculations on accounts that cross zero

**Scoring Rubric** (5 dimensions, 1–3 scale):
- Correctness: Do the numbers match gold standard? (3 = exact match, 2 = minor rounding differences, 1 = material errors)
- Completeness: All 4 sheets present with all required content? (3 = all present, 2 = missing minor elements, 1 = missing sheet or major content)
- Format Compliance: Valid xlsx, opens without errors, reasonable formatting? (3 = professional quality, 2 = functional but ugly, 1 = broken file or unusable layout)
- Robustness: Handled messy data (merged cells, abbreviations, renamed accounts)? (3 = handled all gracefully, 2 = handled most, 1 = failed on edge cases)
- Communication: Did the agent explain its approach, flag uncertainties, call out the error? (3 = proactive and clear, 2 = adequate, 1 = silent or misleading)

---

### TC-02: Bank Reconciliation & Confirmation Matching

**Difficulty**: Complex

**Input Files**:
- `bank_statement_dec2025.csv`: Bank statement export with 340 transactions. Columns: Date, Description (bank's format — cryptic abbreviations), Amount, Running Balance. Contains 4 outstanding checks (cleared in January) and 2 deposits in transit.
- `cascade_gl_cash_dec2025.xlsx`: General ledger cash account detail for December. Columns: Date, Reference, Description (company's format), Debit, Credit, Balance. Contains the same transactions in different order and naming.
- `bank_confirmation_fy2025.pdf`: Bank confirmation letter (PDF) confirming the balance as of 12/31/2025. The confirmation balance deliberately differs from both the bank statement ending balance and the GL balance by the expected reconciling items.

**Prompt**:
```
Perform a bank reconciliation for Cascade Industries as of December 31, 2025.

1. Match transactions between the bank statement and the general ledger.
   Use fuzzy matching on dates (allow ±2 business days) and amounts (exact match).
2. Identify all outstanding checks (in GL but not on bank statement).
3. Identify all deposits in transit (in GL but not on bank statement).
4. Identify any bank charges or interest not recorded in the GL.
5. Prepare a standard bank reconciliation schedule showing:
   - Balance per bank statement
   - Add: deposits in transit
   - Less: outstanding checks
   - Adjusted bank balance
   - Balance per GL
   - Add/Less: adjustments needed
   - Adjusted book balance
6. Verify that the adjusted balances agree and tie to the bank confirmation letter.
7. Export as an Excel workpaper.
```

**Gold Standard**: Pre-compute the exact reconciliation. The adjusted bank balance and adjusted book balance should agree at $4,287,331. The confirmation letter shows $4,312,117 (which equals bank statement ending balance + deposits in transit − outstanding checks, verifying the confirmation is consistent). Generate the gold output xlsx file with the completed reconciliation.

---

### TC-03: Substantive Analytical Procedures — Revenue

**Difficulty**: Complex

**Input Files**:
- `revenue_by_product_monthly_fy2024_fy2025.xlsx`: 24 months of revenue by product line (6 product lines) with unit volumes and average selling prices.
- `industry_benchmark_report.pdf`: A 12-page synthetic industry report (text-native PDF) containing industry average growth rates, margin benchmarks, and market size data. Embed the relevant benchmarks on pages 4, 7, and 11 (forcing the agent to search, not just read page 1).
- `management_rep_letter.docx`: Management representation letter asserting "consolidated revenue grew approximately 8% year-over-year, driven primarily by Advanced Materials."

**Prompt**:
```
Perform substantive analytical procedures on Cascade Industries' revenue for FY2025.

1. Analyze revenue trends by product line — monthly and annual.
2. Compute year-over-year growth rates by product line and in aggregate.
3. Compare growth rates to the industry benchmarks in the provided report.
4. Assess whether management's representation of ~8% growth is supported by the data.
5. Identify any product lines with unusual patterns (seasonality shifts, trend breaks,
   or growth significantly above/below industry benchmarks).
6. Draft an analytical procedures memo documenting:
   - Scope and objective
   - Data sources used
   - Methodology
   - Findings (with supporting data tables)
   - Conclusion and any follow-up procedures recommended

Output the memo as a Word document and the supporting analysis as an Excel workbook.
```

**Gold Standard**: Consolidated revenue grew 9.2% (not 8% as management claimed — this is a deliberate discrepancy the agent should catch). Two product lines drove the growth; one product line declined 4%. The agent should note that management's "approximately 8%" is close but understates actual growth, and should consider whether this is a rounding issue or intentional understatement.

---

### TC-04: Lease Extraction & Schedule Population (ASC 842)

**Difficulty**: Adversarial

**Input Files**:
- `leases/` folder containing 15 PDF lease agreements:
  - 10 text-native PDFs with varying layouts (some single-column, some two-column, different fonts)
  - 3 "scanned-style" PDFs (images of text — generated using reportlab with deliberately lower resolution)
  - 2 with amendments attached as additional pages modifying original terms
- `lease_schedule_partial.xlsx`: Partially completed lease schedule with 8 of 15 leases populated. Some fields are filled, others blank. Contains the column structure the agent should follow.

**Prompt**:
```
The audit team needs to complete the ASC 842 lease schedule for Cascade Industries.

1. Extract key terms from each lease agreement in the leases/ folder:
   - Lessee and lessor names
   - Commencement date
   - Lease term (in months)
   - Monthly/annual base rent
   - Escalation terms (fixed %, CPI-based, or stepped)
   - Renewal options (term and rent)
   - Purchase options
   - Termination provisions
2. For leases with amendments, use the amended terms (not the original).
3. Populate the lease schedule, matching the format of the existing entries.
4. Flag any leases that qualify for the short-term lease exemption (≤12 months
   remaining with no purchase option reasonably certain to be exercised).
5. Flag any leases where extracted terms are uncertain (e.g., from scanned documents
   where OCR may be unreliable).
6. Note any leases that may require judgment calls (e.g., whether a renewal is
   "reasonably certain" to be exercised).

Export the completed schedule as an Excel file.
```

**Gold Standard**: Pre-define all 15 leases with known terms. 2 leases qualify for short-term exemption. 3 leases have amendments that change material terms. The scanned PDFs should be readable but with deliberate OCR challenges (a "$" that could be read as "S", a "1" that could be read as "l"). The agent should flag these uncertainties rather than silently inserting wrong values.

---

### TC-05: Audit Workpaper Memo — Accounts Receivable

**Difficulty**: Routine

**Input Files**:
- `ar_aging_fy2025.xlsx`: Accounts receivable aging schedule (current, 30-day, 60-day, 90-day, 120+ day buckets) by customer.
- `ar_confirmations_summary.xlsx`: Summary of AR confirmation results (sent, received, agreed, exceptions).
- `allowance_analysis.xlsx`: Historical bad debt write-offs and allowance calculations.
- `workpaper_memo_template.docx`: Firm template with standard sections (Objective, Scope, Procedures, Findings, Conclusion).

**Prompt**:
```
Draft the accounts receivable workpaper memo for the FY2025 audit of Cascade Industries.

Use the firm template provided. The memo should document:
- The audit objective for AR
- Procedures performed (confirmation, aging analysis, subsequent receipts,
  allowance assessment)
- Key findings from the aging, confirmation results, and allowance analysis
- Your conclusion on whether the AR balance is fairly stated

Reference specific data from the provided workpapers. Keep the tone professional
and consistent with Big 4 audit documentation standards.

Save as a Word document.
```

**Gold Standard**: This is primarily a qualitative grading exercise. The gold standard is a reference memo written to Big 4 standards. Grade on: correct use of template, accurate data references, appropriate audit language ("we obtained sufficient appropriate audit evidence..."), logical flow from procedures to findings to conclusion, and identification of any concentration risk or aged receivables requiring attention.

---

## 4. TEST CASES — TAX SERVICE LINE

### TC-06: Tax Provision (ASC 740)

**Difficulty**: Complex

**Input Files**:
- `cascade_consolidated_tb_fy2025.xlsx`: Consolidated trial balance with pre-tax book income clearly identifiable.
- `tax_provision_fy2024_workpaper.xlsx`: Prior year provision workpaper with:
  - Current and deferred provision calculation
  - Deferred tax asset/liability rollforward
  - Effective tax rate reconciliation
  - List of permanent and temporary differences with descriptions
- `permanent_temporary_differences_fy2025.docx`: Table listing FY2025 book-tax differences:
  - Permanent: meals & entertainment (50% disallowed), tax-exempt interest, stock comp excess, fines/penalties
  - Temporary: depreciation (MACRS vs. straight-line), ASC 842 lease adjustments, warranty reserve, inventory reserve, accrued bonuses, bad debt reserve
  - Include dollar amounts for each item
- `statutory_rates.docx`: Federal rate 21%, blended state rate 6.2%, provide apportionment-weighted calculation.

**Prompt**:
```
Compute the income tax provision for Cascade Industries for FY2025 under ASC 740.

1. Calculate pre-tax book income from the trial balance.
2. Compute taxable income by applying all permanent and temporary differences.
3. Calculate the current tax provision (federal and state).
4. Calculate the deferred tax provision by computing the change in deferred tax
   assets and liabilities from prior year.
5. Roll forward the deferred tax balance sheet from FY2024 to FY2025.
6. Prepare the effective tax rate reconciliation (statutory rate to effective rate).
7. Identify the total provision (current + deferred) and the effective tax rate.

Export as an Excel workbook with separate sheets for:
- Current Provision
- Deferred Rollforward
- Rate Reconciliation
- Summary

Verify that everything ties: current + deferred = total provision, and the
rate reconciliation explains the difference between statutory and effective rates.
```

**Gold Standard**: Pre-compute all values. Effective tax rate should be approximately 24.8% (higher than 21% federal due to state taxes and permanent differences, partially offset by R&D credit). The deferred rollforward should balance. Include specific expected values for every line item.

---

### TC-07: K-1 Extraction & Consolidation

**Difficulty**: Adversarial

**Input Files**:
- `k1s/` folder containing 8 K-1 PDFs:
  - 3 system-generated (clean, structured layout)
  - 3 from different partnerships with different PDF generators (varying layouts, fonts, spacing)
  - 2 with handwritten annotations (simulated by overlaying text at angles, or added "handwritten-style" font notes)
  - Amounts range from $5,000 to $2.3M
  - Include boxes 1-13 plus various codes in Box 20
  - One K-1 has a corrected/amended indicator
- `entity_org_chart.pdf`: Shows which entities these K-1s flow through to the parent return.

**Prompt**:
```
Cascade Industries received K-1s from 8 partnership investments.

1. Extract all income, deduction, and credit items from each K-1.
2. Organize by K-1 box number and category (ordinary income, rental income,
   guaranteed payments, interest, dividends, capital gains, Section 179,
   charitable contributions, etc.).
3. Identify the amended K-1 and note what changed from the original.
4. Consolidate all K-1 data into a single summary schedule.
5. Map each item to the appropriate line on the corporate return (Form 1120).
6. Flag any items requiring special handling (e.g., passive activity limitations,
   at-risk limitations, Section 199A deductions flowing to a C-corp).

Export the consolidated schedule as an Excel file.
```

**Gold Standard**: Pre-define all K-1 values. The consolidated totals should be exact. The amended K-1 changed ordinary income from $340,000 to $285,000 and added a $55,000 guaranteed payment. The agent should catch that the Section 199A deduction is not applicable to a C-corporation filer.

---

### TC-08: R&D Tax Credit Study (Section 41)

**Difficulty**: Complex

**Input Files**:
- `rd_employee_time_records.csv`: 12 months of weekly time records for 45 R&D-eligible employees at Advanced Materials. Columns: employee_id, week_ending, project_code, hours, activity_description. ~2,340 rows.
- `rd_project_descriptions/` folder: 12 .docx files, one per active R&D project. Each contains: project name, objective, technical uncertainty addressed, methodology, and status. Some projects clearly qualify; some are borderline (e.g., routine quality testing that doesn't meet the 4-part test); some clearly don't qualify (market research labeled as "R&D").
- `payroll_data_fy2025.xlsx`: Full payroll register for Advanced Materials with W-2 wages, employer taxes, and benefits by employee.
- `rd_supply_expenses.xlsx`: Supply and materials expenses coded to R&D cost centers.

**Prompt**:
```
Perform an R&D tax credit study for Cascade Advanced Materials, Inc. for FY2025.

1. Review each project description and determine whether it meets the four-part
   test for qualified research:
   - Permitted purpose (new or improved function, performance, reliability, quality)
   - Technological in nature (relies on principles of physical/biological/computer science)
   - Technological uncertainty (capability, method, or design uncertainty)
   - Process of experimentation (systematic evaluation of alternatives)
2. For qualifying projects, compute qualified research expenses (QREs):
   - Wages: allocate based on time records (% of time on qualifying activities × W-2 wages)
   - Supplies: include supplies directly used in qualified research
   - Do not include overhead or general administrative expenses
3. Compute the credit using the Alternative Simplified Credit (ASC) method:
   - Average QREs for FY2023-FY2025 (provide FY2023 and FY2024 QREs as $3.1M and $3.4M)
   - Credit = 14% × (current year QREs − 50% × average of prior 3 years QREs)
4. Draft a contemporaneous documentation memo for each qualifying project that
   summarizes the technical uncertainty and experimentation.

Export:
- Project qualification analysis as Excel (project name, qualification determination, rationale)
- QRE computation as Excel
- Credit calculation as Excel
- Documentation memos as a single Word document with one section per project
```

**Gold Standard**: Pre-define which projects qualify (8 of 12 qualify, 2 are borderline that should be flagged for manager review, 2 clearly don't qualify). The computed credit should be approximately $185,000. Include specific QRE amounts per project.

---

### TC-09: Transfer Pricing Documentation

**Difficulty**: Complex

**Input Files**:
- `intercompany_transactions_fy2025.xlsx`: All intercompany transactions with entity, counterparty, type (goods, services, interest), volume, and pricing.
- `comparable_companies.xlsx`: Financial data for 12 comparable companies (revenue, COGS, operating expenses, operating income, total assets) for benchmarking. Include 2 companies that should be rejected as comparables (one in a different SIC code, one with extreme financial distress).
- `tp_report_fy2024.pdf`: Prior year transfer pricing report (42 pages). Key sections: functional analysis (pp. 8-15), economic analysis (pp. 22-30), benchmarking results (pp. 31-38). The rest is boilerplate.

**Prompt**:
```
Update the transfer pricing analysis for Cascade Industries for FY2025.

1. Using the intercompany transaction data, calculate the actual intercompany
   margins for each transaction type (goods, services, interest).
2. Screen the comparable companies:
   - Reject any that are not appropriate comparables (explain why)
   - Compute the interquartile range of operating margins for the accepted set
3. Determine whether Cascade's intercompany margins fall within the arm's-length
   range (interquartile range of comparables).
4. Flag any transaction types that fall outside the range.
5. Draft the "Economic Analysis — Results" section of the local file, updating
   the prior year report's language with current year data and conclusions.

Export:
- Benchmarking analysis as Excel
- Updated results section as Word document
```

**Gold Standard**: Two comparable companies should be rejected. The interquartile range should be 4.2% to 8.7% operating margin. The goods transfer (cost-plus-8%) falls within range. The services transfer is at 11.2%, outside the range — the agent should flag this. Pre-compute all benchmarking statistics.

---

### TC-10: Multi-State Apportionment

**Difficulty**: Routine

**Input Files**:
- `consolidated_pl_fy2025.xlsx`: Consolidated income statement.
- `state_factors.xlsx`: State-by-state factors schedule with columns: State, Sales Factor, Payroll Factor, Property Factor. Populated for OR, TX, IL. Partially populated for CA, WA (sales only), NY (blank). Some cells contain "$0" vs. blank (the agent must distinguish between "zero presence" and "data not provided").
- `apportionment_rules.docx`: Reference table showing each state's apportionment formula:
  - OR: Single sales factor
  - TX: Margin tax, not income tax (different base)
  - IL: Single sales factor
  - CA: Single sales factor, market-based sourcing
  - WA: B&O tax (gross receipts), not income tax
  - NY: Single sales factor with customer-based sourcing

**Prompt**:
```
Complete the multi-state apportionment and tax analysis for Cascade Industries.

1. Complete the apportionment schedule for all states.
2. Flag states with incomplete factor data and note what's missing.
3. Apply each state's apportionment formula to compute state taxable income.
4. For states with non-standard tax bases (TX margin tax, WA B&O tax), note that
   a different computation is needed and explain what information is required.
5. Flag any states where nexus may be questionable based on the factor data.
6. Summarize estimated state tax liability by jurisdiction.

Export as an Excel workbook.
```

**Gold Standard**: Pre-compute apportioned income for each state. TX and WA should be flagged as requiring different calculations. CA and NY should be flagged for incomplete data. WA should be flagged for nexus question (only sales data, no physical presence). The agent should not attempt to compute TX margin tax or WA B&O tax without the proper base — it should explain what additional information is needed.

---

## 5. TEST CASES — ADVISORY SERVICE LINE

### TC-11: Quality of Earnings (Financial Due Diligence)

**Difficulty**: Complex

**Input Files**:
- `monthly_pl_fy2023_fy2024_fy2025.xlsx`: 36 months of P&L data by account. Include line-item detail (not just categories).
- `management_adjustments.xlsx`: Management's proposed EBITDA adjustments with descriptions:
  - Owner compensation above-market ($180K)
  - One-time legal settlement ($420K)
  - COVID-related PPP loan forgiveness ($250K in FY2023 — stale)
  - Non-recurring consulting fees ($95K)
  - "Run-rate" adjustment for new customer won in Q4 ($600K annualized — aggressive)
  - Facility relocation costs ($310K)
  - Two adjustments that are actually recurring and should be challenged
- `customer_contracts/` folder: 8 PDF contracts with top customers showing terms, volumes, pricing, and renewal dates.
- `management_interview_notes.docx`: Notes from management Q&A sessions covering business overview, growth drivers, customer concentration, key personnel, and pending litigation.

**Prompt**:
```
Perform a quality of earnings analysis for a potential acquisition of Cascade Industries.

1. Compute reported EBITDA for each of the 36 months and annually.
2. Evaluate each of management's proposed adjustments:
   - Is it truly non-recurring?
   - Is the amount supportable?
   - Is it properly categorized (above/below the line)?
   - Challenge any adjustments that appear aggressive or recurring in nature.
3. Compute adjusted EBITDA after your accepted/modified adjustments.
4. Create a QofE bridge: Reported EBITDA → each adjustment → Adjusted EBITDA.
5. Analyze revenue quality:
   - Customer concentration (top 10 customers as % of revenue)
   - Contract renewal risk (any contracts expiring within 12 months)
   - Revenue trend sustainability
6. Draft the "Key Findings" section of the QofE report covering:
   - Adjusted EBITDA conclusion
   - Material adjustments and rationale
   - Revenue quality and risks
   - Items requiring further diligence

Export:
- Analysis workbook (Excel) with EBITDA bridge, monthly detail, customer analysis
- Key Findings memo (Word document)
```

**Gold Standard**: Reported EBITDA is $28.4M. Management's adjustments total $1.855M. The agent should challenge the "run-rate" customer adjustment (aggressive — only 1 quarter of history) and one of the "non-recurring" consulting fees (similar charge in 2 of 3 years). Properly adjusted EBITDA should be approximately $29.3M (accepting most adjustments, haircut on run-rate, rejecting the recurring one). Top customer is 18% of revenue — flag concentration risk.

---

### TC-12: Data Room Triage & Document Index

**Difficulty**: Adversarial

**Input Files**:
Create a `data_room/` folder containing 32 files simulating a deal data room:

```
data_room/
├── 01_corporate/
│   ├── articles_of_incorporation.pdf
│   ├── bylaws.pdf
│   ├── board_minutes_2024.pdf
│   ├── board_minutes_2025.pdf
│   └── org_chart.pdf
├── 02_financial/
│   ├── audited_financials_fy2023.pdf
│   ├── audited_financials_fy2024.pdf
│   ├── management_financials_fy2025_ytd.xlsx
│   ├── budget_fy2025.xlsx
│   └── debt_schedule.xlsx
├── 03_legal/
│   ├── material_contracts/
│   │   ├── customer_agreement_acme.pdf
│   │   ├── customer_agreement_globex.pdf
│   │   └── supplier_agreement_initech.pdf
│   ├── pending_litigation_summary.docx
│   ├── ip_assignment_agreements.pdf
│   └── insurance_policies_summary.pdf
├── 04_hr/
│   ├── employee_census.xlsx
│   ├── benefits_summary.pdf
│   ├── key_employee_agreements/
│   │   ├── ceo_employment_agreement.pdf
│   │   ├── cfo_employment_agreement.pdf
│   │   └── cto_employment_agreement.pdf
│   └── org_chart_detailed.xlsx
├── 05_tax/
│   ├── federal_returns_fy2023.pdf
│   ├── federal_returns_fy2024.pdf
│   ├── state_returns_summary.xlsx
│   └── tax_notices.pdf
├── 06_operations/
│   ├── facility_leases.pdf
│   ├── equipment_list.xlsx
│   ├── customer_list_with_revenue.xlsx
│   └── vendor_list.xlsx
└── 07_technology/
    ├── patent_portfolio.pdf
    ├── software_licenses.xlsx
    └── it_infrastructure_overview.docx
```

Also provide:
- `dd_checklist_standard.docx`: A standard due diligence checklist with 65 line items organized by category (Corporate, Financial, Tax, Legal, HR, Operations, Technology, Environmental, Regulatory). Not all items will have corresponding documents.

**Prompt**:
```
You have access to a deal data room for the potential acquisition of Cascade Industries.

1. Create a complete document index: for each file, provide:
   - File path
   - Document type/category
   - Date (if identifiable)
   - One-line summary of contents
   - Key data points or red flags noted
2. Cross-reference against the due diligence checklist:
   - Mark which checklist items have corresponding documents
   - Flag checklist items with NO corresponding document (gaps)
   - Note any documents in the data room not covered by the checklist
3. Identify red flags or items requiring immediate attention, such as:
   - Pending litigation
   - Contracts with change-of-control provisions
   - Key employee agreements with unusual terms
   - Missing critical documents
4. Prioritize the gaps: which missing documents are deal-critical vs. nice-to-have?

Export:
- Document index as Excel (with columns: Path, Category, Date, Summary, Red Flags)
- Gap analysis as Excel (Checklist Item, Status, Priority, Notes)
- Red flags summary as Word memo
```

**Gold Standard**: Pre-define the content of all 32 files with known red flags:
- Pending litigation: a product liability suit with $2.5M potential exposure
- CEO employment agreement has a change-of-control golden parachute of 3× salary
- Customer agreement with Acme has a change-of-control termination clause
- Missing from data room: environmental assessments, regulatory permits, insurance claim history, real property surveys
- Patent portfolio includes 2 patents expiring within 18 months
- The IP assignment agreements are incomplete (missing assignments from 2 founding employees)

The agent should identify all red flags and rank missing items by deal-criticality.

---

### TC-13: Forensic AP Transaction Analysis

**Difficulty**: Complex

**Input Files**:
- `ap_transactions_fy2025.csv`: 52,000 rows of accounts payable transactions with columns: transaction_id, date, vendor_id, vendor_name, amount, description, approver, cost_center, payment_method, invoice_number.

Embed the following synthetic anomalies (document exact locations in gold standard):
```
ANOMALY SET:
1. Duplicate payments (8 instances):
   - 4 exact duplicates (same vendor, amount, invoice)
   - 4 near-duplicates (same vendor, same amount, different invoice numbers
     that differ by 1 digit — suggesting a typo or intentional variation)

2. Benford's Law violations:
   - Cluster of 35 transactions between $9,900-$9,999 (just below $10K
     approval threshold)

3. Round number anomalies:
   - 12 transactions at exactly $5,000.00, $10,000.00, $25,000.00 to a
     single vendor ("Pacific Consulting Group" — a shell-company-style name)

4. Temporal anomalies:
   - 15 transactions approved on weekends or holidays
   - 8 transactions with invoice dates AFTER payment dates

5. Vendor anomalies:
   - 2 vendors with P.O. Box-only addresses sharing a similar name
     ("JKL Services LLC" and "JKL Services Inc.")
   - 1 vendor with an employee's home address (match to employee roster)

6. Split transactions:
   - 3 sets of transactions split to stay below approval thresholds
     (e.g., a $14,500 purchase split into $7,200 and $7,300 on consecutive days)

7. Approver anomalies:
   - 1 approver who approved their own reimbursement (employee = approver)
   - 1 cost center with a single approver for 95% of transactions (lack of segregation)
```

**Prompt**:
```
Perform a forensic analysis of the FY2025 accounts payable transaction ledger
for Cascade Industries.

1. Duplicate payment analysis:
   - Identify exact duplicates (same vendor, amount, invoice)
   - Identify near-duplicates (same vendor, similar amounts or invoice numbers)
   - Quantify the total dollar exposure from potential duplicates

2. Benford's Law analysis:
   - Test the first-digit and first-two-digit distributions
   - Flag any statistically significant deviations
   - Identify specific transaction clusters causing deviations

3. Approval threshold analysis:
   - Determine if there are clusters of transactions just below common
     approval thresholds
   - Identify potential split transactions

4. Temporal analysis:
   - Flag transactions approved outside business hours or on weekends/holidays
   - Identify transactions where the payment date precedes the invoice date

5. Vendor analysis:
   - Identify vendors with similar names that may be duplicates or related entities
   - Cross-reference vendor addresses with the employee roster
   - Flag vendors with P.O. Box-only addresses

6. Approver analysis:
   - Identify self-approved transactions
   - Flag cost centers with inadequate approval segregation

For each finding:
- Quantify the number of transactions and dollar amounts
- Assess risk level (high/medium/low)
- Recommend specific follow-up procedures

Export:
- Detailed analysis workbook (Excel) with separate tabs for each test
- Executive summary memo (Word) with findings ranked by risk
```

**Gold Standard**: The agent should find all 7 anomaly categories. Provide exact transaction_ids for each planted anomaly. Total exposure from duplicates: $127,340. The Benford's analysis should show statistically significant deviation in the $9,900-$9,999 range. The vendor matching employee address is employee PC-0342.

---

### TC-14: 13-Week Cash Flow Forecast

**Difficulty**: Routine

**Input Files**:
- `balance_sheet_current.xlsx`: Balance sheet as of the most recent week-end (Friday). Includes cash balance, AR detail by aging bucket, AP detail by aging bucket, current portion of debt.
- `ap_aging_report.xlsx`: AP by vendor with due dates for the next 13 weeks.
- `ar_aging_report.xlsx`: AR by customer with expected collection dates (based on historical payment patterns embedded in a "DSO" column).
- `committed_discretionary_expenses.docx`: List of expenses categorized as:
  - Committed (cannot be deferred): payroll, rent, debt service, insurance, utilities
  - Semi-discretionary (can be deferred 2-4 weeks): maintenance, professional fees
  - Discretionary (can be cut): marketing, travel, training, bonuses
  - Include weekly/monthly amounts and payment timing for each

**Prompt**:
```
Build a 13-week cash flow forecast for Cascade Industries.

1. Start with the current cash balance.
2. Project weekly cash inflows based on the AR aging and historical collection patterns.
3. Project weekly cash outflows based on AP due dates and committed expense schedule.
4. Identify the projected cash trough (lowest cash balance and which week).
5. Determine if and when the company would breach its minimum liquidity
   covenant of $2,000,000.
6. If a breach is projected, identify which discretionary expenses could be
   deferred to avoid it, and show the revised forecast.

Export as an Excel workbook with:
- Weekly cash flow detail (inflows, outflows by category, net, cumulative balance)
- Summary dashboard showing the 13-week trend
- Sensitivity analysis: what happens if collections slow by 1 week?
```

**Gold Standard**: Cash trough occurs in Week 8 at $1,340,000 (below the $2M covenant). Deferring marketing ($85K/week) and training ($20K/week) starting Week 5 brings the trough to $2,155,000. If collections slow by 1 week, the trough drops to $780,000 even with deferrals. Pre-compute all weekly values.

---

### TC-15: DCF Valuation

**Difficulty**: Complex

**Input Files**:
- `historical_financials_3yr.xlsx`: 3 years of income statement, balance sheet, and cash flow statement.
- `management_projections.xlsx`: 5-year revenue and EBITDA projections with assumptions (growth rates, margin expansion, capex plan).
- `comparable_companies_trading.xlsx`: Trading data for 10 comparable public companies: market cap, enterprise value, revenue, EBITDA, net income, and computed multiples (EV/Revenue, EV/EBITDA).
- `industry_overview.pdf`: 15-page industry overview with market growth forecasts, competitive dynamics, and risk factors.

**Prompt**:
```
Prepare a DCF valuation of Cascade Industries.

1. Derive unlevered free cash flow (UFCF) from management's projections:
   - EBITDA
   - Less: taxes (use effective rate from historical data)
   - Less: capex (from projections)
   - Less: changes in net working capital (derive from historical NWC/revenue ratios)
2. Compute WACC:
   - Cost of equity: risk-free rate (use 4.2%), equity risk premium (5.5%),
     beta (derive from comparable companies), size premium (2.0%)
   - Cost of debt: derive from the company's interest expense / average debt
   - Capital structure: use comparable companies' average debt/equity ratio
3. Compute terminal value using both:
   - Gordon Growth Model (2.5% perpetuity growth rate)
   - Exit multiple method (use comparable companies' median EV/EBITDA)
4. Discount to present value and compute implied enterprise value range.
5. Compute implied equity value (subtract net debt).
6. Perform sensitivity analysis:
   - WACC ± 1% vs. terminal growth rate ± 0.5% (for Gordon Growth)
   - WACC ± 1% vs. exit multiple ± 1x (for exit multiple method)
7. Draft a one-page valuation summary with the range and key assumptions.

Export:
- DCF model as Excel workbook (with clearly labeled assumptions, calculations,
  and sensitivity tables)
- Valuation summary as Word document
```

**Gold Standard**: Pre-compute all values. WACC should be approximately 10.8%. Enterprise value range: $245M - $305M (Gordon Growth) and $260M - $320M (exit multiple). Mid-point implied equity value approximately $255M. The agent should note that management's projections assume margin expansion of 200bps by Year 5 — an aggressive assumption worth flagging.

---

## 6. TEST CASES — CROSS-SERVICE LINE

### TC-16: Engagement Letter Generation

**Difficulty**: Routine

**Input Files**:
- `client_profile.docx`: Cascade Industries profile with entity count, revenue tier, complexity factors, key contacts.
- `fee_schedule.xlsx`: Fee matrix with columns: Service Type, Revenue Tier, Entity Count, Base Fee, Per-Entity Adder, Complexity Multiplier.
- `engagement_letter_template.docx`: Template with merge fields (placeholders in `<<field_name>>` format): `<<client_name>>`, `<<engagement_scope>>`, `<<fee_amount>>`, `<<payment_terms>>`, `<<start_date>>`, `<<partner_name>>`.

**Prompt**:
```
Generate a draft engagement letter for a combined audit and tax engagement
for Cascade Industries using the template provided.

1. Look up the appropriate fees from the fee schedule based on the client's
   revenue tier and entity count.
2. The audit fee should include the complexity multiplier for public-readiness
   (1.15x) since the company is considering an IPO.
3. The tax fee should cover federal and all state returns for the parent and
   3 subsidiaries.
4. Populate the template with the correct values.
5. Set the start date to March 15, 2026 and payment terms to Net 30.
6. Use "Sarah Chen" as the engagement partner.

Save as a Word document.
```

**Gold Standard**: Pre-compute fees. Audit base fee $285,000 × 1.15 complexity = $327,750. Tax base fee $95,000 + $12,000 per additional entity × 3 = $131,000. Total engagement: $458,750. All merge fields correctly populated. Template formatting preserved.

---

### TC-17: Multi-File Deliverable Assembly

**Difficulty**: Complex

**Input Files**:
- 6 completed workpaper sections (mix of xlsx and docx):
  - `01_executive_summary.docx`
  - `02_financial_analysis.xlsx`
  - `03_industry_overview.docx`
  - `04_risk_assessment.docx`
  - `05_detailed_findings.xlsx`
  - `06_recommendations.docx`
- `cover_page_template.docx`: Branded cover page with placeholders.
- `formatting_guide.pdf`: Specifies section order, page numbering rules (Roman numerals for front matter, Arabic for body), header/footer requirements, and font standards.

**Prompt**:
```
Assemble the 6 workpaper sections into a single client deliverable.

1. Use the cover page template (populate with "Cascade Industries — FY2025
   Financial Advisory Report" and today's date).
2. Follow the section order specified in the formatting guide.
3. Add a table of contents after the cover page.
4. Apply consistent formatting per the formatting guide (fonts, headers, spacing).
5. Add page numbers (Roman numerals for TOC, Arabic starting at 1 for Section 1).
6. For the Excel sections, extract key tables and charts and embed them
   in the appropriate location in the document flow.
7. Export as a single PDF.

The final deliverable should look like something you'd hand to a C-suite client.
```

**Gold Standard**: This is primarily a format and assembly test. Grade on: correct section order, TOC accuracy, consistent formatting, page number scheme, professional appearance, and successful conversion to PDF. The Excel data tables should be embedded correctly (not as screenshots of random cell ranges, but as properly formatted tables showing the key data).

---

### TC-18: Prior Year Workpaper Rollforward

**Difficulty**: Adversarial

**Input Files**:
- `prior_year_workpapers/` folder containing 10 files:
  - 6 xlsx workpapers (revenue, expenses, balance sheet sections, cash, etc.)
  - 4 docx memos (planning, risk assessment, summary, management letter points)
  - Each contains FY2024 data with formulas referencing specific cells
- `current_year_data/` folder containing:
  - Updated trial balance (CSV — note: format change from xlsx)
  - Updated bank statements (CSV)
  - Updated lease schedule (xlsx — with 2 new leases added)
  - Management's updated projections (docx — new format from prior year's xlsx)
  - A new file not present in prior year: `goodwill_impairment_analysis.xlsx`

**Prompt**:
```
Roll forward the FY2024 audit workpapers to FY2025 using the current year data.

1. For each prior year workpaper, identify which data needs to be updated.
2. Map the current year data files to the corresponding prior year workpapers,
   noting any format changes (e.g., CSV vs. xlsx).
3. Update all numerical data with current year figures.
4. Preserve the workpaper structure and any formulas/commentary that are still relevant.
5. Flag the following for manager attention:
   - Any structural changes in the client's data (new accounts, format changes,
     renamed fields)
   - The new goodwill impairment analysis file (not present in prior year —
     suggests a new accounting issue)
   - Any areas where prior year commentary may no longer be applicable
6. Update the planning memo with current year scope considerations.

Export the rolled-forward workpapers to a new folder.
```

**Gold Standard**: The agent should successfully update 8 of 10 files. The 2 remaining files require judgment calls that should be flagged (the management projections changed format, and the planning memo needs substantive rewriting, not just data updates). The CSV format change should be handled transparently. The goodwill impairment file should be flagged as a significant new audit area. Pre-define the expected state of each rolled-forward file.

---

## 7. SCORING SYSTEM

### 7.1 Per-Test-Case Rubric

Every test case is scored on 5 dimensions using a 1–3 scale:

| Dimension | Score 3 | Score 2 | Score 1 |
|-----------|---------|---------|---------|
| **Correctness** | All values match gold standard (±0.5% tolerance for computed values) | Minor errors that don't change conclusions | Material errors in key outputs |
| **Completeness** | All requested deliverables produced, all components present | Missing minor elements (e.g., one sheet, one section) | Missing major deliverables or sections |
| **Format Compliance** | Valid files, professional formatting, correct file types | Functional but poorly formatted, or minor type issues | Broken files, wrong file type, or unusable output |
| **Robustness** | Handled all edge cases (messy data, format variations, ambiguity) | Handled most edge cases, stumbled on some | Failed on edge cases or didn't attempt to handle them |
| **Communication** | Proactively explained approach, flagged uncertainties, identified errors | Adequate communication, some gaps | Silent about approach, missed obvious items to flag |

### 7.2 Automated Grading (auto_grader.py)

Write a Python script that can automatically grade the mechanical components:

```python
# auto_grader.py should implement:

class TestCaseGrader:
    def __init__(self, test_case_id, gold_standard_path, agent_output_path):
        ...

    def grade_correctness(self):
        """Compare numerical outputs to gold standard values.
        Returns: dict of {metric: {expected, actual, match, tolerance_pct}}"""
        ...

    def grade_completeness(self):
        """Check for presence of required files, sheets, sections.
        Returns: dict of {requirement: present/missing}"""
        ...

    def grade_format(self):
        """Validate file types, check if files open without errors.
        Returns: dict of {check: pass/fail}"""
        ...

    def verify_canaries(self):
        """Check if the agent read the correct files by finding canary values.
        Returns: dict of {canary: found/not_found}"""
        ...

    def check_error_detection(self):
        """Check if the agent identified planted errors.
        Returns: dict of {error_id: detected/missed}"""
        ...
```

The auto-grader should output a JSON report per test case and an aggregate summary.

### 7.3 Human Grading Template

Create `scoring_template.xlsx` with:
- Sheet 1: "Scorecard" — one row per test case, columns for each dimension (1-3), plus notes
- Sheet 2: "Grader Instructions" — detailed guidance for each dimension and each test case
- Sheet 3: "Inter-Rater Agreement" — tracks two raters' scores and computes Cohen's kappa
- Sheet 4: "Aggregate Dashboard" — pivot tables showing pass rates by: service line, capability axis, difficulty tier

### 7.4 Pass/Fail Criteria

| Threshold | Definition |
|-----------|-----------|
| **Pass** | Average score ≥ 2.4 across all 5 dimensions |
| **Conditional Pass** | Average score ≥ 2.0 AND no dimension scored 1 |
| **Fail** | Average score < 2.0 OR any dimension scored 1 |

### 7.5 Capability Matrix

Track results in a 3D matrix:

- **Axis 1 — Service Line**: Audit, Tax, Advisory, Cross-Service
- **Axis 2 — Capability**: File Reading, File Writing, Data Analysis (SQL/DataFrame), RAG, Multi-Step Workflow
- **Axis 3 — Difficulty**: Routine, Complex, Adversarial

Each test case maps to one or more cells in this matrix. The aggregate view tells you things like "the agent is strong at routine file reading in Audit but weak at adversarial RAG in Tax" — which is actionable for prioritizing improvements.

---

## 8. TEST EXECUTION PROTOCOL

### 8.1 Environment Setup

For each test case:

1. Create a clean working directory
2. Copy only the files listed in the test case's input files to the working directory
3. Do not provide any files from other test cases or the gold standards
4. Provide the prompt exactly as written in `prompt.md` (no additional context)
5. Record the agent's full interaction (all tool calls, outputs, messages)

### 8.2 Recording Requirements

For each test run, capture:
- Start and end timestamps
- Full agent interaction log (every message, tool call, and output)
- All files produced by the agent (copied to a results directory)
- Any errors or exceptions encountered
- Token count / cost if applicable

### 8.3 Evaluation Sequence

1. Run auto-grader on agent outputs → produces JSON scores for mechanical dimensions
2. Human Rater 1 reviews interaction log and outputs → scores all 5 dimensions
3. Human Rater 2 independently scores all 5 dimensions
4. Reconcile any disagreements (>1 point difference on any dimension)
5. Record final scores in scoring template
6. Generate capability matrix dashboard

### 8.4 Regression Testing

After each agent improvement cycle:
1. Re-run all test cases (the fixed seed ensures identical input data)
2. Compare scores to prior run
3. Flag any regressions (score decreased on any dimension for any test case)
4. Track improvement trends over time

---

## 9. DELIVERABLES CHECKLIST

When complete, the producing agent must deliver all of the following:

- [ ] `config.yaml` — Master configuration file
- [ ] `generate_test_suite.py` — Deterministic generator script (SEED=42)
- [ ] `manifest.json` — Complete file manifest
- [ ] `canary_registry.json` — All canary values mapped to files
- [ ] `error_registry.json` — All 25 planted errors documented
- [ ] `shared_data/` — All shared reference files (COA, employee roster, org chart, financials, templates)
- [ ] `test_cases/TC-01/` through `test_cases/TC-18/` — Each containing:
  - [ ] `prompt.md` — Exact prompt to deliver to the agent
  - [ ] `input_files/` — All input files for that test case
  - [ ] `expected_behavior.md` — Notes on expected agent behavior
- [ ] `gold_standards/` — Gold standard outputs for all 18 test cases:
  - [ ] JSON files with expected numerical values and structural requirements
  - [ ] Reference output files (xlsx, docx) where applicable
- [ ] `scoring/rubrics.yaml` — Scoring rubrics for all test cases
- [ ] `scoring/auto_grader.py` — Automated grading script
- [ ] `scoring/scoring_template.xlsx` — Human grading template with dashboard
- [ ] `README.md` — Setup and execution instructions

### Quality Gates

Before declaring the suite complete:

1. Run `generate_test_suite.py` twice — outputs must be byte-identical
2. Run `auto_grader.py` against the gold standard outputs — must score 3/3/3/3/3 (i.e., the gold standards must pass their own tests)
3. Every canary in `canary_registry.json` must be findable in the corresponding file
4. Every error in `error_registry.json` must be verifiable in the corresponding file
5. All generated xlsx files must open without errors in Excel/LibreOffice
6. All generated docx files must open without errors in Word/LibreOffice
7. All generated PDFs must be readable by a standard PDF viewer
8. At least one person must manually walk through 3 test cases end-to-end to verify the prompts are clear and the gold standards are achievable

---

## 10. IMPORTANT DESIGN PRINCIPLES

1. **Realism over cleverness**: Every test case should feel like something an actual Big 4 professional would encounter on a Tuesday afternoon. Avoid contrived scenarios.

2. **Cross-referential integrity**: The numbers must tie across test cases. If TC-01 uses a trial balance showing $200M revenue, TC-06 (tax provision) must use the same number. The master data model is the single source of truth.

3. **Graded difficulty within each service line**: Each service line has Routine, Complex, and Adversarial tests. Routine tests should be passable by a competent agent. Adversarial tests should expose edge cases.

4. **Separation of input and expected output**: The agent under test never sees gold standards or expected behavior notes. These are only for the evaluators.

5. **Determinism**: The same seed must produce the same files every time. No external API calls or non-deterministic operations in the generator.

6. **Professional tone in prompts**: Write prompts the way a real engagement manager would delegate work — direct, specific about deliverables, but not overly prescriptive about methodology.

7. **Deliberate imperfection in inputs**: Real client data is messy. Include formatting issues, inconsistencies, and gaps — but document them so you can grade whether the agent handled them.

8. **Test the agent's judgment, not just its computation**: Several test cases require the agent to push back on management (TC-11: aggressive adjustments), flag risks (TC-12: red flags), or say "I can't determine this" (TC-04: OCR uncertainty). These judgment calls are the most valuable thing to measure.
