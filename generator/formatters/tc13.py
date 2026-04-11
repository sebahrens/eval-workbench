"""Formatter: TC-13 — Forensic AP Transaction Analysis (Advisory, Complex).

Emits:
- test_cases/TC-13/input_files/ap_transactions_fy2025.csv
  52,000-row AP transaction ledger with 7 planted anomaly categories
- test_cases/TC-13/prompt.md
- test_cases/TC-13/expected_behavior.md
- gold_standards/TC-13_gold.json

Anomaly categories:
1. Duplicate payments (4 exact + 4 near-dup, exposure $127,340)
2. Benford's Law violation (35 txns $9,900-$9,999)
3. Round-number payments (12 txns to Pacific Consulting Group)
4. Temporal anomalies (15 weekend/holiday + 8 invoice-after-payment)
5. Vendor anomalies (2 similar-name + 1 employee-address vendor)
6. Split transactions (3 sets below $10K threshold)
7. Approver anomalies (1 self-approved + 1 single-approver cost center)

Uses the canonical model for employees; ap_ledger.py for transaction generation.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any

from generator.canaries import CanaryRegistry, embed_canary_csv_comment
from generator.errors import ErrorRegistry
from generator.golds.framework import GoldStandard, register_gold
from generator.manifest import Manifest
from generator.model.ap_ledger import (
    _EMPLOYEE_VENDOR_NAME,
    _JKL_INC_NAME,
    _JKL_LLC_NAME,
    _PACIFIC_CONSULTING_NAME,
    _TARGET_EMPLOYEE_ID,
    APLedgerResult,
)
from generator.model.build import CascadeModel

# ── Constants ────────────────────────────────────────────────────────────────

_TC = "TC-13"
_INPUT_DIR = f"test_cases/{_TC}/input_files"
_CSV_FILE_KEY = "tc13_ap_transactions"

# CSV column order matches prompt.md §TC-13.
_CSV_COLUMNS = [
    "transaction_id",
    "date",
    "vendor_id",
    "vendor_name",
    "amount",
    "description",
    "approver",
    "cost_center",
    "payment_method",
    "invoice_number",
]


# ── CSV writer ───────────────────────────────────────────────────────────────


def _write_ap_csv(
    model: CascadeModel,
    output_dir: Path,
    canaries: CanaryRegistry,
    manifest: Manifest,
) -> APLedgerResult:
    """Write ap_transactions_fy2025.csv and return the ledger result."""
    result = model.ap_ledger

    rel_path = f"{_INPUT_DIR}/ap_transactions_fy2025.csv"
    abs_path = output_dir / rel_path

    canary = canaries.canary_for(_CSV_FILE_KEY)
    canary_line = embed_canary_csv_comment(canary)

    abs_path.parent.mkdir(parents=True, exist_ok=True)
    with open(abs_path, "w", newline="") as f:
        f.write(canary_line)
        f.write(",".join(_CSV_COLUMNS) + "\n")

        for txn in result.transactions:
            # Format amount with 2 decimal places
            amount_str = f"{txn.amount:.2f}"
            # Escape any commas in description by quoting
            desc = txn.description
            if "," in desc or '"' in desc:
                desc = '"' + desc.replace('"', '""') + '"'
            # Vendor name may contain commas
            vname = txn.vendor_name
            if "," in vname or '"' in vname:
                vname = '"' + vname.replace('"', '""') + '"'

            row = [
                txn.transaction_id,
                txn.date.isoformat(),
                txn.vendor_id,
                vname,
                amount_str,
                desc,
                txn.approver,
                txn.cost_center,
                txn.payment_method,
                txn.invoice_number,
            ]
            f.write(",".join(row) + "\n")

    canaries.set_location(
        _CSV_FILE_KEY,
        rel_path,
        "Line 1 comment: # CANARY: ...",
    )
    manifest.register(rel_path, "csv")

    return result


# ── Prompt & Expected Behavior ───────────────────────────────────────────────


def _write_prompt(output_dir: Path) -> None:
    """Write test_cases/TC-13/prompt.md per spec."""
    text = """\
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
"""
    path = output_dir / f"test_cases/{_TC}/prompt.md"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text)


def _write_expected_behavior(output_dir: Path) -> None:
    """Write test_cases/TC-13/expected_behavior.md per spec."""
    text = """\
# TC-13: Forensic AP Transaction Analysis — Expected Behavior

## Key Findings the Agent Should Identify

### 1. Duplicate Payments (8 instances, exposure $127,340)
- **4 exact duplicates**: Same vendor, amount, and invoice number processed
  on different dates (1-5 days apart). The agent should match on the
  (vendor_id, amount, invoice_number) tuple.
- **4 near-duplicates**: Same vendor and amount, but invoice numbers differ
  by exactly 1 digit in the last position — suggesting a typo or intentional
  variation. The agent should use fuzzy matching on invoice numbers.
- Total duplicate exposure: $127,340 (sum of the duplicated amounts).

### 2. Benford's Law Violation
- 35 transactions clustered between $9,900-$9,999, just below the $10,000
  approval threshold. First-two-digit analysis should show statistically
  significant overrepresentation of digits 99 compared to Benford's expected
  distribution.
- The agent should compute a chi-squared or similar goodness-of-fit test.

### 3. Round-Number Payments to Shell-Company-Style Vendor
- 12 transactions to "Pacific Consulting Group" (VEND-050) at exactly
  $5,000.00, $10,000.00, or $25,000.00. Round-number payments to a single
  vendor with a generic consulting name are a classic fraud indicator.

### 4. Temporal Anomalies (23 total)
- **15 weekend/holiday approvals**: Transactions approved on dates that
  fall on weekends or US federal holidays. Legitimate AP processing should
  not occur on non-business days.
- **8 invoice-after-payment**: The invoice date (embedded in the description
  field) is 3-15 days after the payment date. This is a red flag for
  fictitious invoices or backdated approvals.

### 5. Vendor Anomalies
- **Similar-name vendors**: "JKL Services LLC" (VEND-051) and "JKL Services
  Inc." (VEND-052) — two vendors with P.O. Box-only addresses sharing a
  nearly identical name. May indicate a related-party scheme or vendor
  duplication.
- **Employee-address vendor**: "Willow Creek Consulting" (VEND-053) has an
  address matching employee PC-0342 from the employee roster. This is a
  classic conflict-of-interest / ghost vendor indicator.

### 6. Split Transactions (3 sets)
- Three pairs of transactions to the same vendor on consecutive business
  days, with combined amounts exceeding the $10,000 approval threshold:
  - $7,200 + $7,300 = $14,500
  - $4,800 + $4,950 = $9,750
  - $6,100 + $5,900 = $12,000
- Same approver and cost center within each pair.

### 7. Approver Anomalies
- **Self-approved reimbursement**: One transaction where the vendor name
  (payee) matches the approver name — an employee approved their own
  reimbursement.
- **Single-approver concentration**: Cost center 2600 has 95% (38/40) of
  transactions approved by the same person, indicating inadequate
  segregation of duties.

## Data Challenges

- **Volume**: 52,000 rows require efficient analysis — the agent cannot
  manually inspect each row.
- **Invoice-date encoding**: The invoice date for temporal anomaly #4b is
  embedded in the description field, not a separate column. The agent must
  parse descriptions to extract invoice dates.
- **Fuzzy matching required**: Near-duplicate detection and similar-vendor
  identification require string similarity algorithms, not exact matching.
- **Cross-reference needed**: The employee-address vendor requires joining
  AP data with the employee roster (shared_data).

## Expected Output Structure

### Analysis Workbook (Excel):
- Tab per analysis type (duplicates, Benford's, thresholds, temporal,
  vendor, approver)
- Transaction IDs and details for each flagged item
- Summary statistics and risk ratings

### Executive Summary Memo (Word):
- Findings ranked by risk (high/medium/low)
- Dollar exposure quantification
- Recommended follow-up procedures for each category
"""
    path = output_dir / f"test_cases/{_TC}/expected_behavior.md"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text)


# ── Gold Standard ────────────────────────────────────────────────────────────


@register_gold("TC-13")
def _tc13_gold(
    canaries: CanaryRegistry,
    errors: ErrorRegistry,
    **model_kwargs: Any,
) -> GoldStandard:
    """TC-13 gold standard: forensic AP transaction analysis."""
    model: CascadeModel = model_kwargs["model"]

    # Use canonical AP ledger from the model.
    result = model.ap_ledger

    # Build expected outputs from the anomaly index.
    anomaly_index = result.anomaly_index

    # Compute duplicate exposure
    dup_txn_ids = (
        anomaly_index.get("exact_duplicate", [])
        + anomaly_index.get("near_duplicate", [])
    )
    dup_txns = [t for t in result.transactions if t.transaction_id in dup_txn_ids]
    # Exposure = sum of the second copy in each pair (every other txn)
    exact_exposure = sum(
        t.amount for i, t in enumerate(
            [t for t in dup_txns if t.anomaly_type == "exact_duplicate"]
        ) if i % 2 == 1
    )
    near_exposure = sum(
        t.amount for i, t in enumerate(
            [t for t in dup_txns if t.anomaly_type == "near_duplicate"]
        ) if i % 2 == 1
    )
    total_exposure = exact_exposure + near_exposure

    expected_outputs: dict[str, Any] = {
        "total_transactions": len(result.transactions),
        "anomaly_categories": {
            "duplicate_payments": {
                "exact_duplicates": {
                    "count": len(anomaly_index.get("exact_duplicate", [])),
                    "transaction_ids": sorted(
                        anomaly_index.get("exact_duplicate", [])
                    ),
                },
                "near_duplicates": {
                    "count": len(anomaly_index.get("near_duplicate", [])),
                    "transaction_ids": sorted(
                        anomaly_index.get("near_duplicate", [])
                    ),
                },
                "total_exposure": str(total_exposure),
            },
            "benford_violation": {
                "count": len(anomaly_index.get("benford_violation", [])),
                "transaction_ids": sorted(
                    anomaly_index.get("benford_violation", [])
                ),
                "range": "$9,900-$9,999",
            },
            "round_number": {
                "count": len(anomaly_index.get("round_number", [])),
                "vendor": _PACIFIC_CONSULTING_NAME,
                "transaction_ids": sorted(
                    anomaly_index.get("round_number", [])
                ),
            },
            "temporal_anomalies": {
                "weekend_holiday": {
                    "count": len(anomaly_index.get("weekend_holiday", [])),
                    "transaction_ids": sorted(
                        anomaly_index.get("weekend_holiday", [])
                    ),
                },
                "invoice_after_payment": {
                    "count": len(
                        anomaly_index.get("invoice_after_payment", [])
                    ),
                    "transaction_ids": sorted(
                        anomaly_index.get("invoice_after_payment", [])
                    ),
                },
            },
            "vendor_anomalies": {
                "similar_name_vendors": {
                    "count": len(
                        anomaly_index.get("similar_name_vendor", [])
                    ),
                    "vendors": [_JKL_LLC_NAME, _JKL_INC_NAME],
                    "transaction_ids": sorted(
                        anomaly_index.get("similar_name_vendor", [])
                    ),
                },
                "employee_address_vendor": {
                    "count": len(
                        anomaly_index.get("employee_address_vendor", [])
                    ),
                    "vendor": _EMPLOYEE_VENDOR_NAME,
                    "matching_employee": _TARGET_EMPLOYEE_ID,
                    "transaction_ids": sorted(
                        anomaly_index.get("employee_address_vendor", [])
                    ),
                },
            },
            "split_transactions": {
                "count": len(anomaly_index.get("split_transaction", [])),
                "sets": 3,
                "transaction_ids": sorted(
                    anomaly_index.get("split_transaction", [])
                ),
            },
            "approver_anomalies": {
                "self_approved": {
                    "count": len(anomaly_index.get("self_approved", [])),
                    "transaction_ids": sorted(
                        anomaly_index.get("self_approved", [])
                    ),
                },
                "single_approver_cc": {
                    "count": len(
                        anomaly_index.get("single_approver_cc", [])
                    ),
                    "cost_center": "2600",
                    "transaction_ids": sorted(
                        anomaly_index.get("single_approver_cc", [])
                    ),
                },
            },
        },
    }

    canary_verification = {
        "read_ap_transactions": canaries.canary_for(_CSV_FILE_KEY),
    }

    scoring_hints = {
        "completeness": (
            "Agent must identify all 7 anomaly categories. "
            "Partial credit for finding 5-6 of 7."
        ),
        "accuracy": (
            "Duplicate exposure must be $127,340. "
            "Benford analysis must flag $9,900-$9,999 cluster. "
            f"Employee-address vendor must link to {_TARGET_EMPLOYEE_ID}."
        ),
        "methodology": (
            "Agent should use statistical tests (chi-squared for Benford's), "
            "fuzzy matching (for near-duplicates and similar vendors), "
            "and cross-referencing (employee roster for vendor addresses)."
        ),
    }

    return GoldStandard(
        test_case=_TC,
        expected_outputs=expected_outputs,
        canary_verification=canary_verification,
        error_detection={},  # No planted ERR-xxx errors — anomalies are the test
        scoring_hints=scoring_hints,
    )


# ── Public API ───────────────────────────────────────────────────────────────


def emit_tc13(
    model: CascadeModel,
    output_dir: Path,
    canaries: CanaryRegistry,
    errors: ErrorRegistry,
    manifest: Manifest,
) -> None:
    """Emit all TC-13 files."""
    _write_ap_csv(model, output_dir, canaries, manifest)
    _write_prompt(output_dir)
    _write_expected_behavior(output_dir)
