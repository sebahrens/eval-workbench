"""Tests for the chart of accounts (generator/model/coa.py)."""

from generator.model.coa import (
    ACCOUNTS,
    ACCOUNTS_BY_NUMBER,
    AccountType,
    NormalBalance,
    validate_coa,
)


def test_account_count():
    """~120 accounts as specified in prompt.md §1.5."""
    assert 100 <= len(ACCOUNTS) <= 150, f"Expected ~120 accounts, got {len(ACCOUNTS)}"


def test_all_numbers_unique():
    assert len(ACCOUNTS_BY_NUMBER) == len(ACCOUNTS), "Duplicate account numbers"


def test_four_digit_numbers():
    for acct in ACCOUNTS:
        assert len(acct.number) == 4 and acct.number.isdigit(), f"Bad number: {acct.number}"


def test_prefix_matches_type():
    """Account type must match the leading digit."""
    prefix_type = {
        "1": AccountType.ASSET,
        "2": AccountType.LIABILITY,
        "3": AccountType.EQUITY,
        "4": AccountType.REVENUE,
        "5": AccountType.COGS,
        "6": AccountType.OPEX,
        "7": AccountType.OTHER,
        "8": AccountType.TAX,
        "9": AccountType.INTERCOMPANY,
    }
    for acct in ACCOUNTS:
        expected = prefix_type[acct.number[0]]
        assert acct.account_type == expected, (
            f"{acct.number} ({acct.name}): type={acct.account_type.value}, expected={expected.value}"
        )


def test_all_nine_prefixes_present():
    prefixes = {acct.number[0] for acct in ACCOUNTS}
    assert prefixes == {"1", "2", "3", "4", "5", "6", "7", "8", "9"}


def test_sorted_by_number():
    numbers = [acct.number for acct in ACCOUNTS]
    assert numbers == sorted(numbers)


def test_ic_pairs_present():
    """Every IC receivable has a matching payable and vice-versa."""
    ic = [a for a in ACCOUNTS if a.number.startswith("90")]
    receivables = {a.name.replace("IC Receivable — ", "") for a in ic if "Receivable" in a.name}
    payables = {a.name.replace("IC Payable — ", "") for a in ic if "Payable" in a.name}
    assert receivables == payables, f"Mismatched IC pairs: recv={receivables}, pay={payables}"


def test_rd_accounts_exist():
    """Granular R&D expense accounts required by TC-08."""
    rd_accounts = [a for a in ACCOUNTS if a.number.startswith("63")]
    assert len(rd_accounts) >= 5, f"Need ≥5 granular R&D accounts, got {len(rd_accounts)}"


def test_lease_accounts_exist():
    """Separate ROU asset + lease liability accounts required by TC-04 (ASC 842)."""
    rou = [a for a in ACCOUNTS if "ROU" in a.name]
    lease_liab = [a for a in ACCOUNTS if "Lease Liability" in a.name]
    assert len(rou) >= 2, f"Need ≥2 ROU asset accounts, got {len(rou)}"
    assert len(lease_liab) >= 2, f"Need ≥2 lease liability accounts, got {len(lease_liab)}"


def test_contra_accounts_have_opposite_balance():
    """Contra-asset accounts (allowance, accum depr) should be credit-normal."""
    contras = [
        a for a in ACCOUNTS
        if (a.account_type == AccountType.ASSET and "Accum" in a.name) or "Allowance" in a.name
    ]
    for acct in contras:
        if acct.account_type == AccountType.ASSET:
            assert acct.normal_balance == NormalBalance.CREDIT, (
                f"Contra-asset {acct.number} ({acct.name}) should be credit-normal"
            )


def test_validate_coa_clean():
    """The built-in validator should find zero errors."""
    errors = validate_coa()
    assert errors == [], f"Validation errors: {errors}"
