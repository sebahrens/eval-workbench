"""Cascade Industries general ledger posting engine.

Provides the JournalEntry dataclass and Ledger class for posting balanced
journal entries and querying balances. No entries are posted at import time —
this is the engine, not the data.
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass
from decimal import Decimal

from generator.model.coa import ACCOUNTS_BY_NUMBER


@dataclass(frozen=True)
class JournalEntryLine:
    """A single line (leg) of a journal entry."""

    account: str  # 4-digit account number
    debit: Decimal
    credit: Decimal
    memo: str = ""

    def __post_init__(self) -> None:
        if self.debit < 0 or self.credit < 0:
            raise ValueError(f"Debit/credit must be non-negative: debit={self.debit}, credit={self.credit}")
        if self.debit > 0 and self.credit > 0:
            raise ValueError(f"Line cannot have both debit and credit: {self.account}")
        if self.debit == 0 and self.credit == 0:
            raise ValueError(f"Line must have a debit or credit: {self.account}")


@dataclass(frozen=True)
class JournalEntry:
    """A balanced journal entry to be posted to the ledger."""

    date: datetime.date
    entity_code: str  # Two-letter entity code (CI, PC, AM, DS)
    description: str
    lines: tuple[JournalEntryLine, ...]

    def total_debits(self) -> Decimal:
        return sum(line.debit for line in self.lines)

    def total_credits(self) -> Decimal:
        return sum(line.credit for line in self.lines)

    def is_balanced(self) -> bool:
        return self.total_debits() == self.total_credits()


class UnbalancedEntryError(Exception):
    """Raised when attempting to post an unbalanced journal entry."""


class InvalidAccountError(Exception):
    """Raised when a journal entry references an account not in the COA."""


class Ledger:
    """General ledger that accepts balanced journal entries and supports queries."""

    def __init__(self) -> None:
        self._entries: list[JournalEntry] = []

    @property
    def entries(self) -> list[JournalEntry]:
        """All posted journal entries (read-only copy)."""
        return list(self._entries)

    def post(self, entry: JournalEntry) -> None:
        """Post a balanced journal entry to the ledger.

        Raises UnbalancedEntryError if debits != credits.
        Raises InvalidAccountError if any line references an unknown account.
        """
        # Validate accounts exist in COA
        for line in entry.lines:
            if line.account not in ACCOUNTS_BY_NUMBER:
                raise InvalidAccountError(f"Account {line.account} not in chart of accounts")

        if not entry.is_balanced():
            raise UnbalancedEntryError(
                f"Entry '{entry.description}' is unbalanced: "
                f"debits={entry.total_debits()}, credits={entry.total_credits()}"
            )

        self._entries.append(entry)

    def balance_by_account(
        self,
        entity_code: str,
        as_of_date: datetime.date | None = None,
    ) -> dict[str, Decimal]:
        """Return net balance per account for a given entity.

        Balance = sum(debits) - sum(credits) for each account, filtered by
        entity and optionally by date (<= as_of_date).

        Returns a dict of account_number → net balance (positive = debit balance).
        Only accounts with non-zero balances are included.
        """
        balances: dict[str, Decimal] = {}
        for entry in self._entries:
            if entry.entity_code != entity_code:
                continue
            if as_of_date is not None and entry.date > as_of_date:
                continue
            for line in entry.lines:
                bal = balances.get(line.account, Decimal(0))
                bal += line.debit - line.credit
                balances[line.account] = bal

        return {acct: bal for acct, bal in sorted(balances.items()) if bal != 0}

    def filter_by_date_range(
        self,
        start: datetime.date,
        end: datetime.date,
    ) -> list[JournalEntry]:
        """Return entries with date in [start, end] inclusive."""
        return [e for e in self._entries if start <= e.date <= end]

    def filter_by_entity(self, entity_code: str) -> list[JournalEntry]:
        """Return all entries for a given entity code."""
        return [e for e in self._entries if e.entity_code == entity_code]

    def account_detail(
        self,
        account: str,
        entity_code: str | None = None,
    ) -> list[tuple[JournalEntry, JournalEntryLine]]:
        """Return all (entry, line) pairs touching a specific account."""
        results = []
        for entry in self._entries:
            if entity_code is not None and entry.entity_code != entity_code:
                continue
            for line in entry.lines:
                if line.account == account:
                    results.append((entry, line))
        return results
