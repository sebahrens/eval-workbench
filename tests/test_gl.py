"""Tests for the GL posting engine (generator/model/gl.py)."""

import datetime
from decimal import Decimal

import pytest

from generator.model.gl import (
    InvalidAccountError,
    JournalEntry,
    JournalEntryLine,
    Ledger,
    UnbalancedEntryError,
)

# ── Helpers ──────────────────────────────────────────────────────────────────


def _line(account: str, debit: str = "0", credit: str = "0", memo: str = "") -> JournalEntryLine:
    return JournalEntryLine(account=account, debit=Decimal(debit), credit=Decimal(credit), memo=memo)


def _je(
    entity: str = "PC",
    desc: str = "test",
    lines: tuple[JournalEntryLine, ...] = (),
    date: datetime.date | None = None,
) -> JournalEntry:
    return JournalEntry(
        date=date or datetime.date(2025, 1, 15),
        entity_code=entity,
        description=desc,
        lines=lines,
    )


# ── JournalEntryLine validation ─────────────────────────────────────────────


class TestJournalEntryLine:
    def test_valid_debit_line(self) -> None:
        line = _line("1010", debit="1000")
        assert line.debit == Decimal("1000")
        assert line.credit == Decimal("0")

    def test_valid_credit_line(self) -> None:
        line = _line("2010", credit="500")
        assert line.credit == Decimal("500")

    def test_both_debit_and_credit_raises(self) -> None:
        with pytest.raises(ValueError, match="both debit and credit"):
            _line("1010", debit="100", credit="100")

    def test_zero_both_raises(self) -> None:
        with pytest.raises(ValueError, match="must have a debit or credit"):
            _line("1010", debit="0", credit="0")

    def test_negative_raises(self) -> None:
        with pytest.raises(ValueError, match="non-negative"):
            _line("1010", debit="-100")


# ── JournalEntry ─────────────────────────────────────────────────────────────


class TestJournalEntry:
    def test_balanced_entry(self) -> None:
        entry = _je(lines=(
            _line("1010", debit="5000"),
            _line("4010", credit="5000"),
        ))
        assert entry.is_balanced()
        assert entry.total_debits() == Decimal("5000")
        assert entry.total_credits() == Decimal("5000")

    def test_unbalanced_entry(self) -> None:
        entry = _je(lines=(
            _line("1010", debit="5000"),
            _line("4010", credit="4999"),
        ))
        assert not entry.is_balanced()


# ── Ledger.post ──────────────────────────────────────────────────────────────


class TestLedgerPost:
    def test_post_balanced_entry(self) -> None:
        ledger = Ledger()
        entry = _je(lines=(
            _line("1010", debit="1000"),
            _line("4010", credit="1000"),
        ))
        ledger.post(entry)
        assert len(ledger.entries) == 1

    def test_post_unbalanced_raises(self) -> None:
        ledger = Ledger()
        entry = _je(lines=(
            _line("1010", debit="1000"),
            _line("4010", credit="999"),
        ))
        with pytest.raises(UnbalancedEntryError, match="unbalanced"):
            ledger.post(entry)
        assert len(ledger.entries) == 0

    def test_post_invalid_account_raises(self) -> None:
        ledger = Ledger()
        entry = _je(lines=(
            _line("9999", debit="100"),
            _line("1010", credit="100"),
        ))
        with pytest.raises(InvalidAccountError, match="9999"):
            ledger.post(entry)


# ── Ledger queries ───────────────────────────────────────────────────────────


class TestLedgerQueries:
    @pytest.fixture()
    def populated_ledger(self) -> Ledger:
        ledger = Ledger()
        # PC: Cash debit 5000, Revenue credit 5000 — Jan 15
        ledger.post(_je(
            entity="PC",
            desc="Sale",
            date=datetime.date(2025, 1, 15),
            lines=(
                _line("1010", debit="5000"),
                _line("4010", credit="5000"),
            ),
        ))
        # PC: COGS debit 2000, Inventory credit 2000 — Jan 20
        ledger.post(_je(
            entity="PC",
            desc="COGS",
            date=datetime.date(2025, 1, 20),
            lines=(
                _line("5010", debit="2000"),
                _line("1200", credit="2000"),
            ),
        ))
        # AM: Cash debit 3000, Revenue credit 3000 — Feb 1
        ledger.post(_je(
            entity="AM",
            desc="AM Sale",
            date=datetime.date(2025, 2, 1),
            lines=(
                _line("1010", debit="3000"),
                _line("4020", credit="3000"),
            ),
        ))
        return ledger

    def test_balance_by_account(self, populated_ledger: Ledger) -> None:
        balances = populated_ledger.balance_by_account("PC")
        assert balances["1010"] == Decimal("5000")  # Cash
        assert balances["4010"] == Decimal("-5000")  # Revenue (credit normal)
        assert balances["5010"] == Decimal("2000")  # COGS
        assert balances["1200"] == Decimal("-2000")  # Inventory drawn down

    def test_balance_by_account_with_date(self, populated_ledger: Ledger) -> None:
        balances = populated_ledger.balance_by_account("PC", as_of_date=datetime.date(2025, 1, 15))
        assert balances["1010"] == Decimal("5000")
        assert "5010" not in balances  # COGS posted Jan 20, excluded

    def test_filter_by_entity(self, populated_ledger: Ledger) -> None:
        am_entries = populated_ledger.filter_by_entity("AM")
        assert len(am_entries) == 1
        assert am_entries[0].description == "AM Sale"

    def test_filter_by_date_range(self, populated_ledger: Ledger) -> None:
        jan = populated_ledger.filter_by_date_range(
            datetime.date(2025, 1, 1),
            datetime.date(2025, 1, 31),
        )
        assert len(jan) == 2  # Both PC entries

    def test_account_detail(self, populated_ledger: Ledger) -> None:
        detail = populated_ledger.account_detail("1010")
        assert len(detail) == 2  # PC cash + AM cash

    def test_account_detail_filtered(self, populated_ledger: Ledger) -> None:
        detail = populated_ledger.account_detail("1010", entity_code="AM")
        assert len(detail) == 1
        assert detail[0][1].debit == Decimal("3000")
