"""Validators for augmentation outputs against canonical model facts.

Ensures that LLM-generated narrative text:
1. Includes required canonical names, dates, and amounts.
2. Does not contradict any canonical model facts.
3. Preserves canary values (if the augmented text replaces canary-bearing content).
4. Does not mask planted errors.

Validators return field-level messages and are usable by DataDesigner
or any local augmentation pipeline.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from decimal import Decimal
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from generator.config import Config


# ---------------------------------------------------------------------------
# Result types
# ---------------------------------------------------------------------------

@dataclass
class ValidationMessage:
    """A single validation finding, scoped to a specific field or fact."""

    level: str  # "error" or "warning"
    field: str  # canonical fact identifier (e.g. "company.name", "subsidiary.PC.revenue")
    message: str

    def __str__(self) -> str:
        return f"[{self.level.upper()}] {self.field}: {self.message}"


@dataclass
class ValidationResult:
    """Aggregated result from one or more validators."""

    messages: list[ValidationMessage] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not any(m.level == "error" for m in self.messages)

    @property
    def errors(self) -> list[ValidationMessage]:
        return [m for m in self.messages if m.level == "error"]

    @property
    def warnings(self) -> list[ValidationMessage]:
        return [m for m in self.messages if m.level == "warning"]

    def merge(self, other: ValidationResult) -> None:
        """Merge another result's messages into this one."""
        self.messages.extend(other.messages)


# ---------------------------------------------------------------------------
# Canonical fact extraction
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class CanonicalFact:
    """A single fact that must appear in (or not be contradicted by) augmented text."""

    field: str  # dot-path identifier
    value: str  # canonical string representation
    required: bool = True  # must appear in output (vs. must-not-contradict)


def extract_canonical_facts(config: Config) -> list[CanonicalFact]:
    """Extract the set of canonical facts from a Config that augmented
    narrative text must preserve.

    Returns facts for: company identity, subsidiary names, fiscal year,
    and key financial figures.
    """
    facts: list[CanonicalFact] = []
    co = config.company

    # Company identity
    facts.append(CanonicalFact("company.name", co.name))
    facts.append(CanonicalFact("company.headquarters", co.headquarters))
    facts.append(CanonicalFact("company.industry", co.industry))
    facts.append(CanonicalFact("company.fiscal_year_end", co.fiscal_year_end))

    # Subsidiary names — must appear when referenced
    for key, sub in sorted(co.subsidiaries.items()):
        facts.append(CanonicalFact(
            f"subsidiary.{sub.entity_code}.legal_name",
            sub.legal_name,
            required=False,  # only required if the subsidiary is referenced
        ))
        facts.append(CanonicalFact(
            f"subsidiary.{sub.entity_code}.entity_code",
            sub.entity_code,
            required=False,
        ))

    # Financial figures — must not be contradicted
    facts.append(CanonicalFact(
        "company.consolidated_revenue",
        str(co.consolidated_revenue),
        required=False,
    ))
    for key, sub in sorted(co.subsidiaries.items()):
        facts.append(CanonicalFact(
            f"subsidiary.{sub.entity_code}.revenue",
            str(sub.revenue),
            required=False,
        ))

    # Fiscal years
    for year in co.years:
        facts.append(CanonicalFact(
            f"company.year.{year}",
            str(year),
            required=False,
        ))

    return facts


# ---------------------------------------------------------------------------
# Validators
# ---------------------------------------------------------------------------

def validate_required_facts(
    text: str,
    facts: list[CanonicalFact],
) -> ValidationResult:
    """Check that all required canonical facts appear in the text.

    A fact is considered present if its value appears as a substring
    (case-insensitive for names, exact for codes and numbers).
    """
    result = ValidationResult()
    for fact in facts:
        if not fact.required:
            continue
        if not _fact_present(text, fact):
            result.messages.append(ValidationMessage(
                level="error",
                field=fact.field,
                message=f"Required canonical fact missing: {fact.value!r}",
            ))
    return result


def validate_no_contradictions(
    text: str,
    facts: list[CanonicalFact],
) -> ValidationResult:
    """Check that financial figures in the text don't contradict canonical values.

    Scans for dollar-amount patterns near fact keywords and flags
    contradictions where the text states a different number for a
    canonical quantity.
    """
    result = ValidationResult()

    for fact in facts:
        if not _is_financial_fact(fact.field):
            continue

        canonical_amount = _parse_amount(fact.value)
        if canonical_amount is None:
            continue

        # Look for the entity/company name near a dollar figure
        context_label = _fact_context_label(fact.field)
        contradictions = _find_contradicting_amounts(
            text, context_label, canonical_amount,
        )
        for found_amount in contradictions:
            result.messages.append(ValidationMessage(
                level="error",
                field=fact.field,
                message=(
                    f"Contradicting amount found: text states "
                    f"${found_amount:,.0f} but canonical value is "
                    f"${canonical_amount:,.0f}"
                ),
            ))

    return result


def validate_canary_preserved(
    text: str,
    canary: str,
    file_id: str,
) -> ValidationResult:
    """Check that a canary value is still present in augmented text."""
    result = ValidationResult()
    if canary and canary not in text:
        result.messages.append(ValidationMessage(
            level="error",
            field=f"canary.{file_id}",
            message=f"Canary value {canary!r} missing from augmented output",
        ))
    return result


def validate_planted_error_preserved(
    text: str,
    error_id: str,
    error_marker: str,
) -> ValidationResult:
    """Check that a planted error marker is still detectable in augmented text."""
    result = ValidationResult()
    if error_marker and error_marker not in text:
        result.messages.append(ValidationMessage(
            level="error",
            field=f"planted_error.{error_id}",
            message=f"Planted error {error_id!r} marker no longer detectable",
        ))
    return result


def validate_augmentation_output(
    text: str,
    config: Config,
    *,
    canary: str = "",
    file_id: str = "",
    error_id: str = "",
    error_marker: str = "",
) -> ValidationResult:
    """Run all validators on an augmented text output.

    This is the main entry point for the augmentation pipeline.
    """
    facts = extract_canonical_facts(config)
    result = ValidationResult()

    result.merge(validate_required_facts(text, facts))
    result.merge(validate_no_contradictions(text, facts))

    if canary:
        result.merge(validate_canary_preserved(text, canary, file_id))

    if error_id:
        result.merge(validate_planted_error_preserved(
            text, error_id, error_marker,
        ))

    return result


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------

def _fact_present(text: str, fact: CanonicalFact) -> bool:
    """Check if a fact's value appears in the text.

    Names are matched case-insensitively. Codes and numbers are exact.
    """
    if _is_financial_fact(fact.field) or fact.field.endswith(".entity_code"):
        return fact.value in text
    return fact.value.lower() in text.lower()


def _is_financial_fact(field_path: str) -> bool:
    return field_path.endswith(".revenue") or field_path.endswith(".consolidated_revenue")


def _fact_context_label(field_path: str) -> str:
    """Extract a human-readable context label for proximity matching.

    e.g. "subsidiary.PC.revenue" -> "PC"
         "company.consolidated_revenue" -> "revenue"
    """
    parts = field_path.split(".")
    if len(parts) >= 3 and parts[0] == "subsidiary":
        return parts[1]  # entity code
    # For company-level facts, use "revenue" as the context keyword
    last = parts[-1]
    if "revenue" in last.lower():
        return "revenue"
    return last


def _parse_amount(value_str: str) -> int | None:
    """Parse a numeric string into an integer amount, or None if not numeric."""
    cleaned = value_str.replace(",", "").replace("$", "").strip()
    try:
        return int(Decimal(cleaned))
    except Exception:
        return None


# Matches dollar amounts like $200,000,000 or 200000000 or $200M
_DOLLAR_PATTERN = re.compile(
    r"\$?\s*(\d[\d,]*(?:\.\d+)?)\s*(?:million|M|billion|B)?",
    re.IGNORECASE,
)


def _normalize_amount(match_str: str, suffix: str) -> int | None:
    """Normalize a matched amount string to an integer."""
    cleaned = match_str.replace(",", "")
    try:
        value = float(cleaned)
    except ValueError:
        return None

    suffix_lower = suffix.strip().lower() if suffix else ""
    if suffix_lower in ("million", "m"):
        value *= 1_000_000
    elif suffix_lower in ("billion", "b"):
        value *= 1_000_000_000
    return int(value)


_AMOUNT_WITH_SUFFIX = re.compile(
    r"\$?\s*(\d[\d,]*(?:\.\d+)?)\s*(million|M|billion|B)?",
    re.IGNORECASE,
)


def _find_contradicting_amounts(
    text: str,
    context_label: str,
    canonical_amount: int,
) -> list[int]:
    """Find dollar amounts near a context label that differ from the canonical value.

    Uses a proximity window of 200 characters around the context label.
    """
    contradictions: list[int] = []

    # Find all occurrences of the context label
    label_lower = context_label.lower()
    text_lower = text.lower()
    pos = 0
    while True:
        idx = text_lower.find(label_lower, pos)
        if idx == -1:
            break

        # Extract a window around the label
        window_start = max(0, idx - 200)
        window_end = min(len(text), idx + len(context_label) + 200)
        window = text[window_start:window_end]

        for m in _AMOUNT_WITH_SUFFIX.finditer(window):
            amount = _normalize_amount(m.group(1), m.group(2) or "")
            if amount is None:
                continue
            # Only flag if the amount is in a similar order of magnitude
            # but different from canonical — avoid flagging unrelated numbers
            if amount == canonical_amount:
                continue
            if _same_order_of_magnitude(amount, canonical_amount):
                contradictions.append(amount)

        pos = idx + 1

    return contradictions


def _same_order_of_magnitude(a: int, b: int) -> bool:
    """True if a and b are within one order of magnitude of each other."""
    if a == 0 or b == 0:
        return False
    ratio = max(a, b) / min(a, b)
    return ratio < 10
