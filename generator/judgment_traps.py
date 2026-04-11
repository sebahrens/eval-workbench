"""Judgment trap registry for non-numeric professional judgment issues.

Judgment traps differ from planted errors: planted errors are numeric
discrepancies an agent must detect; judgment traps are qualitative
reasoning challenges where the agent must demonstrate professional
judgment (flagging contradictions, noting missing evidence, etc.).

Trap types:
- summary_contradiction: summary text conflicts with source data
- missing_evidence: claim lacks supporting documentation
- stale_document: document date predates the analysis period
- scope_boundary: issue falls outside the engagement scope
- overconfident_conclusion: conclusion stated without appropriate caveats

Expected responses:
- flag: agent should flag the issue for review
- caveat: agent should include a caveat or qualification
- deprioritize: agent should note but deprioritize
- do_not_assert: agent should not make an assertion without more evidence
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from pathlib import Path

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class JudgmentTrap:
    """One entry in the judgment trap registry."""

    trap_id: str                       # e.g. "JT-001"
    test_case: str                     # e.g. "TC-19"
    trap_type: str                     # One of _VALID_TRAP_TYPES
    description: str                   # Human-readable description
    source_refs: tuple[str, ...]       # Document/section references
    expected_response: str             # One of _VALID_RESPONSES
    rationale: str                     # Why this response is expected


_VALID_TRAP_TYPES = frozenset({
    "summary_contradiction",
    "missing_evidence",
    "stale_document",
    "scope_boundary",
    "overconfident_conclusion",
})

_VALID_RESPONSES = frozenset({
    "flag",
    "caveat",
    "deprioritize",
    "do_not_assert",
})


# ---------------------------------------------------------------------------
# Registry
# ---------------------------------------------------------------------------

@dataclass
class JudgmentTrapRegistry:
    """Complete registry of judgment traps for the test suite."""

    entries: dict[str, JudgmentTrap] = field(default_factory=dict)

    def add(self, trap: JudgmentTrap) -> None:
        if trap.trap_type not in _VALID_TRAP_TYPES:
            raise ValueError(
                f"Invalid trap type {trap.trap_type!r}. "
                f"Must be one of: {sorted(_VALID_TRAP_TYPES)}"
            )
        if trap.expected_response not in _VALID_RESPONSES:
            raise ValueError(
                f"Invalid expected_response {trap.expected_response!r}. "
                f"Must be one of: {sorted(_VALID_RESPONSES)}"
            )
        if trap.trap_id in self.entries:
            raise ValueError(f"Duplicate trap_id: {trap.trap_id}")
        self.entries[trap.trap_id] = trap

    def get(self, trap_id: str) -> JudgmentTrap:
        return self.entries[trap_id]

    def by_test_case(self, tc: str) -> list[JudgmentTrap]:
        """Return all traps for *tc*, sorted by trap_id."""
        return sorted(
            (t for t in self.entries.values() if t.test_case == tc),
            key=lambda t: t.trap_id,
        )

    def by_type(self, trap_type: str) -> list[JudgmentTrap]:
        """Return all traps of *trap_type*, sorted by trap_id."""
        return sorted(
            (t for t in self.entries.values() if t.trap_type == trap_type),
            key=lambda t: t.trap_id,
        )

    def by_response(self, expected_response: str) -> list[JudgmentTrap]:
        """Return all traps with *expected_response*, sorted by trap_id."""
        return sorted(
            (t for t in self.entries.values()
             if t.expected_response == expected_response),
            key=lambda t: t.trap_id,
        )

    # -- serialisation --------------------------------------------------------

    def to_dict(self) -> list[dict]:
        """Return a sorted list of entry dicts (deterministic order)."""
        result = []
        for k in sorted(self.entries):
            d = asdict(self.entries[k])
            # Convert tuple to list for JSON compatibility
            d["source_refs"] = list(d["source_refs"])
            result.append(d)
        return result

    def write_json(self, path: str | Path) -> None:
        """Write the registry to *path* as formatted JSON."""
        path = Path(path)
        path.parent.mkdir(parents=True, exist_ok=True)
        with open(path, "w") as f:
            json.dump(self.to_dict(), f, indent=2, sort_keys=True)
            f.write("\n")
