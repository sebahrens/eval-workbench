"""Scenario pack registry — v1 static registry.

Each pack declares its test cases, canary file keys, and ordered emitters.
The orchestrator resolves selected packs, validates dependencies, and runs
emitters in topological order.
"""

from __future__ import annotations

from collections.abc import Callable
from dataclasses import dataclass, field

from generator.config import ConfigError


@dataclass(frozen=True)
class ScenarioPack:
    """A self-contained scenario pack definition."""

    pack_id: str
    display_name: str
    test_cases: list[str]
    canary_file_keys: list[str]
    emitters: list[Callable[..., None]]
    dependencies: list[str] = field(default_factory=list)


# Ordered dict — insertion order = default execution order.
REGISTRY: dict[str, ScenarioPack] = {}

# Default pack IDs when no --packs flag is provided.
_DEFAULT_PACKS: list[str] = ["cascade_accounting_core"]


def register_pack(pack: ScenarioPack) -> None:
    """Register a pack in the global registry."""
    if pack.pack_id in REGISTRY:
        raise ConfigError(f"duplicate pack_id: {pack.pack_id!r}")
    REGISTRY[pack.pack_id] = pack


def get_pack(pack_id: str) -> ScenarioPack:
    """Look up a pack by ID or raise ConfigError."""
    try:
        return REGISTRY[pack_id]
    except KeyError:
        valid = ", ".join(sorted(REGISTRY))
        raise ConfigError(
            f"unknown pack {pack_id!r}; valid packs: {valid}"
        ) from None


def list_packs() -> list[ScenarioPack]:
    """Return all registered packs in insertion order."""
    return list(REGISTRY.values())


def resolve_packs(pack_ids: list[str] | None) -> list[ScenarioPack]:
    """Resolve a list of pack IDs into validated, dependency-sorted packs.

    If *pack_ids* is None, returns the default pack list.
    The special value ``["all"]`` selects every registered pack.
    """
    if pack_ids is None:
        pack_ids = list(_DEFAULT_PACKS)
    elif pack_ids == ["all"]:
        pack_ids = list(REGISTRY)

    selected = {pid: get_pack(pid) for pid in pack_ids}

    # Validate dependencies
    for pid, pack in selected.items():
        for dep in pack.dependencies:
            if dep not in selected:
                raise ConfigError(
                    f"pack {pid!r} depends on {dep!r} which is not selected"
                )

    # Topological sort (Kahn's algorithm)
    in_degree: dict[str, int] = {pid: 0 for pid in selected}
    for pid, pack in selected.items():
        for dep in pack.dependencies:
            in_degree[pid] += 1

    queue = sorted(pid for pid, deg in in_degree.items() if deg == 0)
    order: list[str] = []
    while queue:
        pid = queue.pop(0)
        order.append(pid)
        for other_pid, other_pack in selected.items():
            if pid in other_pack.dependencies:
                in_degree[other_pid] -= 1
                if in_degree[other_pid] == 0:
                    queue.append(other_pid)
                    queue.sort()

    if len(order) != len(selected):
        raise ConfigError("circular dependency among selected packs")

    return [selected[pid] for pid in order]


def collect_canary_keys(packs: list[ScenarioPack]) -> list[str]:
    """Merge and sort canary file keys from all selected packs."""
    keys: list[str] = []
    for pack in packs:
        keys.extend(pack.canary_file_keys)
    return sorted(keys)


def collect_test_case_count(packs: list[ScenarioPack]) -> int:
    """Total number of test cases across selected packs."""
    return sum(len(p.test_cases) for p in packs)


# ── Auto-register built-in packs on import ──────────────────────────
from generator.packs import accounting_core as _ac  # noqa: E402

register_pack(_ac.PACK)
