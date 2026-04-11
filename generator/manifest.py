"""Manifest emitter — tracks every file produced by the generator.

Provides a Manifest context manager that formatters register files against
during generation. At the end of the run, writes manifest.json with
deterministic ordering.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class ManifestEntry:
    """One entry in the manifest."""

    path: str                     # Relative path from output root
    type: str                     # File type: "xlsx", "docx", "pdf", "csv", "json", "yaml", etc.
    size: int = 0                 # File size in bytes (populated after write)
    canary: str = ""              # 8-char canary code, if applicable
    test_cases: list[str] = field(default_factory=list)  # Which TCs use this file
    scenario_pack: str = ""       # Pack ID that produced this file (optional)
    augmentation_source: str | None = None  # None=no augmentation, "cache:<hash>", "baseline"


class Manifest:
    """Accumulates file entries during generation, writes manifest.json on exit.

    Usage::

        with Manifest(output_dir) as manifest:
            manifest.register("shared_data/coa.xlsx", "xlsx",
                              canary="AB12CD34", test_cases=["TC-01", "TC-03"])
            # ... formatters write files ...
        # manifest.json is written automatically on __exit__
    """

    def __init__(self, output_dir: Path) -> None:
        self._output_dir = output_dir
        self._entries: dict[str, ManifestEntry] = {}
        self._current_pack: str = ""

    # -- context manager ------------------------------------------------------

    def __enter__(self) -> Manifest:
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:  # type: ignore[no-untyped-def]
        # Only write if no exception — don't produce a partial manifest
        if exc_type is None:
            self._resolve_sizes()
            self.write_json()
        return None

    # -- pack context ---------------------------------------------------------

    def set_current_pack(self, pack_id: str) -> None:
        """Set the scenario pack ID for subsequent register() calls."""
        self._current_pack = pack_id

    # -- registration ---------------------------------------------------------

    def register(
        self,
        path: str,
        file_type: str,
        *,
        canary: str = "",
        test_cases: list[str] | None = None,
        scenario_pack: str = "",
    ) -> None:
        """Register a file that has been (or will be) written.

        Parameters
        ----------
        path:
            Path relative to the output root (e.g. "shared_data/coa.xlsx").
        file_type:
            File extension / type label.
        canary:
            Canary code embedded in this file, if any.
        test_cases:
            List of test case IDs that use this file (e.g. ["TC-01"]).
        scenario_pack:
            Pack ID that produced this file. If empty, uses the current
            pack context set via :meth:`set_current_pack`.
        """
        self._entries[path] = ManifestEntry(
            path=path,
            type=file_type,
            canary=canary,
            test_cases=sorted(test_cases) if test_cases else [],
            scenario_pack=scenario_pack or self._current_pack,
        )

    @property
    def entries(self) -> dict[str, ManifestEntry]:
        """Read-only access to registered entries."""
        return self._entries

    # -- size resolution ------------------------------------------------------

    def _resolve_sizes(self) -> None:
        """Populate file sizes from the actual files on disk."""
        for entry in self._entries.values():
            full_path = self._output_dir / entry.path
            if full_path.exists():
                entry.size = full_path.stat().st_size

    # -- serialisation --------------------------------------------------------

    def to_dict(self) -> list[dict]:
        """Return a sorted list of entry dicts (deterministic key order)."""
        result = []
        for key in sorted(self._entries):
            e = self._entries[key]
            d: dict = {
                "path": e.path,
                "type": e.type,
                "size": e.size,
                "canary": e.canary,
                "test_cases": e.test_cases,
            }
            if e.scenario_pack:
                d["scenario_pack"] = e.scenario_pack
            if e.augmentation_source is not None:
                d["augmentation_source"] = e.augmentation_source
            result.append(d)
        return result

    def write_json(self, filename: str = "manifest.json") -> Path:
        """Write manifest.json to the output directory. Returns the path."""
        out_path = self._output_dir / filename
        out_path.parent.mkdir(parents=True, exist_ok=True)
        with open(out_path, "w") as f:
            json.dump(self.to_dict(), f, indent=2, sort_keys=False)
            f.write("\n")
        return out_path
