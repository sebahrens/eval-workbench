"""Scenario draft service for TUI state management.

Non-UI service that maintains a mutable draft override, applies typed
field updates, validates via the config loader, computes diffs against
the base, and performs atomic YAML saves.

Bead: synth-data-2u6.6.3
"""

from __future__ import annotations

import copy
import tempfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml

from generator.config import (
    Config,
    ConfigError,
    deep_merge,
    load_layered_config,
)
from generator.schema_metadata import (
    FieldMeta,
    get_field_by_path,
    is_unsupported_path,
)

# ---------------------------------------------------------------------------
# Types
# ---------------------------------------------------------------------------


@dataclass
class FieldDiff:
    """A single field difference between base and draft."""

    path: str
    base_value: Any
    draft_value: Any


@dataclass
class ValidationResult:
    """Result of validating the current draft state."""

    valid: bool
    errors: list[str] = field(default_factory=list)
    config: Config | None = None


# ---------------------------------------------------------------------------
# Draft Service
# ---------------------------------------------------------------------------


class DraftService:
    """Manages a mutable override draft on top of an immutable base config.

    The service maintains:
    - ``base_raw``: the merged raw dict from base + file overlays (immutable)
    - ``draft``: the override-only dict (mutable, starts empty)

    The effective config at any point is ``deep_merge(base_raw, draft)``.
    """

    def __init__(
        self,
        base_path: str | Path,
        overlays: list[str | Path] | None = None,
    ) -> None:
        self._base_path = Path(base_path)
        self._overlays = [Path(p) for p in (overlays or [])]
        self._base_raw: dict = {}
        self._draft: dict = {}
        self._load()

    def _load(self) -> None:
        """Load the base config and file overlays into base_raw."""
        from generator.config import _load_yaml

        self._base_raw = _load_yaml(self._base_path)
        for overlay_path in self._overlays:
            overlay_raw = _load_yaml(overlay_path)
            self._base_raw = deep_merge(self._base_raw, overlay_raw)
        self._draft = {}

    # ------------------------------------------------------------------
    # Read accessors
    # ------------------------------------------------------------------

    @property
    def base_raw(self) -> dict:
        """The immutable base config (merged with file overlays)."""
        return copy.deepcopy(self._base_raw)

    @property
    def draft(self) -> dict:
        """The current override-only draft dict."""
        return copy.deepcopy(self._draft)

    @property
    def merged_raw(self) -> dict:
        """The effective merged config (base + draft)."""
        return deep_merge(self._base_raw, self._draft)

    def get_field_value(self, path: str) -> Any:
        """Get the effective value at a dotted path from the merged config."""
        merged = self.merged_raw
        return _get_nested(merged, path)

    def get_base_value(self, path: str) -> Any:
        """Get the base value at a dotted path (before draft overrides)."""
        return _get_nested(self._base_raw, path)

    def is_modified(self, path: str) -> bool:
        """Return True if the draft has an override for the given path."""
        return _has_nested(self._draft, path)

    # ------------------------------------------------------------------
    # Edit operations
    # ------------------------------------------------------------------

    def set_field(self, path: str, value: Any) -> None:
        """Apply a typed field update to the draft.

        Raises ValueError if the path is unsupported or the value fails
        basic type checking against schema metadata.
        """
        if is_unsupported_path(path):
            raise ValueError(
                f"Field '{path}' is not user-editable in v1"
            )

        meta = get_field_by_path(path)
        if meta is not None:
            value = _coerce_value(meta, value)

        _set_nested(self._draft, path, value)

    def reset_field(self, path: str) -> None:
        """Remove an override for a field, reverting to the base value."""
        _delete_nested(self._draft, path)

    def reset_all(self) -> None:
        """Clear all draft overrides."""
        self._draft = {}

    # ------------------------------------------------------------------
    # Validation
    # ------------------------------------------------------------------

    def validate(self) -> ValidationResult:
        """Validate the merged config (base + draft) through the config loader.

        Returns a ValidationResult with the parsed Config on success,
        or a list of error messages on failure.
        """
        merged = self.merged_raw
        # Write to a temp file and load through the standard validator
        with tempfile.NamedTemporaryFile(
            mode="w", suffix=".yaml", delete=False
        ) as f:
            yaml.dump(merged, f, default_flow_style=False, sort_keys=True)
            tmp_path = Path(f.name)

        try:
            config = load_layered_config(tmp_path)
            return ValidationResult(valid=True, config=config)
        except ConfigError as exc:
            return ValidationResult(valid=False, errors=[str(exc)])
        finally:
            tmp_path.unlink(missing_ok=True)

    # ------------------------------------------------------------------
    # Diff
    # ------------------------------------------------------------------

    def diff(self) -> list[FieldDiff]:
        """Compute field-level differences between base and the draft.

        Only returns diffs for fields that are explicitly overridden
        in the draft (not computed differences from deep_merge semantics).
        """
        diffs: list[FieldDiff] = []
        flat_draft = _flatten_to_leaves(self._draft)
        for path, draft_value in sorted(flat_draft.items()):
            base_value = _get_nested(self._base_raw, path)
            if base_value != draft_value:
                diffs.append(FieldDiff(
                    path=path,
                    base_value=base_value,
                    draft_value=draft_value,
                ))
        return diffs

    # ------------------------------------------------------------------
    # Save
    # ------------------------------------------------------------------

    def save(
        self,
        output_path: str | Path,
        *,
        force: bool = False,
        validate: bool = True,
    ) -> None:
        """Atomically save the draft override to a YAML file.

        Only the override (delta) is saved, not the full merged config.
        The base config.yaml is never modified.

        Uses write-to-temp + rename for atomicity.

        Parameters
        ----------
        output_path : str | Path
            Destination for the override YAML.
        force : bool
            If False (default) and *output_path* already exists, raises
            ``FileExistsError`` instead of silently overwriting.
        validate : bool
            If True (default), validates the merged config before saving.
            Raises ``ConfigError`` if the merged config is invalid.
        """
        output_path = Path(output_path)

        # Guard: validate merged config before writing anything
        if validate:
            result = self.validate()
            if not result.valid:
                raise ConfigError(
                    "Cannot save: merged config is invalid. "
                    + "; ".join(result.errors)
                )

        # Guard: refuse to overwrite without force
        if not force and output_path.exists():
            raise FileExistsError(
                f"Override file already exists: {output_path}. "
                "Pass force=True to overwrite."
            )

        output_path.parent.mkdir(parents=True, exist_ok=True)

        # Write to a temp file in the same directory, then rename
        tmp_fd = tempfile.NamedTemporaryFile(
            mode="w",
            suffix=".yaml",
            dir=output_path.parent,
            delete=False,
        )
        try:
            yaml.dump(
                self._draft,
                tmp_fd,
                default_flow_style=False,
                sort_keys=True,
            )
            tmp_fd.close()
            Path(tmp_fd.name).replace(output_path)
        except Exception:
            Path(tmp_fd.name).unlink(missing_ok=True)
            raise

    def load_draft(self, draft_path: str | Path) -> None:
        """Load an existing override file into the draft state."""
        draft_path = Path(draft_path)
        if not draft_path.exists():
            raise ConfigError(f"Draft file not found: {draft_path}")
        with open(draft_path) as f:
            raw = yaml.safe_load(f)
        if raw is None:
            self._draft = {}
        elif not isinstance(raw, dict):
            raise ConfigError(
                f"Draft file root must be a mapping, got {type(raw).__name__}"
            )
        else:
            self._draft = raw


# ---------------------------------------------------------------------------
# Internal helpers
# ---------------------------------------------------------------------------


def _get_nested(d: dict, path: str) -> Any:
    """Get a value from a nested dict using a dotted path."""
    parts = path.split(".")
    current: Any = d
    for part in parts:
        if not isinstance(current, dict):
            return None
        current = current.get(part)
        if current is None:
            return None
    return current


def _has_nested(d: dict, path: str) -> bool:
    """Return True if a dotted path exists in a nested dict."""
    parts = path.split(".")
    current: Any = d
    for part in parts:
        if not isinstance(current, dict) or part not in current:
            return False
        current = current[part]
    return True


def _set_nested(d: dict, path: str, value: Any) -> None:
    """Set a value in a nested dict at a dotted path, creating intermediates."""
    parts = path.split(".")
    current = d
    for part in parts[:-1]:
        if part not in current or not isinstance(current[part], dict):
            current[part] = {}
        current = current[part]
    current[parts[-1]] = value


def _delete_nested(d: dict, path: str) -> None:
    """Delete a value from a nested dict at a dotted path.

    Cleans up empty parent dicts after deletion.
    """
    parts = path.split(".")
    # Walk down to the parent, tracking the path for cleanup
    parents: list[tuple[dict, str]] = []
    current = d
    for part in parts[:-1]:
        if not isinstance(current, dict) or part not in current:
            return  # Path doesn't exist, nothing to delete
        parents.append((current, part))
        current = current[part]

    if isinstance(current, dict) and parts[-1] in current:
        del current[parts[-1]]

    # Clean up empty parents bottom-up
    for parent_dict, key in reversed(parents):
        if isinstance(parent_dict[key], dict) and not parent_dict[key]:
            del parent_dict[key]


def _flatten_to_leaves(d: dict, prefix: str = "") -> dict[str, Any]:
    """Flatten a nested dict to dotted-path keys mapping to leaf values."""
    result: dict[str, Any] = {}
    for k, v in d.items():
        full = f"{prefix}.{k}" if prefix else k
        if isinstance(v, dict):
            result.update(_flatten_to_leaves(v, full))
        else:
            result[full] = v
    return result


def _coerce_value(meta: FieldMeta, value: Any) -> Any:
    """Coerce and validate a value against field metadata.

    Raises ValueError on type mismatch or constraint violation.
    """
    from generator.schema_metadata import InputType

    if value is None:
        return None

    try:
        if meta.input_type == InputType.INTEGER:
            value = int(value)
        elif meta.input_type == InputType.FLOAT:
            value = float(value)
        elif meta.input_type == InputType.TEXT:
            value = str(value)
        elif meta.input_type == InputType.CHOICE:
            value = str(value)
            if meta.choices and value not in meta.choices:
                raise ValueError(
                    f"Invalid choice for '{meta.path}': {value!r}. "
                    f"Allowed: {meta.choices}"
                )
        elif meta.input_type == InputType.MULTI_CHOICE:
            if not isinstance(value, list):
                raise ValueError(
                    f"Field '{meta.path}' expects a list, got {type(value).__name__}"
                )
            if meta.choices:
                invalid = [v for v in value if v not in meta.choices]
                if invalid:
                    raise ValueError(
                        f"Invalid choices for '{meta.path}': {invalid}. "
                        f"Allowed: {meta.choices}"
                    )
        elif meta.input_type in (InputType.LIST_TEXT, InputType.LIST_INT):
            if not isinstance(value, list):
                raise ValueError(
                    f"Field '{meta.path}' expects a list, got {type(value).__name__}"
                )
    except (TypeError, ValueError) as exc:
        if "Invalid choice" in str(exc) or "expects a list" in str(exc):
            raise
        raise ValueError(
            f"Cannot coerce value for '{meta.path}': {exc}"
        ) from exc

    # Range validation
    if meta.range_min is not None and isinstance(value, (int, float)):
        if value < meta.range_min:
            raise ValueError(
                f"Field '{meta.path}' value {value} is below minimum {meta.range_min}"
            )
    if meta.range_max is not None and isinstance(value, (int, float)):
        if value > meta.range_max:
            raise ValueError(
                f"Field '{meta.path}' value {value} is above maximum {meta.range_max}"
            )

    return value
