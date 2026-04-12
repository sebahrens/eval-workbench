"""Validation error panel with field-path navigation.

Displays validation errors grouped by config section.  Each error row
is clickable — clicking posts a ``ValidationPanel.GoToField`` message
that parent screens can handle to scroll to the offending widget.

Bead: synth-data-2u6.6.8
"""

from __future__ import annotations

import re
from dataclasses import dataclass

from textual.app import ComposeResult
from textual.containers import VerticalScroll
from textual.message import Message
from textual.reactive import reactive
from textual.widget import Widget
from textual.widgets import Label, Static

from generator.tui.draft_service import ValidationResult

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

# Patterns that commonly appear in ConfigError messages to identify field paths.
# Examples:
#   "difficulty.error_density must be 0.0–1.0, got 2.0"
#   "Missing required keys in company: ['name']"
#   "Unknown keys in company.subsidiaries.US: ['bogus']"
_PATH_PATTERN = re.compile(
    r"(?:^|\b)"
    r"((?:company|difficulty|output|errors|augmentation)"
    r"(?:\.[a-z_][a-z0-9_.]*)?)"
    r"(?:\b|$)",
    re.IGNORECASE,
)

_SECTION_LABELS: dict[str, str] = {
    "company": "Company",
    "difficulty": "Difficulty",
    "output": "Output",
    "errors": "Errors",
    "augmentation": "Augmentation",
}


@dataclass
class GroupedError:
    """A single validation error with extracted section and field path."""

    message: str
    section: str = "general"
    field_path: str = ""


def _extract_field_path(message: str) -> str:
    """Try to extract a dotted field path from an error message."""
    m = _PATH_PATTERN.search(message)
    return m.group(1) if m else ""


def _group_errors(errors: list[str]) -> dict[str, list[GroupedError]]:
    """Group flat error strings by config section."""
    groups: dict[str, list[GroupedError]] = {}
    for msg in errors:
        path = _extract_field_path(msg)
        if path:
            section = path.split(".")[0]
        else:
            section = "general"
        ge = GroupedError(message=msg, section=section, field_path=path)
        groups.setdefault(section, []).append(ge)
    return groups


# ---------------------------------------------------------------------------
# Widget
# ---------------------------------------------------------------------------


class ValidationPanel(Widget):
    """Collapsible panel that shows validation errors grouped by section.

    Set ``result`` to update the display.  When there are no errors
    the panel hides itself automatically.
    """

    DEFAULT_CSS = """
    ValidationPanel {
        height: auto;
        max-height: 14;
        display: none;
        margin: 0 1;
        border: solid $error;
        padding: 0 1;
    }
    ValidationPanel.has-errors {
        display: block;
    }
    ValidationPanel .vp-title {
        color: $error;
        text-style: bold;
        padding: 0;
    }
    ValidationPanel .vp-section-heading {
        color: $warning;
        text-style: bold;
        margin-top: 1;
        padding: 0;
    }
    ValidationPanel .vp-error-row {
        padding: 0 0 0 2;
    }
    ValidationPanel .vp-error-row:hover {
        background: $surface-lighten-1;
    }
    ValidationPanel .vp-error-clickable {
        padding: 0 0 0 2;
        color: $text;
    }
    ValidationPanel .vp-error-clickable:hover {
        background: $surface-lighten-1;
        text-style: underline;
    }
    ValidationPanel .vp-valid {
        color: $success;
        padding: 0;
    }
    """

    # -- Messages ----------------------------------------------------------

    class GoToField(Message):
        """Posted when the user clicks an error row with a known field path."""

        def __init__(self, field_path: str) -> None:
            super().__init__()
            self.field_path = field_path

    # -- Reactive state ----------------------------------------------------

    result: reactive[ValidationResult | None] = reactive(None)

    def watch_result(self, result: ValidationResult | None) -> None:
        """Re-render whenever the validation result changes."""
        self._rebuild()

    # -- Public API --------------------------------------------------------

    @property
    def has_errors(self) -> bool:
        """True when the current result contains validation errors."""
        r = self.result
        return r is not None and not r.valid

    def clear(self) -> None:
        """Remove any displayed errors."""
        self.result = None

    # -- Compose / rebuild -------------------------------------------------

    def compose(self) -> ComposeResult:
        with VerticalScroll(id="vp-scroll"):
            yield Static("", id="vp-content")

    def _rebuild(self) -> None:
        """Rebuild the panel content from the current result."""
        result = self.result
        # Toggle visibility
        if result is None or result.valid:
            self.remove_class("has-errors")
            return

        self.add_class("has-errors")

        # Remove old dynamic children
        try:
            scroll = self.query_one("#vp-scroll", VerticalScroll)
        except Exception:
            return

        scroll.remove_children()

        # Title
        n = len(result.errors)
        scroll.mount(
            Static(
                f"[bold red]Validation failed — {n} error{'s' if n != 1 else ''}[/bold red]",
                classes="vp-title",
            )
        )

        groups = _group_errors(result.errors)

        # Render each section group
        section_order = ["company", "difficulty", "output", "errors", "augmentation", "general"]
        for section in section_order:
            if section not in groups:
                continue
            label = _SECTION_LABELS.get(section, section.title())
            scroll.mount(
                Static(f"[bold yellow]{label}[/bold yellow]", classes="vp-section-heading")
            )
            for ge in groups[section]:
                if ge.field_path:
                    row = Label(
                        f"[red]●[/red] {ge.message}",
                        classes="vp-error-clickable",
                        id=f"vperr-{ge.field_path.replace('.', '-')}",
                    )
                    row._vp_field_path = ge.field_path  # type: ignore[attr-defined]
                    scroll.mount(row)
                else:
                    scroll.mount(
                        Static(f"[red]●[/red] {ge.message}", classes="vp-error-row")
                    )

        # Render any sections not in the predefined order
        for section, items in groups.items():
            if section in section_order:
                continue
            label = _SECTION_LABELS.get(section, section.title())
            scroll.mount(
                Static(f"[bold yellow]{label}[/bold yellow]", classes="vp-section-heading")
            )
            for ge in items:
                scroll.mount(
                    Static(f"[red]●[/red] {ge.message}", classes="vp-error-row")
                )

    def on_click(self, event) -> None:
        """Bubble GoToField when an error label with a field path is clicked."""
        target = event.widget
        if target is None:
            return
        # Walk up to find a label with a vp_field_path attribute
        widget = target
        for _ in range(5):
            path = getattr(widget, "_vp_field_path", None)
            if path:
                self.post_message(self.GoToField(path))
                return
            widget = widget.parent
            if widget is None:
                break
