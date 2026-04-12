"""Merged config preview and diff screen.

Shows a deterministic diff from base config to the merged draft,
a raw YAML override preview, and a concise change summary.
This screen is read-only — it never mutates the draft.

Bead: synth-data-2u6.6.9
"""

from __future__ import annotations

import yaml
from textual.app import ComposeResult
from textual.containers import Vertical, VerticalScroll
from textual.screen import Screen
from textual.widgets import Footer, Header, Static

from generator.tui.draft_service import DraftService, FieldDiff

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_PREVIEW_HELP = (
    "[bold]Configuration Preview[/bold]\n"
    "Read-only view of your pending overrides.  The diff section shows "
    "each field you changed (base → draft).  The YAML section shows the "
    "raw override that would be saved to disk."
)

_NO_CHANGES = (
    "[dim]No overrides — the draft matches the base configuration.[/dim]"
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _format_value(value: object) -> str:
    """Format a value for display in the diff table."""
    if value is None:
        return "[dim](unset)[/dim]"
    if isinstance(value, list):
        if not value:
            return "[dim](empty list)[/dim]"
        return ", ".join(str(v) for v in value)
    return str(value)


def _build_diff_text(diffs: list[FieldDiff]) -> str:
    """Build a Rich-markup table of field diffs."""
    if not diffs:
        return _NO_CHANGES

    lines: list[str] = []
    lines.append("[bold underline]Changed Fields[/bold underline]\n")

    # Compute column widths for alignment
    path_width = max(len(d.path) for d in diffs)

    for d in diffs:
        base_str = _format_value(d.base_value)
        draft_str = _format_value(d.draft_value)
        lines.append(
            f"  [bold]{d.path:<{path_width}}[/bold]  "
            f"[red]{base_str}[/red] → [green]{draft_str}[/green]"
        )

    return "\n".join(lines)


def _build_summary(diffs: list[FieldDiff]) -> str:
    """Build a one-line change summary."""
    if not diffs:
        return "No changes."
    n = len(diffs)
    word = "field" if n == 1 else "fields"
    return f"{n} {word} modified."


def _build_yaml_preview(draft: dict) -> str:
    """Render the override dict as YAML for display."""
    if not draft:
        return "[dim]# (empty — no overrides)[/dim]"
    rendered = yaml.dump(draft, default_flow_style=False, sort_keys=True)
    return rendered.rstrip()


# ---------------------------------------------------------------------------
# Screen
# ---------------------------------------------------------------------------


class ConfigPreviewScreen(Screen):
    """Read-only screen showing the diff and YAML preview of draft overrides."""

    BINDINGS = [
        ("escape", "dismiss", "Back"),
        ("q", "dismiss", "Back"),
    ]

    def __init__(
        self,
        service: DraftService,
        name: str | None = None,
        id: str | None = None,
        classes: str | None = None,
    ) -> None:
        super().__init__(name=name, id=id, classes=classes)
        self.service = service

    def compose(self) -> ComposeResult:
        # Snapshot state once at compose time — no mutation
        diffs = self.service.diff()
        draft = self.service.draft
        summary = _build_summary(diffs)

        yield Header()
        with VerticalScroll():
            yield Static(_PREVIEW_HELP, classes="section-help")

            # Summary
            yield Static(
                f"[bold]Summary:[/bold] {summary}",
                id="preview-summary",
                classes="preview-section",
            )

            # Diff table
            yield Static(
                _build_diff_text(diffs),
                id="preview-diff",
                classes="preview-section",
            )

            # YAML override preview
            with Vertical(classes="preview-section"):
                yield Static(
                    "[bold underline]Override YAML[/bold underline]\n",
                    id="preview-yaml-header",
                )
                yield Static(
                    _build_yaml_preview(draft),
                    id="preview-yaml",
                )
        yield Footer()
