"""Textual application shell for scenario configuration.

This module contains the top-level App subclass and the ``configure``
entry point.  It does not mutate config or generate output on launch —
it delegates all state management to a service object (synth-data-2u6.6.3).
"""

from __future__ import annotations

import sys


def _check_textual() -> None:
    """Verify that textual is installed, exit with a helpful message if not."""
    try:
        import textual  # noqa: F401
    except ImportError:
        print(
            "Error: the 'textual' package is required for the TUI.\n"
            "Install it with:  uv pip install synth-data[tui]",
            file=sys.stderr,
        )
        sys.exit(1)


def configure(config_path: str = "config.yaml", overlay: list[str] | None = None) -> None:
    """Launch the scenario configuration TUI.

    Parameters
    ----------
    config_path : str
        Path to the base configuration YAML file.
    overlay : list[str] | None
        Optional overlay YAML paths to merge before editing.
    """
    _check_textual()

    from textual.app import App, ComposeResult
    from textual.widgets import Footer, Header, Static

    class ConfigureApp(App):
        """Scenario configuration TUI — shell for future screens."""

        TITLE = "Cascade Industries — Scenario Configuration"
        BINDINGS = [
            ("q", "quit", "Quit"),
            ("?", "help", "Help"),
        ]

        def __init__(self, config_path: str, overlay: list[str] | None = None) -> None:
            super().__init__()
            self.config_path = config_path
            self.overlay = overlay or []

        def compose(self) -> ComposeResult:
            yield Header()
            yield Static(
                f"Config: {self.config_path}\n"
                f"Overlays: {', '.join(self.overlay) if self.overlay else '(none)'}\n\n"
                "Screens will be added by subsequent beads.\n"
                "Press [bold]q[/bold] to quit.",
                id="welcome",
            )
            yield Footer()

        def action_help(self) -> None:
            """Show help — placeholder for future help screen."""
            self.notify("Help screen coming soon.", title="Help")

    app = ConfigureApp(config_path=config_path, overlay=overlay)
    app.run()
