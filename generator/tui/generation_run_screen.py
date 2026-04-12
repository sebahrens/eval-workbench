"""Generation validation and run screen.

After saving an override, this screen shows the constructed CLI command
for generating the test suite and optionally runs it in a subprocess.

Bead: synth-data-2u6.6.11
"""

from __future__ import annotations

import shlex
import subprocess
import sys

from textual.app import ComposeResult
from textual.containers import Vertical, VerticalScroll
from textual.screen import Screen
from textual.widgets import Button, Footer, Header, Static

from generator.tui.draft_service import DraftService

# ---------------------------------------------------------------------------
# Command builder (pure function, no side effects)
# ---------------------------------------------------------------------------


def build_generate_command(
    *,
    config_path: str = "config.yaml",
    overlay: str | None = None,
    output_dir: str = "/tmp/test_suite",
    packs: list[str] | None = None,
) -> list[str]:
    """Build the CLI argument list for ``generate_test_suite.py``.

    Returns a list suitable for :func:`subprocess.run` or display.

    Parameters
    ----------
    config_path : str
        Path to the base configuration YAML.
    overlay : str | None
        Path to the saved override YAML file (if any).
    output_dir : str
        Target directory for generated output.
    packs : list[str] | None
        Scenario pack IDs to pass via ``--packs``.  ``None`` means default.
    """
    cmd = [sys.executable, "generate_test_suite.py"]

    if config_path != "config.yaml":
        cmd.extend(["--config", config_path])

    if overlay:
        cmd.extend(["--overlay", overlay])

    cmd.extend(["--output", output_dir])

    if packs:
        cmd.extend(["--packs"] + list(packs))

    return cmd


def format_command(cmd: list[str]) -> str:
    """Format a command list as a copy-pasteable shell string."""
    return " ".join(shlex.quote(c) for c in cmd)


# ---------------------------------------------------------------------------
# Screen
# ---------------------------------------------------------------------------

_HELP_TEXT = (
    "[bold]Generation Run[/bold]\n\n"
    "Review the generation command below.  Press [bold]Run[/bold] to execute "
    "it, or [bold]Back[/bold] to return without running.\n\n"
    "The command uses your saved override as an overlay on top of the base "
    "config and writes output to the directory shown."
)


class GenerationRunScreen(Screen):
    """Screen that shows and optionally runs the test suite generator."""

    BINDINGS = [
        ("escape", "dismiss", "Back"),
        ("q", "dismiss", "Back"),
    ]

    def __init__(
        self,
        service: DraftService,
        *,
        overlay_path: str | None = None,
        output_dir: str = "/tmp/test_suite",
        packs: list[str] | None = None,
        name: str | None = None,
        id: str | None = None,
        classes: str | None = None,
    ) -> None:
        super().__init__(name=name, id=id, classes=classes)
        self.service = service
        self.overlay_path = overlay_path
        self.output_dir = output_dir
        self.packs = packs
        self._cmd = build_generate_command(
            config_path=str(service._base_path),
            overlay=overlay_path,
            output_dir=output_dir,
            packs=packs,
        )

    def compose(self) -> ComposeResult:
        yield Header()
        with VerticalScroll():
            yield Static(_HELP_TEXT, classes="section-help")

            # Validation status
            result = self.service.validate()
            if result.valid:
                yield Static(
                    "[green bold]✓ Configuration is valid[/green bold]",
                    id="gen-validation-status",
                )
            else:
                n = len(result.errors)
                yield Static(
                    f"[red bold]✗ Configuration has {n} error{'s' if n != 1 else ''}[/red bold]\n"
                    "[dim]Fix validation errors before running generation.[/dim]",
                    id="gen-validation-status",
                )

            # Command preview
            with Vertical(classes="preview-section"):
                yield Static(
                    "[bold underline]Command[/bold underline]",
                    id="gen-cmd-header",
                )
                yield Static(
                    format_command(self._cmd),
                    id="gen-cmd-preview",
                )

            # Output directory
            yield Static(
                f"[bold]Output:[/bold] {self.output_dir}",
                id="gen-output-dir",
            )

            # Run / Back buttons
            yield Button("Run Generation", id="gen-run-btn", variant="primary")
            yield Static("", id="gen-run-output")
        yield Footer()

    def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "gen-run-btn":
            self._run_generation()

    def _run_generation(self) -> None:
        """Execute the generation command and display results."""
        output_widget = self.query_one("#gen-run-output", Static)
        output_widget.update("[dim]Running generation...[/dim]")

        try:
            proc = subprocess.run(
                self._cmd,
                capture_output=True,
                text=True,
                timeout=300,
            )
            if proc.returncode == 0:
                output_widget.update(
                    f"[green bold]✓ Generation completed successfully[/green bold]\n"
                    f"[dim]{proc.stdout.strip() if proc.stdout.strip() else 'Done.'}[/dim]"
                )
            else:
                output_widget.update(
                    f"[red bold]✗ Generation failed (exit code {proc.returncode})[/red bold]\n"
                    f"{proc.stderr.strip()}"
                )
        except subprocess.TimeoutExpired:
            output_widget.update(
                "[red bold]✗ Generation timed out after 5 minutes[/red bold]"
            )
        except Exception as exc:
            output_widget.update(
                f"[red bold]✗ Error: {exc}[/red bold]"
            )
