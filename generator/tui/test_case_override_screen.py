"""Test-case override configuration screen.

Lists TC-01 through TC-18 with their titles and exposes supported v1
per-TC override fields.  In v1, no per-TC override fields exist, so each
TC shows an empty-state message.

Bead: synth-data-2u6.6.7
"""

from __future__ import annotations

from textual.app import ComposeResult
from textual.containers import Vertical, VerticalScroll
from textual.screen import Screen
from textual.widgets import Footer, Header, Label, Static

from generator.tui.draft_service import DraftService

# ---------------------------------------------------------------------------
# Test case catalogue
# ---------------------------------------------------------------------------

TC_CATALOGUE: list[tuple[str, str]] = [
    ("TC-01", "Trial Balance Reconciliation"),
    ("TC-02", "Bank Reconciliation & Confirmation Matching"),
    ("TC-03", "Substantive Analytical Procedures — Revenue"),
    ("TC-04", "Lease Extraction & Schedule Population (ASC 842)"),
    ("TC-05", "Audit Workpaper Memo — Accounts Receivable"),
    ("TC-06", "Tax Provision (ASC 740)"),
    ("TC-07", "K-1 Extraction & Consolidation"),
    ("TC-08", "R&D Tax Credit Study (Section 41)"),
    ("TC-09", "Transfer Pricing Documentation"),
    ("TC-10", "Multi-State Apportionment"),
    ("TC-11", "Quality of Earnings (Financial Due Diligence)"),
    ("TC-12", "Data Room Triage & Document Index"),
    ("TC-13", "Forensic AP Transaction Analysis"),
    ("TC-14", "13-Week Cash Flow Forecast"),
    ("TC-15", "DCF Valuation"),
    ("TC-16", "Engagement Letter Generation"),
    ("TC-17", "Multi-File Deliverable Assembly"),
    ("TC-18", "Prior Year Workpaper Rollforward"),
]

_NO_OVERRIDES_MSG = (
    "[dim]No per-test-case override fields are available in v1. "
    "Use the Difficulty & Output screen to enable/disable entire test cases.[/dim]"
)

_SCREEN_HELP = (
    "[bold]Test Case Overrides[/bold]\n"
    "This screen lists all 18 test cases. In the current version (v1), "
    "per-TC parameter tweaks are not yet supported — each test case "
    "inherits its behavior from the company profile, difficulty, and "
    "output settings.\n\n"
    "To control which TCs are generated, use the "
    "[bold]Difficulty & Output[/bold] screen."
)


# ---------------------------------------------------------------------------
# Screen
# ---------------------------------------------------------------------------


class TestCaseOverrideScreen(Screen):
    """Configuration screen for per-test-case overrides."""

    BINDINGS = [
        ("escape", "dismiss", "Back"),
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
        yield Header()
        with VerticalScroll():
            yield Static(_SCREEN_HELP, classes="section-help")
            for tc_id, tc_title in TC_CATALOGUE:
                with Vertical(classes="field-section", id=f"section-{tc_id}"):
                    yield Label(
                        f"[bold]{tc_id}[/bold]: {tc_title}",
                        id=f"lbl-{tc_id}",
                    )
                    yield Static(
                        _NO_OVERRIDES_MSG,
                        id=f"empty-{tc_id}",
                        classes="tc-empty-state",
                    )
        yield Footer()

    def get_tc_ids(self) -> list[str]:
        """Return the list of test case IDs shown on this screen."""
        return [tc_id for tc_id, _ in TC_CATALOGUE]

    def get_tc_title(self, tc_id: str) -> str | None:
        """Return the title for a given TC ID, or None if not found."""
        for tid, title in TC_CATALOGUE:
            if tid == tc_id:
                return title
        return None

    def has_overrides(self, tc_id: str) -> bool:
        """Return True if the given TC has any supported v1 override fields.

        In v1, this is always False.
        """
        return False
