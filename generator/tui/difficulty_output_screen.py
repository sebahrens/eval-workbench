"""Difficulty and output profile configuration screen.

Displays editable fields for the DIFFICULTY and OUTPUT field groups,
explaining presentation noise versus canonical model changes and
rejecting unsupported fields.

Bead: synth-data-2u6.6.6
"""

from __future__ import annotations

from textual.app import ComposeResult
from textual.containers import Horizontal, Vertical, VerticalScroll
from textual.screen import Screen
from textual.widgets import (
    Button,
    Footer,
    Header,
    Input,
    Label,
    Select,
    SelectionList,
    Static,
)

from generator.schema_metadata import (
    FieldGroup,
    FieldMeta,
    InputType,
    get_fields_by_group,
    is_unsupported_path,
)
from generator.tui.draft_service import DraftService

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_DIFFICULTY_HELP = (
    "[bold]Difficulty Profile[/bold]\n"
    "These knobs control [italic]presentation noise[/italic] — how hard the "
    "generated test suite is for the AI agent. They affect error injection "
    "density, canary visibility, and judgment-trap frequency.\n\n"
    "They do [bold]not[/bold] change the canonical data model (company "
    "structure, financials, or intercompany relationships). To change those, "
    "use the Company Profile screen."
)

_OUTPUT_HELP = (
    "[bold]Output Profile[/bold]\n"
    "Select which test cases and scenario packs are generated. "
    "Leaving a list empty means \"generate all.\""
)


# ---------------------------------------------------------------------------
# Screen
# ---------------------------------------------------------------------------


class DifficultyOutputScreen(Screen):
    """Configuration screen for difficulty and output profile fields."""

    BINDINGS = [
        ("escape", "dismiss", "Back"),
        ("ctrl+s", "save", "Save"),
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
        self._difficulty_fields = get_fields_by_group(FieldGroup.DIFFICULTY)
        self._output_fields = get_fields_by_group(FieldGroup.OUTPUT)

    # ------------------------------------------------------------------
    # Compose
    # ------------------------------------------------------------------

    def compose(self) -> ComposeResult:
        yield Header()
        with VerticalScroll():
            # Difficulty section
            yield Static(_DIFFICULTY_HELP, classes="section-help")
            with Vertical(classes="field-section", id="difficulty-section"):
                for meta in self._difficulty_fields:
                    yield from self._compose_field(meta)

            # Output section
            yield Static(_OUTPUT_HELP, classes="section-help")
            with Vertical(classes="field-section", id="output-section"):
                for meta in self._output_fields:
                    yield from self._compose_field(meta)

            # Action buttons
            with Horizontal(classes="button-bar"):
                yield Button("Reset All", id="btn-reset", variant="warning")
                yield Button("Validate", id="btn-validate", variant="default")
        yield Footer()

    def _compose_field(self, meta: FieldMeta) -> ComposeResult:
        """Yield widgets for a single field based on its input type."""
        field_id = meta.path.replace(".", "-")
        current = self.service.get_field_value(meta.path)
        modified = self.service.is_modified(meta.path)
        label_text = f"{'* ' if modified else ''}{meta.label}"

        with Vertical(classes="field-row", id=f"row-{field_id}"):
            yield Label(f"{label_text}  [dim]{meta.help_text}[/dim]", id=f"lbl-{field_id}")

            if meta.input_type == InputType.CHOICE:
                options = [(c, c) for c in meta.choices]
                yield Select(
                    options,
                    value=current if current is not None else (meta.default or Select.BLANK),
                    id=f"fld-{field_id}",
                )

            elif meta.input_type == InputType.MULTI_CHOICE:
                selections = [
                    (c, c, c in (current or []))
                    for c in meta.choices
                ]
                yield SelectionList(*selections, id=f"fld-{field_id}")

            elif meta.input_type == InputType.LIST_TEXT:
                display = ", ".join(current) if isinstance(current, list) else (current or "")
                yield Input(
                    value=str(display),
                    placeholder="Comma-separated values (empty = all)",
                    id=f"fld-{field_id}",
                )

            elif meta.input_type in (InputType.FLOAT, InputType.INTEGER):
                range_hint = ""
                if meta.range_min is not None and meta.range_max is not None:
                    range_hint = f" ({meta.range_min}–{meta.range_max})"
                yield Input(
                    value=str(current) if current is not None else "",
                    placeholder=f"Enter a number{range_hint}",
                    id=f"fld-{field_id}",
                )
            else:
                yield Input(
                    value=str(current) if current is not None else "",
                    id=f"fld-{field_id}",
                )

    # ------------------------------------------------------------------
    # Event handlers
    # ------------------------------------------------------------------

    def on_input_changed(self, event: Input.Changed) -> None:
        """Handle text/numeric field changes."""
        widget_id = event.input.id or ""
        if not widget_id.startswith("fld-"):
            return
        path = widget_id.removeprefix("fld-").replace("-", ".")
        meta = self._find_meta(path)
        if meta is None:
            return

        raw = event.value.strip()

        try:
            if meta.input_type == InputType.LIST_TEXT:
                value = [v.strip() for v in raw.split(",") if v.strip()] if raw else []
                self.service.set_field(meta.path, value)
            else:
                if raw == "":
                    return  # don't push empty string to numeric fields
                self.service.set_field(meta.path, raw)
            self._update_label(meta)
            self._clear_error(meta)
        except ValueError as exc:
            self._show_field_error(meta, str(exc))

    def on_select_changed(self, event: Select.Changed) -> None:
        """Handle choice field changes."""
        widget_id = event.select.id or ""
        if not widget_id.startswith("fld-"):
            return
        path = widget_id.removeprefix("fld-").replace("-", ".")
        meta = self._find_meta(path)
        if meta is None:
            return

        try:
            self.service.set_field(meta.path, str(event.value))
            self._update_label(meta)
            self._clear_error(meta)
        except ValueError as exc:
            self._show_field_error(meta, str(exc))

    def on_selection_list_selected_changed(
        self, event: SelectionList.SelectedChanged
    ) -> None:
        """Handle multi-choice (SelectionList) changes."""
        widget_id = event.selection_list.id or ""
        if not widget_id.startswith("fld-"):
            return
        path = widget_id.removeprefix("fld-").replace("-", ".")
        meta = self._find_meta(path)
        if meta is None:
            return

        selected = list(event.selection_list.selected)
        try:
            self.service.set_field(meta.path, selected)
            self._update_label(meta)
            self._clear_error(meta)
        except ValueError as exc:
            self._show_field_error(meta, str(exc))

    def on_button_pressed(self, event: Button.Pressed) -> None:
        if event.button.id == "btn-reset":
            self._reset_all_fields()
        elif event.button.id == "btn-validate":
            self._run_validation()

    # ------------------------------------------------------------------
    # Actions
    # ------------------------------------------------------------------

    def action_save(self) -> None:
        """Validate and notify the parent to trigger save."""
        result = self.service.validate()
        if result.valid:
            self.notify("Configuration is valid.", title="Validation")
            self.dismiss(True)
        else:
            self.notify(
                "\n".join(result.errors), title="Validation Failed", severity="error"
            )

    def _reset_all_fields(self) -> None:
        """Reset all difficulty/output fields to base values."""
        for meta in self._difficulty_fields + self._output_fields:
            self.service.reset_field(meta.path)
        # Refresh screen by re-mounting
        self.notify("All fields reset to base values.", title="Reset")
        # Refresh widget values
        for meta in self._difficulty_fields + self._output_fields:
            self._refresh_widget(meta)
            self._update_label(meta)

    def _run_validation(self) -> None:
        result = self.service.validate()
        if result.valid:
            self.notify("Configuration is valid.", title="Validation")
        else:
            self.notify(
                "\n".join(result.errors), title="Validation Failed", severity="error"
            )

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _find_meta(self, path: str) -> FieldMeta | None:
        """Find field metadata by dotted path."""
        for meta in self._difficulty_fields + self._output_fields:
            if meta.path == path:
                return meta
        return None

    def _update_label(self, meta: FieldMeta) -> None:
        """Update the label to reflect modified state."""
        field_id = meta.path.replace(".", "-")
        try:
            lbl = self.query_one(f"#lbl-{field_id}", Label)
            modified = self.service.is_modified(meta.path)
            prefix = "* " if modified else ""
            lbl.update(f"{prefix}{meta.label}  [dim]{meta.help_text}[/dim]")
        except Exception:
            pass

    def _show_field_error(self, meta: FieldMeta, msg: str) -> None:
        """Show a notification for a field error."""
        self.notify(msg, title=f"Error: {meta.label}", severity="error")

    def _clear_error(self, meta: FieldMeta) -> None:
        """Placeholder for clearing field-level errors."""
        pass

    def _refresh_widget(self, meta: FieldMeta) -> None:
        """Refresh a widget's displayed value from the service."""
        field_id = meta.path.replace(".", "-")
        current = self.service.get_field_value(meta.path)

        try:
            if meta.input_type == InputType.CHOICE:
                widget = self.query_one(f"#fld-{field_id}", Select)
                widget.value = current if current is not None else Select.BLANK
            elif meta.input_type == InputType.MULTI_CHOICE:
                widget = self.query_one(f"#fld-{field_id}", SelectionList)
                current_list = current or []
                for idx, choice in enumerate(meta.choices):
                    if choice in current_list:
                        widget.select(idx)
                    else:
                        widget.deselect(idx)
            elif meta.input_type == InputType.LIST_TEXT:
                widget = self.query_one(f"#fld-{field_id}", Input)
                display = ", ".join(current) if isinstance(current, list) else ""
                widget.value = display
            else:
                widget = self.query_one(f"#fld-{field_id}", Input)
                widget.value = str(current) if current is not None else ""
        except Exception:
            pass


def is_unsupported_difficulty_output_path(path: str) -> bool:
    """Check if a path is unsupported in the difficulty/output context.

    Delegates to the schema-level unsupported check and additionally
    rejects any path not in the DIFFICULTY or OUTPUT groups.
    """
    if is_unsupported_path(path):
        return True
    from generator.schema_metadata import get_field_by_path

    meta = get_field_by_path(path)
    if meta is None:
        return True
    return meta.group not in (FieldGroup.DIFFICULTY, FieldGroup.OUTPUT)
