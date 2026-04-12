"""Company profile configuration screen.

Displays editable fields for the COMPANY, SUBSIDIARY, and FINANCIAL
field groups.  Subsidiary fields are templated — the screen reads the
current subsidiary keys from the draft service and expands one section
per subsidiary.

Bead: synth-data-2u6.6.5
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
    expand_subsidiary_fields,
    get_fields_by_group,
    is_unsupported_path,
)
from generator.tui.draft_service import DraftService
from generator.tui.validation_panel import ValidationPanel

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

_COMPANY_HELP = (
    "[bold]Company Profile[/bold]\n"
    "Core identity of the parent entity — legal name, structure, "
    "fiscal calendar, and consolidated revenue.  Changes here flow "
    "into every generated test case."
)

_SUBSIDIARY_HELP = (
    "[bold]Subsidiaries[/bold]\n"
    "Each subsidiary's legal name, location, revenue, margin, and "
    "headcount.  Entity codes are derived and cannot be edited."
)

_FINANCIAL_HELP = (
    "[bold]Financial Parameters[/bold]\n"
    "Growth rates, intercompany terms, employee headcount, and "
    "seasonal revenue weights.  Quarterly weights must sum to 1.0."
)


# ---------------------------------------------------------------------------
# Screen
# ---------------------------------------------------------------------------


class CompanyProfileScreen(Screen):
    """Configuration screen for company, subsidiary, and financial fields."""

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
        self._company_fields = get_fields_by_group(FieldGroup.COMPANY)
        self._financial_fields = get_fields_by_group(FieldGroup.FINANCIAL)
        # Subsidiary fields are expanded per key at compose time
        self._sub_keys: list[str] = []
        self._sub_fields: list[FieldMeta] = []

    # ------------------------------------------------------------------
    # Compose
    # ------------------------------------------------------------------

    def compose(self) -> ComposeResult:
        # Discover subsidiary keys from the merged config
        merged = self.service.merged_raw
        subs = (merged.get("company") or {}).get("subsidiaries") or {}
        self._sub_keys = sorted(subs.keys())
        self._sub_fields = []
        for key in self._sub_keys:
            self._sub_fields.extend(expand_subsidiary_fields(key))

        yield Header()
        with VerticalScroll():
            # Company section
            yield Static(_COMPANY_HELP, classes="section-help")
            with Vertical(classes="field-section", id="company-section"):
                for meta in self._company_fields:
                    yield from self._compose_field(meta)

            # Subsidiary sections — one block per subsidiary
            yield Static(_SUBSIDIARY_HELP, classes="section-help")
            for key in self._sub_keys:
                sub_fields = expand_subsidiary_fields(key)
                entity_code = (subs.get(key) or {}).get("entity_code", "")
                heading = f"[bold]{key}[/bold]"
                if entity_code:
                    heading += f"  [dim](entity code: {entity_code})[/dim]"
                with Vertical(
                    classes="field-section",
                    id=f"sub-section-{key}",
                ):
                    yield Static(heading, classes="sub-heading")
                    for meta in sub_fields:
                        yield from self._compose_field(meta)

            # Financial section
            yield Static(_FINANCIAL_HELP, classes="section-help")
            with Vertical(classes="field-section", id="financial-section"):
                for meta in self._financial_fields:
                    yield from self._compose_field(meta)

            # Action buttons
            with Horizontal(classes="button-bar"):
                yield Button("Reset All", id="btn-reset", variant="warning")
                yield Button("Validate", id="btn-validate", variant="default")

        # Validation error panel (outside scroll, pinned at bottom)
        yield ValidationPanel(id="validation-panel")
        yield Footer()

    def _compose_field(self, meta: FieldMeta) -> ComposeResult:
        """Yield widgets for a single field based on its input type."""
        field_id = meta.path.replace(".", "-")
        current = self.service.get_field_value(meta.path)
        modified = self.service.is_modified(meta.path)
        label_text = f"{'* ' if modified else ''}{meta.label}"

        with Vertical(classes="field-row", id=f"row-{field_id}"):
            yield Label(
                f"{label_text}  [dim]{meta.help_text}[/dim]",
                id=f"lbl-{field_id}",
            )

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
                    placeholder="Comma-separated values",
                    id=f"fld-{field_id}",
                )

            elif meta.input_type == InputType.LIST_INT:
                if isinstance(current, list):
                    display = ", ".join(str(v) for v in current)
                else:
                    display = str(current) if current is not None else ""
                yield Input(
                    value=display,
                    placeholder="Comma-separated integers (e.g. 2023, 2024, 2025)",
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
            elif meta.input_type == InputType.LIST_INT:
                if raw == "":
                    return
                try:
                    value = [int(v.strip()) for v in raw.split(",") if v.strip()]
                except ValueError:
                    self._show_field_error(meta, "All values must be integers")
                    return
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
        """Validate and notify the parent to trigger save.

        Blocks dismiss when hard errors are present.
        """
        result = self.service.validate()
        panel = self.query_one("#validation-panel", ValidationPanel)
        panel.result = result
        if result.valid:
            panel.clear()
            self.notify("Configuration is valid.", title="Validation")
            self.dismiss(True)
        else:
            self.notify("Fix validation errors before saving.", severity="error")

    def _reset_all_fields(self) -> None:
        """Reset all company/subsidiary/financial fields to base values."""
        all_fields = self._company_fields + self._sub_fields + self._financial_fields
        for meta in all_fields:
            self.service.reset_field(meta.path)
        self.notify("All fields reset to base values.", title="Reset")
        for meta in all_fields:
            self._refresh_widget(meta)
            self._update_label(meta)

    def _run_validation(self) -> None:
        result = self.service.validate()
        panel = self.query_one("#validation-panel", ValidationPanel)
        panel.result = result
        if result.valid:
            panel.clear()
            self.notify("Configuration is valid.", title="Validation")

    def on_validation_panel_go_to_field(
        self, event: ValidationPanel.GoToField
    ) -> None:
        """Scroll to and focus the widget identified by a field path."""
        field_id = event.field_path.replace(".", "-")
        try:
            widget = self.query_one(f"#fld-{field_id}")
            widget.scroll_visible()
            widget.focus()
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------

    def _find_meta(self, path: str) -> FieldMeta | None:
        """Find field metadata by dotted path."""
        all_fields = self._company_fields + self._sub_fields + self._financial_fields
        for meta in all_fields:
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
            elif meta.input_type == InputType.LIST_INT:
                widget = self.query_one(f"#fld-{field_id}", Input)
                display = ", ".join(str(v) for v in current) if isinstance(current, list) else ""
                widget.value = display
            else:
                widget = self.query_one(f"#fld-{field_id}", Input)
                widget.value = str(current) if current is not None else ""
        except Exception:
            pass


def is_unsupported_company_path(path: str) -> bool:
    """Check if a path is unsupported in the company profile context.

    Delegates to the schema-level unsupported check and additionally
    rejects any path not in the COMPANY, SUBSIDIARY, or FINANCIAL groups.
    """
    if is_unsupported_path(path):
        return True
    from generator.schema_metadata import get_field_by_path

    meta = get_field_by_path(path)
    if meta is None:
        return True
    return meta.group not in (FieldGroup.COMPANY, FieldGroup.SUBSIDIARY, FieldGroup.FINANCIAL)
