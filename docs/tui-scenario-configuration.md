# TUI Scenario Configuration

The TUI provides a keyboard-driven, guided workflow for building scenario override YAML files without hand-editing YAML or mutating `config.yaml`. It covers company profile, difficulty/output profile, test case overrides, validation, diff preview, atomic save, and optional generation run.

## Prerequisites

The TUI requires the `textual` extra:

```bash
uv pip install synth-data[tui]
```

## Launch

```bash
# Default: load config.yaml as the base
uv run python -m generator.tui

# Specify a base config
uv run python -m generator.tui --config path/to/config.yaml

# Pre-apply file overlays before editing
uv run python -m generator.tui --overlay presets/small-company.yaml presets/high-growth.yaml
```

Overlays are merged left-to-right onto the base config before the TUI opens, following the same deep-merge semantics as the CLI (see [customization.md](customization.md)).

## Screens

The TUI is organized into sequential screens. Navigation uses keyboard bindings shown in the footer.

### 1. Company Profile

Edit company-level parameters: name, headquarters, industry, consolidated revenue, headcount, turnover rate, growth rates, and quarterly seasonal weights. Subsidiary fields are templated per the base config's subsidiary keys.

Fields use typed inputs (text, integer, float, choice) with range validation. Modified fields are highlighted to distinguish overrides from base values.

### 2. Difficulty & Output Profile

Configure presentation difficulty (error density, canary visibility, judgment trap density) and output filtering (enabled test cases, enabled packs). This screen controls noise and scope without changing the canonical financial model.

### 3. Test Case Overrides

Lists TC-01 through TC-18 with titles. In v1, no per-test-case override fields are supported, so each TC shows an informational placeholder. This screen will gain fields in future versions.

### 4. Validation Panel

Real-time validation of the merged config (base + draft overrides) through the same schema validator used by the CLI. Errors are grouped by config section. Each error row is clickable to navigate to the offending field.

### 5. Config Preview & Diff

Read-only view showing:

- **Change summary** -- field-level diffs (base value -> draft value) for every override.
- **Raw YAML** -- the override-only YAML that would be saved to disk.

This screen never mutates the draft.

### 6. Save

Atomically saves the draft override to a YAML file using write-to-temp + rename. Key behaviors:

- Only the **override delta** is saved, not the full merged config. The base `config.yaml` is never modified.
- If the target file already exists, the TUI refuses to overwrite unless force is confirmed.
- Before writing, the merged config is validated. Invalid configs cannot be saved.

### 7. Generation Run (optional)

After saving, this screen shows the constructed CLI command for generating the test suite with the saved override applied as an overlay. You can:

- **Review** the command and copy it for manual execution.
- **Run** the command directly from the TUI (executes in a subprocess with a 5-minute timeout).

The screen also shows the validation status before running.

## Equivalent CLI Command

Everything the TUI produces is a standard override YAML file. The equivalent CLI workflow for a TUI-saved override at `my-scenario.yaml`:

```bash
# Generate using the saved override as an overlay
uv run python generate_test_suite.py \
    --overlay my-scenario.yaml \
    --output /tmp/test_suite

# Or with a non-default base config
uv run python generate_test_suite.py \
    --config path/to/config.yaml \
    --overlay my-scenario.yaml \
    --output /tmp/test_suite
```

The saved override YAML has the same format as any overlay file and can be stacked with other overlays or `--set` overrides.

## Override YAML format

The saved file contains only the fields you changed, not the full config. Example:

```yaml
company:
  name: Acme Corp
  consolidated_revenue: 75000000
difficulty:
  error_density: 0.5
  canary_visibility: subtle
```

This file is applied via deep merge onto the base config at generation time.

## Validation workflow

The TUI validates continuously:

1. **On field edit** -- the draft service coerces the value against schema metadata (type, range, allowed choices). Invalid input is rejected inline.
2. **On save** -- the full merged config is validated through `load_layered_config`. Save is blocked if validation fails.
3. **On generation run** -- validation status is shown before the run button becomes active.

Common validation errors:

| Error | Cause |
|---|---|
| Missing required field | Override deleted a required field via `null` |
| Unknown key | Typo in a field path |
| Type mismatch | String where float expected, etc. |
| Range violation | Density value outside 0.0-1.0 |
| Overlap | `errors.include` and `errors.exclude` share entries |

## v1 Limitations

The following are **not editable** in the TUI and will be rejected if attempted:

- Adding or removing subsidiaries
- Changing the chart of accounts structure
- Custom canary assignment overrides
- Custom error injection definitions (beyond include/exclude filtering)
- Per-test-case difficulty overrides (TC override screen is read-only in v1)

These match the CLI v1 customization surface documented in [customization.md](customization.md#future-limitations).

## Keybindings

| Key | Action |
|---|---|
| `q` | Quit the TUI |
| `?` | Show help |
| `Escape` | Go back / dismiss current screen |

Screen-specific bindings are shown in the footer bar.
