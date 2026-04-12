# Customization

The generator supports layered configuration: start with the base `config.yaml`, optionally merge one or more overlay YAML files on top, and apply leaf-level overrides from the command line. The merged result is validated through the same strict schema as a standalone config.

## Config layering

Merge order (last wins):

1. **Base config** (`config.yaml`) — the full Cascade Industries scenario.
2. **Overlay files** (`--overlay`) — partial YAML files merged in order via deep merge.
3. **CLI overrides** (`--set`) — dotted-key=value pairs applied after overlays.

Deep merge semantics:

- **Dicts** are recursively merged (not replaced).
- **Lists** are replaced wholesale (not appended).
- **`null`** in an overlay deletes the key from the result.
- **Scalars** in the overlay override the base.

After merging, the result must still pass full schema validation (all required fields present, no unknown keys).

## CLI usage

### Default generation (no customization)

```bash
uv run python generate_test_suite.py --output /tmp/test_suite
```

### Apply an overlay preset

```bash
uv run python generate_test_suite.py \
    --overlay presets/small-company.yaml \
    --output /tmp/test_suite
```

### Stack multiple overlays

Overlays merge left-to-right. The second overlay wins on any conflicting keys:

```bash
uv run python generate_test_suite.py \
    --overlay presets/high-growth.yaml presets/higher-noise.yaml \
    --output /tmp/test_suite
```

### Apply leaf-level overrides

`--set` overrides are applied after all overlay files. Values are auto-coerced: integers, floats, booleans (`true`/`false`), and `null` are recognized.

```bash
uv run python generate_test_suite.py \
    --set company.name="Acme Corp" seed=99 \
    --output /tmp/test_suite
```

### Combine overlays and overrides

```bash
uv run python generate_test_suite.py \
    --overlay presets/small-company.yaml \
    --set company.consolidated_revenue=75000000 difficulty.error_density=0.5 \
    --output /tmp/test_suite
```

## Seed behavior

The `seed` field (default: `42`) controls all random state: `random.seed()`, `numpy.random.default_rng()`, and `Faker.seed_instance()`. Changing the seed via `--set seed=99` produces a different but equally deterministic suite.

## Supported v1 fields

### Company parameters

| Field | Type | Description |
|---|---|---|
| `company.name` | string | Company display name |
| `company.consolidated_revenue` | int | Total revenue in USD |
| `company.headquarters` | string | HQ city and state |
| `company.industry` | string | Industry descriptor |
| `company.employees.total_count` | int | Headcount |
| `company.employees.annual_turnover_rate` | float | Annual employee turnover (0.0-1.0) |
| `company.growth_rates.fy2023_to_fy2024` | float | Year-over-year growth rate |
| `company.growth_rates.fy2024_to_fy2025` | float | Year-over-year growth rate |
| `company.seasonal_weights.Q1..Q4` | float | Quarterly revenue distribution (must sum to ~1.0) |

### Difficulty profile

| Field | Type | Default | Description |
|---|---|---|---|
| `difficulty.error_density` | float | `1.0` | Fraction of planted errors to include (0.0-1.0) |
| `difficulty.canary_visibility` | string | `"visible"` | `"visible"`, `"subtle"`, or `"hidden"` |
| `difficulty.judgment_trap_density` | float | `1.0` | Fraction of judgment traps to include (0.0-1.0) |

### Output profile

| Field | Type | Default | Description |
|---|---|---|---|
| `output.enabled_test_cases` | list | `[]` (all) | Restrict generation to these test case IDs |
| `output.enabled_packs` | list | `[]` (default) | Restrict generation to these pack IDs |

### Error profile

| Field | Type | Default | Description |
|---|---|---|---|
| `errors.include` | list | `[]` | Only include these error IDs (allowlist) |
| `errors.exclude` | list | `[]` | Exclude these error IDs (blocklist) |
| `errors.density_override` | float/null | `null` | Override error injection probability (0.0-1.0) |

`include` and `exclude` must not overlap.

## Bundled presets

The `presets/` directory contains ready-to-use overlay files:

| Preset | What it changes |
|---|---|
| `small-company.yaml` | Single subsidiary, ~$50M revenue, 120 employees |
| `high-growth.yaml` | 18-25% YoY growth, higher turnover, Q4-heavy seasonality |
| `flat-seasonality.yaml` | Equal quarterly weights (25% each) |
| `higher-noise.yaml` | Full error density, subtle canaries, full judgment traps |
| `easy-mode.yaml` | 30% error density, visible canaries, 50% judgment traps |

## Validation

The merged config is validated against the same schema as a standalone `config.yaml`. Common validation errors:

- **Missing required field** — overlays are partial, but the merged result must be complete.
- **Unknown key** — typos or aspirational fields not yet in the v1 schema are rejected.
- **Type mismatch** — e.g. a string where a float is expected.
- **Range violation** — density values must be 0.0-1.0; `canary_visibility` must be one of the allowed values.
- **Overlap** — `errors.include` and `errors.exclude` must not share entries.

## Future limitations

The v1 customization surface covers company parameters, difficulty tuning, and output/error filtering. The following are **not yet supported** and will be rejected if included in an overlay:

- Adding or removing subsidiaries
- Changing the chart of accounts structure
- Custom canary assignment overrides
- Custom error injection definitions (beyond include/exclude filtering)
- Per-test-case difficulty overrides

These may be added in future versions. To request a new customization field, file a bead.
