# Layered Customization Schema — v1 Design

Decision record for bead `synth-data-2u6.1`. This defines the schema, merge semantics, seed behavior, and v1 field surface for the scenario customization system.

## Layer Stack (merge order, last wins)

Layers are merged bottom-to-top. Each layer is an optional YAML file. The generator loads them in this order and deep-merges, with later layers overriding earlier ones at the leaf-key level.

```
1. defaults           (hardcoded in code — the current config.yaml values)
2. preset             (--preset=<name>  → presets/<name>.yaml)
3. company_profile    (--company=<path> → user-supplied company YAML)
4. industry_profile   (--industry=<name> → presets/industries/<name>.yaml)
5. difficulty_profile (--difficulty=easy|medium|hard|adversarial)
6. output_profile     (--output=<path>  → controls which TCs to emit, formats)
7. error_profile      (--errors=<path>  → override planted error selection/density)
8. test_case_overrides(--tc-override=<path> → per-TC parameter tweaks)
9. CLI flags          (--set key=value  → leaf-level override from command line)
```

## Merge Semantics

- **Deep merge**: dicts are recursively merged, not replaced. `company.subsidiaries.precision_components.revenue: 50_000_000` in a layer only overrides that one field.
- **List replace**: lists (e.g., `company.years`, `company.employees.remote_states`) are replaced wholesale, not appended. Avoids ambiguity about ordering and deduplication.
- **Null deletes**: setting a key to `null` in a higher layer removes it from the merged result. Allows a preset to remove a subsidiary.
- **Type preservation**: a layer cannot change a field's type (int to str). The validator rejects type mismatches after merge.

## Seed Behavior

- `seed` is always present in the final merged config (defaults to 42).
- Each layer can set `seed`, but it's leaf-level — it doesn't compose.
- The seed drives ALL randomness: Faker, numpy, random, and canary generation.
- A preset that changes company parameters but keeps seed=42 will produce different data (because the data model depends on company params), but the generation process remains deterministic.
- Changing the seed changes everything. No per-subsystem seed isolation in v1.

## v1 Supported Fields (safe to customize)

### Company Profile
- `company.name` (str)
- `company.type` (str)
- `company.industry` (str)
- `company.headquarters` (str)
- `company.fiscal_year_end` (str, MM-DD format)
- `company.years` (list[int], 3 consecutive)
- `company.current_year` (int, must be max of years)
- `company.consolidated_revenue` (int, >0)

### Subsidiary Structure
- `company.subsidiaries.*` (full subsidiary dicts)
- Adding/removing subsidiaries is allowed; minimum 1 required
- Each subsidiary must have all required fields

### Financial Parameters
- `company.growth_rates.*` (floats)
- `company.intercompany.*` (floats/ints)
- `company.employees.*` (int/float/list)
- `company.seasonal_weights.*` (floats, must sum to 1.0)

### Generator Control
- `seed` (int)
- `output_dir` (str)

### Output Profile (new)
- `output.enabled_test_cases` (list[str], e.g. `["TC-01", "TC-06", "TC-09"]`)
- `output.enabled_packs` (list[str], e.g. `["accounting", "legal_hr_diligence"]`)
- `output.formats` (reserved for v2 — currently always xlsx/docx/pdf per TC spec)

### Difficulty Profile (new)
- `difficulty.error_density` (float, 0.0-1.0, fraction of possible errors to inject)
- `difficulty.canary_visibility` (str: `"visible"` | `"subtle"` | `"hidden"`)
- `difficulty.judgment_trap_density` (float, 0.0-1.0)

### Error Profile (new)
- `errors.include` (list[str], error IDs to always inject)
- `errors.exclude` (list[str], error IDs to never inject)
- `errors.density_override` (float, overrides difficulty.error_density)

## v1 Unsupported Fields (reject with clear error)

These fields are internal and must NOT be customizable. The validator emits a clear error if a user tries to set them:

- `canary_assignments` — computed from seed, not user-settable
- `error_injections` — derived from error_registry.json + error_profile
- `company.subsidiaries.*.entity_code` — used as join keys across TCs; changing them breaks cross-referential integrity
- Any field not listed above — reject with "unsupported in v1"

## Validation Timing

1. **Per-layer**: each layer file is validated for YAML syntax and type correctness before merge.
2. **Post-merge**: the merged config is validated for semantic constraints (seasonal weights sum, subsidiary count, year ordering, etc.).
3. **Pre-generation**: the final Config dataclass runs `__post_init__` checks (existing behavior).

Validation errors reference the layer that introduced the bad value, not just the merged result.

## Example: Merged Scenario

Given:
- defaults: Cascade Industries, seed=42, $200M revenue
- preset: `presets/small-company.yaml` with revenue=50M, 1 subsidiary
- difficulty: `--difficulty=easy` with error_density=0.3, canary_visibility=visible
- CLI: `--set company.name="Acme Corp"`

Result:
```yaml
seed: 42
company:
  name: "Acme Corp"            # from CLI flag
  consolidated_revenue: 50_000_000  # from preset
  subsidiaries:                # from preset (1 subsidiary)
    main:
      legal_name: "Acme Main LLC"
      ...
difficulty:
  error_density: 0.3           # from difficulty profile
  canary_visibility: "visible"
```

## File Layout

```
presets/
├── default.yaml                # The current config.yaml values, reorganized
├── small-company.yaml          # 1 subsidiary, ~$50M
├── large-conglomerate.yaml     # 5+ subsidiaries, ~$2B
└── industries/
    ├── healthcare.yaml
    ├── technology.yaml
    └── financial-services.yaml
```

## Migration Path

The current `config.yaml` becomes the defaults layer. Existing behavior is preserved when no layers are specified — the generator produces identical output to today. This is the backward-compatibility contract.
