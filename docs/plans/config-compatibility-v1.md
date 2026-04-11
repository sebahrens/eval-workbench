# Config Compatibility and Migration Policy — v1

Decision record for bead `synth-data-2u6.7`. Defines the backward-compatibility contract, unknown-field policy, versioning stance, and migration path for config.yaml and future layered overrides.

## Backward-Compatibility Contract

The current `config.yaml` at the project root is a valid v1 config. It loads without error and produces identical generator output. This contract holds as long as the config contains only v1 supported fields (see `customization-schema-v1.md`).

When no layers, presets, or overrides are specified, the generator behaves exactly as it does today. This is the zero-regression guarantee.

## Unknown-Field Policy

### Strict rejection

All config sections enforce a closed schema. Any key not in the v1 supported-field surface is rejected at load time with a `ConfigError` listing every unknown key.

**Scope:**

| Section | Allowed keys |
|---|---|
| Config root | `seed`, `output_dir`, `company`, `canary_assignments`, `error_injections`, `augmentation` |
| `company` | `name`, `type`, `industry`, `headquarters`, `fiscal_year_end`, `years`, `current_year`, `consolidated_revenue`, `subsidiaries`, `growth_rates`, `intercompany`, `employees`, `seasonal_weights` |
| `company.subsidiaries.<key>` | `legal_name`, `location`, `state`, `entity_code`, `revenue`, `type`, `gross_margin`, `employee_count`, `rd_spend_pct` |

### Rationale

Strict rejection is preferred over warn-and-ignore because:
- Silent ignoring masks typos (e.g., "employes" vs "employees").
- Config drift between what users think they configured and what the generator uses is the primary source of confusion in deterministic generators.
- Adding a new field is a deliberate v2 surface expansion, not an accidental omission.

### Error message format

```
ConfigError: Unknown key(s) in config root: ['difficulty', 'packs'].
Only these keys are supported in v1: ['augmentation', 'canary_assignments', ...]
```

All unknown keys are reported in a single error, not one-at-a-time.

## Build-Time Fields

`canary_assignments` and `error_injections` are build-time populated by the generator. They are accepted in config.yaml (defaulting to empty dicts) but are not user-editable in the customization sense — the generator overwrites them during generation.

`augmentation` controls optional LLM-based data augmentation and is also accepted.

## Config Versioning

### v1 stance

- **No version marker in v1.** The current config.yaml has no `config_version` field and works. Adding a mandatory version field would break existing configs for no gain.
- Configs without a `config_version` field are treated as v1.

### Future v2 migration

When the surface expands (v2):
1. A `config_version: 2` field will be added to the allowed top-level keys.
2. Configs without a version field are treated as v1 automatically.
3. The loader applies v1 defaults for any v2-new fields.
4. A one-time deprecation warning suggests adding `config_version: 2`.

### No breaking changes in v1 → v2

v2 will be a superset of v1. All v1 configs remain valid v2 configs. No fields will be removed or renamed.

## CLI `--config` Behavior

`--config path/to/config.yaml` loads and validates against the v1 surface. The flag's behavior does not change from the current implementation. Unknown fields in the pointed-to file are rejected identically to the default config path.

## Layered Override Behavior (future)

When layered override files are introduced (synth-data-2u6 epic), they follow the same policy:
- Only v1 surface fields may be overridden.
- Unknown fields in an override file are rejected at load time with the same `ConfigError`.
- Each layer file is validated independently before merge.
- Post-merge validation runs semantic checks (seasonal weights sum, subsidiary count, etc.).

## Test Coverage

The following scenarios are tested in `tests/test_config.py`:

| Test | What it verifies |
|---|---|
| `test_legacy_config_loads_without_mutation` | Real config.yaml loads and preserves all default values |
| `test_unknown_top_level_key_rejected` | Unknown top-level key → ConfigError |
| `test_unknown_company_key_rejected` | Unknown company-level key → ConfigError |
| `test_unknown_subsidiary_key_rejected` | Unknown subsidiary key → ConfigError |
| `test_multiple_unknown_top_keys_all_reported` | All unknown keys appear in one error |
| `test_build_time_fields_accepted` | canary_assignments/error_injections accepted |
| `test_augmentation_field_accepted` | augmentation section accepted |
| `test_rd_spend_pct_accepted_in_subsidiary` | Optional subsidiary field accepted |
| `test_config_without_version_treated_as_v1` | No version field → loads as v1 |
| `test_v1_custom_config_does_not_mutate_defaults` | Loading custom config doesn't affect subsequent default loads |
