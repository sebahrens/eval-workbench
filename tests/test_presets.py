"""Validate that all preset override YAML files are well-formed and merge
cleanly with the base config.

Bead: synth-data-2u6.4
"""

from __future__ import annotations

from pathlib import Path

import pytest
import yaml

from generator.config import load_layered_config

PRESETS_DIR = Path(__file__).resolve().parent.parent / "presets"
BASE_CONFIG = Path(__file__).resolve().parent.parent / "config.yaml"

# Collect all .yaml files in presets/
PRESET_FILES = sorted(PRESETS_DIR.glob("*.yaml"))


@pytest.mark.parametrize(
    "preset_path",
    PRESET_FILES,
    ids=[p.stem for p in PRESET_FILES],
)
def test_preset_loads_with_base(preset_path: Path) -> None:
    """Each preset merges with the base config without validation errors."""
    cfg = load_layered_config(BASE_CONFIG, layers=[preset_path])
    # Basic sanity: the merged config has a company name and positive revenue
    assert cfg.company.name
    assert cfg.company.consolidated_revenue > 0


@pytest.mark.parametrize(
    "preset_path",
    PRESET_FILES,
    ids=[p.stem for p in PRESET_FILES],
)
def test_preset_is_valid_yaml(preset_path: Path) -> None:
    """Each preset is parseable YAML containing only a dict."""
    raw = yaml.safe_load(preset_path.read_text())
    assert isinstance(raw, dict), f"{preset_path.name} root must be a mapping"


@pytest.mark.parametrize(
    "preset_path",
    PRESET_FILES,
    ids=[p.stem for p in PRESET_FILES],
)
def test_preset_uses_only_v1_fields(preset_path: Path) -> None:
    """Presets must not include fields outside the v1 supported surface."""
    raw = yaml.safe_load(preset_path.read_text())

    # v1 allowed top-level keys in overlay files
    allowed_top = {"seed", "output_dir", "company", "difficulty", "output", "errors"}
    unknown_top = set(raw) - allowed_top
    assert not unknown_top, f"Unknown top-level keys: {unknown_top}"
