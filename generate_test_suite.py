#!/usr/bin/env python3
"""Cascade Industries test suite generator — orchestrator entry point.

Usage:
    python generate_test_suite.py --output /tmp/test_suite
    python generate_test_suite.py --config config.yaml --output /tmp/test_suite
"""

from __future__ import annotations

import argparse
import random
import sys
from pathlib import Path

import numpy as np
from faker import Faker

from generator.canaries import build_registry as build_canary_registry
from generator.config import Config, ConfigError, load_config
from generator.errors import ErrorRegistry
from generator.formatters.tc01 import emit_tc01
from generator.formatters.tc02 import emit_tc02
from generator.formatters.tc05 import emit_tc05
from generator.formatters.tc06 import emit_tc06
from generator.formatters.tc07 import emit_tc07
from generator.golds.framework import emit_all_golds
from generator.manifest import Manifest
from generator.model.build import build_model

# Number of test cases defined in prompt.md §3–§6
_TEST_CASE_COUNT = 18

# File keys that need canaries — sorted for deterministic registry.
# Each TC formatter adds its files here as it is wired in.
_CANARY_FILE_KEYS: list[str] = sorted([
    "cascade_tb_fy2025",
    "cascade_tb_fy2024_workpaper",
    "cascade_financials_fy2024_signed",
    "ar_aging_fy2025",
    "ar_confirmations_summary",
    "allowance_analysis",
    "workpaper_memo_template",
    # TC-02 files
    "bank_confirmation_fy2025",
    "bank_statement_dec2025",
    "cascade_gl_cash_dec2025",
    # TC-06 files
    "cascade_consolidated_tb_fy2025",
    "tax_provision_fy2024_workpaper",
    "perm_temp_differences_fy2025",
    "statutory_rates",
    # TC-07 files (8 K-1 PDFs + org chart)
    "tc07_k1_001",
    "tc07_k1_002",
    "tc07_k1_003",
    "tc07_k1_004",
    "tc07_k1_005",
    "tc07_k1_006",
    "tc07_k1_007",
    "tc07_k1_008",
    "tc07_entity_org_chart",
])


def _seed_all(seed: int) -> None:
    """Seed every PRNG used by the generator for determinism."""
    random.seed(seed)
    np.random.seed(seed)
    Faker.seed(seed)


def _create_directory_tree(output: Path) -> None:
    """Create the output directory structure per prompt.md §2.3."""
    output.mkdir(parents=True, exist_ok=True)

    # Top-level directories
    (output / "shared_data").mkdir(exist_ok=True)
    (output / "gold_standards").mkdir(exist_ok=True)
    (output / "scoring").mkdir(exist_ok=True)
    (output / "templates").mkdir(exist_ok=True)

    # Per-test-case directories with input_files subdirectory
    for i in range(1, _TEST_CASE_COUNT + 1):
        tc_dir = output / "test_cases" / f"TC-{i:02d}"
        (tc_dir / "input_files").mkdir(parents=True, exist_ok=True)


def generate(config: Config, output: Path) -> Manifest:
    """Run the full generation pipeline.

    Returns the Manifest so callers (tests) can inspect registered files.
    Model builders and formatters will be wired in by later beads.
    """
    _seed_all(config.seed)
    _create_directory_tree(output)

    # ── Phase 1: Build canonical model ──────────────────────────────
    model = build_model(config.seed)

    # ── Canary & error registries ────────────────────────────────────
    canaries = build_canary_registry(_CANARY_FILE_KEYS, seed=config.seed)
    errors = ErrorRegistry()

    with Manifest(output) as manifest:
        # ── Phase 2: TC formatters ───────────────────────────────────
        emit_tc01(model, output, canaries, errors, manifest)
        emit_tc02(model, output, canaries, errors, manifest)
        emit_tc05(model, output, canaries, errors, manifest)
        emit_tc06(model, output, canaries, errors, manifest)
        emit_tc07(model, output, canaries, errors, manifest)

        # ── Emit gold standards ──────────────────────────────────────
        emit_all_golds(
            canaries, errors, output / "gold_standards",
            model=model,
        )

        # ── Write registries ─────────────────────────────────────────
        canaries.write_json(output / "canary_registry.json")
        errors.write_json(output / "error_registry.json")

        manifest.register("canary_registry.json", "json")
        manifest.register("error_registry.json", "json")

    return manifest


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        description="Generate the Cascade Industries test suite.",
    )
    parser.add_argument(
        "--config",
        default="config.yaml",
        help="Path to the configuration YAML file (default: config.yaml)",
    )
    parser.add_argument(
        "--output",
        required=True,
        help="Output directory for the generated test suite",
    )
    args = parser.parse_args(argv)

    try:
        config = load_config(args.config)
    except ConfigError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)

    output = Path(args.output)
    generate(config, output)


if __name__ == "__main__":
    main()
