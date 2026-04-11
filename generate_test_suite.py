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
from generator.golds.framework import emit_all_golds
from generator.manifest import Manifest
from generator.model.build import build_model
from generator.packs import collect_canary_keys, collect_test_case_count, resolve_packs
from scoring.scoring_template import generate_scoring_template


def _seed_all(seed: int) -> None:
    """Seed every PRNG used by the generator for determinism."""
    random.seed(seed)
    np.random.seed(seed)
    Faker.seed(seed)


def _create_directory_tree(output: Path, tc_count: int) -> None:
    """Create the output directory structure per prompt.md §2.3."""
    output.mkdir(parents=True, exist_ok=True)

    # Top-level directories
    (output / "shared_data").mkdir(exist_ok=True)
    (output / "gold_standards").mkdir(exist_ok=True)
    (output / "scoring").mkdir(exist_ok=True)
    (output / "templates").mkdir(exist_ok=True)

    # Per-test-case directories with input_files subdirectory
    for i in range(1, tc_count + 1):
        tc_dir = output / "test_cases" / f"TC-{i:02d}"
        (tc_dir / "input_files").mkdir(parents=True, exist_ok=True)


def generate(
    config: Config, output: Path, pack_ids: list[str] | None = None,
) -> Manifest:
    """Run the full generation pipeline.

    Parameters
    ----------
    pack_ids : optional list of pack IDs to enable.  ``None`` (default)
        selects the default packs (accounting-core only).  Pass ``["all"]``
        to include every registered pack.

    Returns the Manifest so callers (tests) can inspect registered files.
    Model builders and formatters will be wired in by later beads.
    """
    _seed_all(config.seed)

    # ── Resolve packs ───────────────────────────────────────────────
    packs = resolve_packs(pack_ids)
    tc_count = collect_test_case_count(packs)
    canary_keys = collect_canary_keys(packs)

    _create_directory_tree(output, tc_count)

    # ── Phase 1: Build canonical model ──────────────────────────────
    model = build_model(config)

    # ── Canary & error registries ────────────────────────────────────
    canaries = build_canary_registry(canary_keys, seed=config.seed)
    errors = ErrorRegistry()

    with Manifest(output) as manifest:
        # ── Phase 2: Run pack emitters ───────────────────────────────
        for pack in packs:
            manifest.set_current_pack(pack.pack_id)
            for emitter in pack.emitters:
                emitter(model, output, canaries, errors, manifest)
        manifest.set_current_pack("")

        # ── Emit gold standards ──────────────────────────────────────
        active_tcs = [tc for pack in packs for tc in pack.test_cases]
        emit_all_golds(
            canaries, errors, output / "gold_standards",
            tc_ids=active_tcs,
            model=model,
        )

        # ── Scoring template ─────────────────────────────────────────
        rubrics_path = Path(__file__).resolve().parent / "scoring" / "rubrics.yaml"
        generate_scoring_template(
            rubrics_path, output / "scoring" / "scoring_template.xlsx",
        )
        manifest.register("scoring/scoring_template.xlsx", "xlsx")

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
