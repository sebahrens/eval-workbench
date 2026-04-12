#!/usr/bin/env python3
"""Cascade Industries test suite generator — orchestrator entry point.

Usage:
    python generate_test_suite.py --output /tmp/test_suite
    python generate_test_suite.py --config config.yaml --output /tmp/test_suite
    python generate_test_suite.py --output /tmp/test_suite --packs all
    python generate_test_suite.py --output /tmp/test_suite --packs cascade_legal_hr_diligence
    python generate_test_suite.py --output /tmp/test_suite --overlay small.yaml difficulty.yaml
    python generate_test_suite.py --output /tmp/test_suite --set company.name="Acme Corp" seed=99
"""

from __future__ import annotations

import argparse
import json
import random
import sys
from pathlib import Path

import numpy as np
from faker import Faker

from generator.canaries import build_registry as build_canary_registry
from generator.config import Config, ConfigError, load_config, load_layered_config
from generator.errors import ErrorRegistry
from generator.golds.framework import emit_all_golds
from generator.manifest import Manifest
from generator.model.build import build_model
from generator.model.fidelity_report import build_fidelity_report
from generator.packs import collect_canary_keys, collect_test_case_count, resolve_packs
from generator.scenario_context import ScenarioContext
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
    ctx = ScenarioContext(seed=config.seed)

    with Manifest(output) as manifest:
        # ── Phase 2: Run pack emitters ───────────────────────────────
        for pack in packs:
            manifest.set_current_pack(pack.pack_id)
            for emitter in pack.emitters:
                emitter(model, output, canaries, errors, manifest, ctx=ctx)
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

        # ── Fidelity regression report ──────────────────────────────
        fidelity = build_fidelity_report(model)
        (output / "fidelity_report.json").write_text(
            json.dumps(fidelity, indent=2, sort_keys=True) + "\n",
            encoding="utf-8",
        )
        manifest.register("fidelity_report.json", "json")

    return manifest


def parse_set_overrides(pairs: list[str]) -> dict:
    """Parse ``--set key=value`` pairs into a nested dict.

    Each *pair* is ``dotted.key=value``.  Numeric strings are coerced to
    ``int`` or ``float`` where possible; ``"true"``/``"false"`` become bools;
    ``"null"`` becomes ``None``.

    >>> parse_set_overrides(["company.name=Acme", "seed=99"])
    {'company': {'name': 'Acme'}, 'seed': 99}
    """
    result: dict = {}
    for pair in pairs:
        if "=" not in pair:
            raise ConfigError(f"--set value must be key=value, got {pair!r}")
        key, raw_value = pair.split("=", 1)
        if not key:
            raise ConfigError(f"--set key must not be empty in {pair!r}")

        value: object = _coerce_value(raw_value)

        # Build nested dict from dotted key path
        parts = key.split(".")
        node = result
        for part in parts[:-1]:
            node = node.setdefault(part, {})
        node[parts[-1]] = value
    return result


def _coerce_value(raw: str) -> object:
    """Coerce a CLI string value to the most specific Python type."""
    if raw.lower() == "null":
        return None
    if raw.lower() == "true":
        return True
    if raw.lower() == "false":
        return False
    try:
        return int(raw)
    except ValueError:
        pass
    try:
        return float(raw)
    except ValueError:
        pass
    return raw


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        description="Generate the Cascade Industries test suite.",
    )
    subparsers = parser.add_subparsers(dest="command")

    # ── configure subcommand (TUI) ──────────────────────────────────
    configure_parser = subparsers.add_parser(
        "configure",
        help="Launch the scenario configuration TUI (requires synth-data[tui]).",
    )
    configure_parser.add_argument(
        "--config",
        default="config.yaml",
        help="Path to the base configuration YAML file (default: config.yaml)",
    )
    configure_parser.add_argument(
        "--overlay",
        nargs="+",
        default=None,
        metavar="YAML",
        help="One or more overlay YAML files merged onto the base config in order.",
    )

    # ── generate (default) arguments ────────────────────────────────
    parser.add_argument(
        "--config",
        default="config.yaml",
        help="Path to the base configuration YAML file (default: config.yaml)",
    )
    parser.add_argument(
        "--overlay",
        nargs="+",
        default=None,
        metavar="YAML",
        help=(
            "One or more overlay YAML files merged onto the base config in order. "
            "Each file may contain any subset of v1 supported fields."
        ),
    )
    parser.add_argument(
        "--set",
        nargs="+",
        default=None,
        dest="set_overrides",
        metavar="KEY=VALUE",
        help=(
            "Leaf-level config overrides in dotted-key=value form. "
            "Applied after --overlay files. Example: --set company.name=Acme seed=99"
        ),
    )
    parser.add_argument(
        "--output",
        default=None,
        help="Output directory for the generated test suite",
    )
    parser.add_argument(
        "--packs",
        nargs="*",
        default=None,
        metavar="PACK",
        help=(
            "Scenario packs to generate. Omit for default (accounting-core only). "
            "Use 'all' to generate every registered pack. "
            "Available packs: cascade_accounting_core, cascade_legal_hr_diligence."
        ),
    )
    args = parser.parse_args(argv)

    # ── Dispatch subcommands ────────────────────────────────────────
    if args.command == "configure":
        from generator.tui.app import configure
        configure(config_path=args.config, overlay=args.overlay)
        return

    # ── Generate (default) ──────────────────────────────────────────
    if not args.output:
        parser.error("the following arguments are required: --output")

    try:
        has_layers = args.overlay or args.set_overrides
        if has_layers:
            set_dict = parse_set_overrides(args.set_overrides) if args.set_overrides else None
            config = load_layered_config(
                args.config,
                layers=args.overlay,
                set_overrides=set_dict,
            )
        else:
            config = load_config(args.config)
    except ConfigError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)

    output = Path(args.output)

    try:
        generate(config, output, pack_ids=args.packs)
    except ConfigError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
