"""Allow running the TUI via ``python -m generator.tui``."""

from __future__ import annotations

import argparse

from generator.tui.app import configure


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(
        prog="python -m generator.tui",
        description="Launch the scenario configuration TUI.",
    )
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
        help="One or more overlay YAML files merged onto the base config in order.",
    )
    args = parser.parse_args(argv)
    configure(config_path=args.config, overlay=args.overlay)


if __name__ == "__main__":
    main()
