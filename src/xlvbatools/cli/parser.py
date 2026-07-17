"""Argument parser construction, kept separate from command execution."""

from __future__ import annotations

import argparse
import json
from typing import Any

from xlvbatools import __version__
from xlvbatools.cli import commands
from xlvbatools.cli.discovery import COMMAND_INDEX, command_epilog


_TOP_LEVEL_EPILOG = """Agent and automation discovery:
  xlvba help                  machine-readable command catalog
  xlvba help extract          machine-readable detail for one command
  xlvba agents               agent integration contract
  xlvba agents install       safely install packaged guidance into .agents/

Every non-interactive command emits an OperationResult JSON envelope by
default. Use --text or --table only when presentation output is wanted.
Run 'xlvba COMMAND --help' for options and copy-ready examples.
"""


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="xlvba",
        description="Reliable, isolated Excel/VBA automation for tools and agents",
        epilog=_TOP_LEVEL_EPILOG,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--version",
        action="version",
        version=json.dumps({"version": __version__}),
        help="Print the installed version as JSON and exit",
    )
    parser.add_argument(
        "--agents",
        dest="show_agents",
        action="store_true",
        help="Shortcut for 'xlvba agents'",
    )
    subparsers = parser.add_subparsers(
        dest="command",
        title="commands",
        metavar="COMMAND",
        help="Run 'xlvba help' for machine-readable discovery",
    )
    _register_help(subparsers)
    _register_init(subparsers)
    _register_extract(subparsers)
    _register_inject(subparsers)
    _register_diff(subparsers)
    _register_lint(subparsers)
    _register_run(subparsers)
    _register_snapshot(subparsers)
    _register_dump(subparsers)
    _register_modify(subparsers)
    _register_debug(subparsers)
    _register_search(subparsers)
    _register_fmt(subparsers)
    _register_graph(subparsers)
    _register_agents(subparsers)
    _register_version(subparsers)
    return parser


def _command_parser(subparsers: Any, name: str) -> argparse.ArgumentParser:
    spec = COMMAND_INDEX[name]
    return subparsers.add_parser(
        name,
        help=spec.summary,
        description=spec.summary,
        epilog=command_epilog(name),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )


def _worker_options(parser: argparse.ArgumentParser, *, timeout: float = 120.0) -> None:
    parser.add_argument("--workbook", "-w", help="Workbook path; overrides configuration")
    parser.add_argument(
        "--timeout",
        type=float,
        default=timeout,
        help=f"Maximum isolated-worker runtime in seconds (default: {timeout:g})",
    )
    parser.add_argument(
        "--verbose", "-v", action="store_true", help="Enable verbose diagnostic logging"
    )


def _presentation_options(parser: argparse.ArgumentParser) -> None:
    """Add the common machine-first output contract to one command."""
    group = parser.add_mutually_exclusive_group()
    group.add_argument(
        "--output-format",
        choices=("json", "text", "table"),
        default="json",
        help="Output presentation (default: json OperationResult envelope)",
    )
    group.add_argument(
        "--text",
        dest="output_format",
        action="store_const",
        const="text",
        help="Print concise text instead of JSON",
    )
    group.add_argument(
        "--table",
        dest="output_format",
        action="store_const",
        const="table",
        help="Print an aligned table instead of JSON",
    )


def _register_help(subparsers: Any) -> None:
    parser = _command_parser(subparsers, "help")
    parser.add_argument(
        "target", nargs="?", choices=tuple(COMMAND_INDEX), help="Optional command to describe"
    )
    _presentation_options(parser)
    parser.set_defaults(func=commands._cmd_help)


def _register_init(subparsers: Any) -> None:
    parser = _command_parser(subparsers, "init")
    _presentation_options(parser)
    parser.add_argument("--workbook", "-w", help="Path written to xlvbatools.toml")
    parser.add_argument(
        "--agents", action="store_true", help="Also copy packaged guidance into .agents/"
    )
    parser.add_argument(
        "--force", "-f", action="store_true", help="Overwrite an existing xlvbatools.toml"
    )
    parser.set_defaults(func=commands._cmd_init)


def _register_extract(subparsers: Any) -> None:
    parser = _command_parser(subparsers, "extract")
    _worker_options(parser)
    parser.add_argument("--output", "-o", help="Destination source directory")
    parser.add_argument("--component", "-c", help="Extract only this VBA component")
    parser.add_argument("--list", "-l", action="store_true", help="List components only")
    _presentation_options(parser)
    parser.set_defaults(func=commands._cmd_extract)


def _register_inject(subparsers: Any) -> None:
    parser = _command_parser(subparsers, "inject")
    _worker_options(parser)
    parser.add_argument("--source", "-s", help="VBA source directory")
    parser.add_argument("--component", "-c", help="Inject only this VBA component")
    parser.add_argument("--dry-run", action="store_true", help="Report changes without writing")
    parser.add_argument("--no-backup", action="store_true", help="Do not create a workbook backup")
    _presentation_options(parser)
    parser.set_defaults(func=commands._cmd_inject)


def _register_diff(subparsers: Any) -> None:
    parser = _command_parser(subparsers, "diff")
    _worker_options(parser)
    parser.add_argument("--source", "-s", help="VBA source directory")
    parser.add_argument("--component", "-c", help="Compare only this VBA component")
    parser.add_argument("--summary", action="store_true", help="Suppress detailed line differences")
    _presentation_options(parser)
    parser.set_defaults(func=commands._cmd_diff)


def _register_lint(subparsers: Any) -> None:
    parser = _command_parser(subparsers, "lint")
    _worker_options(parser)
    parser.add_argument("--source", "-s", help="Source directory or one VBA source file")
    _presentation_options(parser)
    parser.set_defaults(func=commands._cmd_lint)


def _register_run(subparsers: Any) -> None:
    parser = _command_parser(subparsers, "run")
    _worker_options(parser)
    parser.add_argument("macro", help="Public VBA macro name, optionally module-qualified")
    parser.add_argument(
        "--named-range",
        action="append",
        metavar="NAME=VALUE",
        help=(
            "Set one named-range input; repeat as needed. Valid JSON values "
            "are typed, otherwise VALUE is text"
        ),
    )
    parser.add_argument(
        "--save",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Save workbook changes after the macro (default: save)",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Show the isolated owned Excel instance while the macro runs",
    )
    _presentation_options(parser)
    parser.set_defaults(func=commands._cmd_run)


def _register_snapshot(subparsers: Any) -> None:
    parser = _command_parser(subparsers, "snapshot")
    parser.add_argument("--workbook", "-w", help="Workbook path; overrides configuration")
    parser.add_argument("--source", "-s", help="VBA source directory; overrides configuration")
    nested = parser.add_subparsers(dest="snapshot_command", title="snapshot commands")
    create = nested.add_parser("create", help="Create a checkpoint")
    create.add_argument("--desc", "-d", dest="description", default="", help="Description")
    create.add_argument("--milestone", "-m", action="store_true", help="Retain as a milestone")
    _presentation_options(create)
    create.set_defaults(func=commands._cmd_snapshot)
    listing = nested.add_parser("list", help="List checkpoints")
    _presentation_options(listing)
    listing.set_defaults(func=commands._cmd_snapshot)
    info = nested.add_parser("info", help="Describe one checkpoint")
    info.add_argument("identifier", help="Snapshot identifier or 'latest'")
    _presentation_options(info)
    info.set_defaults(func=commands._cmd_snapshot)
    restore = nested.add_parser("restore", help="Restore one checkpoint")
    restore.add_argument("identifier", help="Snapshot identifier or 'latest'")
    restore.add_argument("--no-safety", action="store_true", help="Skip pre-restore safety snapshot")
    _presentation_options(restore)
    restore.set_defaults(func=commands._cmd_snapshot)
    diff = nested.add_parser("diff", help="Compare with one checkpoint")
    diff.add_argument("identifier", help="Snapshot identifier or 'latest'")
    _presentation_options(diff)
    diff.set_defaults(func=commands._cmd_snapshot)
    prune = nested.add_parser("prune", help="Remove old rolling checkpoints")
    prune.add_argument("--keep", "-k", type=int, default=10, help="Rolling snapshots to keep")
    _presentation_options(prune)
    prune.set_defaults(func=commands._cmd_snapshot)
    parser.set_defaults(func=commands._cmd_snapshot_help)


def _register_dump(subparsers: Any) -> None:
    parser = _command_parser(subparsers, "dump")
    _worker_options(parser, timeout=60.0)
    parser.add_argument("--sheets", "-s", help="Comma-separated worksheet names")
    parser.add_argument("--screenshot", action="store_true", help="Render worksheet screenshots")
    parser.add_argument("--data", action="store_true", help="Include worksheet cell data")
    parser.add_argument("--range", "-r", help="Optional cell range such as B2:K20")
    parser.add_argument("--write-json", help="Also write workbook data to this JSON path")
    parser.add_argument("--write-markdown", help="Also write workbook data to this Markdown path")
    _presentation_options(parser)
    parser.add_argument(
        "--include-hidden-sheets",
        action="store_true",
        help="Include Hidden and VeryHidden sheets (excluded by default)",
    )
    parser.set_defaults(func=commands._cmd_dump)


def _register_modify(subparsers: Any) -> None:
    parser = _command_parser(subparsers, "modify")
    _worker_options(parser)
    parser.add_argument("--sheet", "-s", help="Worksheet name")
    parser.add_argument("--cell", "-c", help="Cell address")
    parser.add_argument("--value", help="Literal value to write")
    parser.add_argument("--formula", "-f", help="Excel formula to write")
    parser.add_argument("--name", "-n", help="Workbook name to create or update")
    parser.add_argument("--refers-to", help="Reference used with --name")
    parser.add_argument("--delete-name", "-d", action="store_true", help="Delete --name")
    _presentation_options(parser)
    parser.set_defaults(func=commands._cmd_modify)


def _register_debug(subparsers: Any) -> None:
    parser = _command_parser(subparsers, "debug")
    parser.add_argument("--workbook", "-w", help="Workbook path; overrides configuration")
    parser.add_argument("--no-vbe", action="store_true", help="Open Excel without opening the VBE")
    parser.set_defaults(func=commands._cmd_debug)


def _register_search(subparsers: Any) -> None:
    parser = _command_parser(subparsers, "search")
    parser.add_argument("pattern", help="Literal text or regular expression")
    parser.add_argument("--source", "-s", help="VBA source directory")
    parser.add_argument("--regex", "-r", action="store_true", help="Interpret pattern as regex")
    parser.add_argument("--context", "-C", type=int, default=0, help="Context lines per match")
    _presentation_options(parser)
    parser.set_defaults(func=commands._cmd_search)


def _register_fmt(subparsers: Any) -> None:
    parser = _command_parser(subparsers, "fmt")
    parser.add_argument("--source", "-s", help="VBA source directory")
    parser.add_argument("--dry-run", action="store_true", help="Report changes without writing")
    parser.add_argument("--indent", type=int, default=4, help="Spaces per indentation level")
    _presentation_options(parser)
    parser.set_defaults(func=commands._cmd_fmt)


def _register_graph(subparsers: Any) -> None:
    parser = _command_parser(subparsers, "graph")
    parser.add_argument("--source", "-s", help="VBA source directory")
    parser.add_argument(
        "--graph-format",
        choices=("json", "mermaid", "dot"),
        default="json",
        help="Graph payload representation (default: json)",
    )
    parser.add_argument("--output", "-o", help="Also write the graph payload to this path")
    _presentation_options(parser)
    parser.set_defaults(func=commands._cmd_graph)


def _register_agents(subparsers: Any) -> None:
    parser = _command_parser(subparsers, "agents")
    _presentation_options(parser)
    nested = parser.add_subparsers(dest="agents_command", title="agent commands")
    install = nested.add_parser(
        "install",
        help="Copy packaged guidance into a project",
        description="Safely copy packaged guidance into a project's .agents/ directory",
    )
    install.add_argument(
        "--destination", default=".agents", help="Destination directory (default: .agents/)"
    )
    install.add_argument(
        "--force", action="store_true", help="Overwrite packaged files that already exist"
    )
    _presentation_options(install)
    install.set_defaults(func=commands._cmd_agents)
    parser.set_defaults(func=commands._cmd_agents)


def _register_version(subparsers: Any) -> None:
    parser = _command_parser(subparsers, "version")
    _presentation_options(parser)
    parser.set_defaults(func=commands._cmd_version)
