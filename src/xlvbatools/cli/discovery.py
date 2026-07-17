"""Authoritative command and agent-onboarding discovery metadata."""

from __future__ import annotations

import argparse
from dataclasses import asdict, dataclass
from typing import Any


DISCOVERY_SCHEMA_VERSION = "1.0"


@dataclass(frozen=True)
class CommandSpec:
    """Stable discovery information for one public CLI command."""

    name: str
    summary: str
    usage: str
    examples: tuple[str, ...]
    excel_required: bool = False
    interactive: bool = False


COMMAND_SPECS = (
    CommandSpec(
        "help", "Return CLI discovery metadata", "xlvba help [command]",
        ("xlvba help", "xlvba help extract", "xlvba help --table"),
    ),
    CommandSpec(
        "init", "Initialize a configured xlvbatools project", "xlvba init [options]",
        ("xlvba init --workbook workbook/Model.xlsm", "xlvba init --agents"),
    ),
    CommandSpec(
        "extract", "Extract VBA source or list components", "xlvba extract [options]",
        ("xlvba extract", "xlvba extract --list --table"), True,
    ),
    CommandSpec(
        "inject", "Inject VBA source into a workbook", "xlvba inject [options]",
        ("xlvba inject --dry-run", "xlvba inject --component modMain"), True,
    ),
    CommandSpec(
        "diff", "Compare workbook VBA with source", "xlvba diff [options]",
        ("xlvba diff", "xlvba diff --summary --text"), True,
    ),
    CommandSpec(
        "lint", "Analyze extracted source or a live workbook", "xlvba lint [options]",
        ("xlvba lint --source vba_source", "xlvba lint --workbook workbook/Model.xlsm"),
    ),
    CommandSpec(
        "run", "Execute a VBA macro in an isolated worker", "xlvba run MACRO [options]",
        (
            "xlvba run OnCalculate --no-save --timeout 120",
            "xlvba run Module1.Refresh --named-range InputValue=42 --timeout 240",
        ),
        True,
    ),
    CommandSpec(
        "snapshot", "Create and manage workbook checkpoints",
        "xlvba snapshot [options] {create,list,info,restore,diff,prune}",
        ("xlvba snapshot create --desc \"before change\"", "xlvba snapshot list --table"),
    ),
    CommandSpec(
        "dump", "Inspect worksheet data and screenshots", "xlvba dump [options]",
        ("xlvba dump --sheets Input --data", "xlvba dump --sheets Input --screenshot"),
        True,
    ),
    CommandSpec(
        "modify", "Modify cells, formulas, or workbook names", "xlvba modify [options]",
        ("xlvba modify --sheet Input --cell C33 --value 42",), True,
    ),
    CommandSpec(
        "debug", "Open Excel and optionally the VBE visibly", "xlvba debug [options]",
        ("xlvba debug --workbook workbook/Model.xlsm",), True, True,
    ),
    CommandSpec(
        "search", "Search extracted VBA source", "xlvba search PATTERN [options]",
        ("xlvba search FileCount", "xlvba search \"Dim\\s+.*As\" --regex"),
    ),
    CommandSpec(
        "fmt", "Format extracted VBA source", "xlvba fmt [options]",
        ("xlvba fmt --dry-run", "xlvba fmt --source vba_source"),
    ),
    CommandSpec(
        "graph", "Generate a VBA dependency graph", "xlvba graph [options]",
        ("xlvba graph", "xlvba graph --graph-format mermaid --text"),
    ),
    CommandSpec(
        "agents", "Discover or install packaged agent guidance", "xlvba agents [install]",
        ("xlvba agents", "xlvba agents install", "xlvba agents install --force"),
    ),
    CommandSpec(
        "version", "Report installed package provenance", "xlvba version [options]",
        ("xlvba version", "xlvba version --text"),
    ),
)

COMMAND_INDEX = {spec.name: spec for spec in COMMAND_SPECS}


def command_epilog(name: str) -> str:
    """Return copy-ready examples for conventional argparse help."""
    examples = "\n".join(f"  {example}" for example in COMMAND_INDEX[name].examples)
    return f"Examples:\n{examples}\n\nOutput is JSON by default; use --text or --table explicitly."


def discovery_payload(command: str | None = None) -> dict[str, Any]:
    """Return versioned machine-readable CLI and template discovery data."""
    parser_metadata = _parser_metadata()
    if command is not None:
        try:
            spec = COMMAND_INDEX[command]
        except KeyError as error:
            choices = ", ".join(COMMAND_INDEX)
            raise ValueError(f"Unknown command {command!r}; choose one of: {choices}") from error
        command_data = asdict(spec)
        command_data.update(parser_metadata[command])
        return {
            "discovery_schema_version": DISCOVERY_SCHEMA_VERSION,
            "command": command_data,
            "output_contract": _output_contract(),
        }
    commands = []
    for spec in COMMAND_SPECS:
        command_data = asdict(spec)
        command_data.update(parser_metadata[spec.name])
        commands.append(command_data)
    return {
        "discovery_schema_version": DISCOVERY_SCHEMA_VERSION,
        "commands": commands,
        "output_contract": _output_contract(),
        "agent_templates": {
            "destination": ".agents/",
            "install_existing_project": "xlvba agents install",
            "initialize_new_project": "xlvba init --agents",
            "overwrite_packaged_files": "xlvba agents install --force",
            "read_first": ".agents/AGENTS.md",
            "behavior": "missing files are copied; existing files are preserved unless --force",
        },
    }


def _output_contract() -> dict[str, Any]:
    return {
        "default": "json",
        "schema": "OperationResult",
        "presentation_flags": ["--text", "--table", "--output-format text|table"],
        "interactive_exception": "debug",
    }


def _parser_metadata() -> dict[str, dict[str, Any]]:
    """Derive flags from argparse so discovery cannot drift from execution."""
    from xlvbatools.cli.parser import build_parser

    root = build_parser()
    subparsers = next(
        action
        for action in root._actions
        if isinstance(action, argparse._SubParsersAction)
    )
    return {
        name: {
            "options": _action_metadata(command_parser),
            "subcommands": _subcommand_metadata(command_parser),
        }
        for name, command_parser in subparsers.choices.items()
    }


def _action_metadata(parser: argparse.ArgumentParser) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    for action in parser._actions:
        if action.dest == "help" or isinstance(action, argparse._SubParsersAction):
            continue
        default = action.default
        if not isinstance(default, (str, int, float, bool, type(None))):
            default = str(default)
        items.append(
            {
                "name": action.dest,
                "flags": list(action.option_strings),
                "positional": not action.option_strings,
                "required": action.required,
                "nargs": action.nargs,
                "choices": list(action.choices) if action.choices is not None else None,
                "default": default,
                "description": action.help,
            }
        )
    return items


def _subcommand_metadata(parser: argparse.ArgumentParser) -> list[dict[str, Any]]:
    for action in parser._actions:
        if isinstance(action, argparse._SubParsersAction):
            return [
                {
                    "name": name,
                    "summary": subparser.description,
                    "options": _action_metadata(subparser),
                }
                for name, subparser in action.choices.items()
            ]
    return []
