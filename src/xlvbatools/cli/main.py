"""Small command-line entry point for xlvbatools."""

from __future__ import annotations

import json
from importlib import import_module
import sys
from typing import Sequence

from xlvbatools.cli.parser import build_parser


def main(args: Sequence[str] | None = None) -> None:
    """Parse arguments and dispatch one command handler."""
    if sys.platform == "win32":
        for stream in (sys.stdout, sys.stderr):
            reconfigure = getattr(stream, "reconfigure", None)
            if callable(reconfigure):
                reconfigure(encoding="utf-8")

    parser = build_parser()
    namespace = parser.parse_args(args)
    if getattr(namespace, "show_agents", False):
        from xlvbatools.cli.commands import _cmd_agents

        _cmd_agents(namespace)
        raise SystemExit(0)
    if namespace.command is None:
        parser.print_help()
        raise SystemExit(0)

    try:
        namespace.func(namespace)
    except Exception as error:
        is_com = _is_com_error(error)
        if getattr(namespace, "output_format", "json") == "json":
            from xlvbatools import OperationResult

            result = OperationResult.failed(
                namespace.command or "cli",
                error,
                phase="cli",
                code="excel_com_error" if is_com else "unhandled_cli_error",
            )
            print(json.dumps(result.to_dict(), indent=2, default=str))
            raise SystemExit(1) from error
        if is_com:
            sys.stderr.write(f"\n[Error] Excel COM automation failure: {error}\n")
            sys.stderr.write("Please check:\n")
            sys.stderr.write("  1. Excel is installed and can open the workbook.\n")
            sys.stderr.write(
                "  2. Trust Center allows access to the VBA project object model.\n"
            )
            sys.stderr.write("  3. No other Excel instance is blocking the file.\n")
            raise SystemExit(1) from error
        raise


def _is_com_error(error: Exception) -> bool:
    try:
        pywintypes = import_module("pywintypes")
    except ImportError:
        return False
    return isinstance(error, pywintypes.com_error)


if __name__ == "__main__":
    main()
