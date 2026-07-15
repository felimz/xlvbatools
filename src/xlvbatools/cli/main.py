"""
xlvba -- Unified CLI for xlvbatools
====================================
Entry point for all xlvbatools CLI commands.

Usage:
    xlvba <command> [options]

Commands:
    init        Initialize xlvbatools.toml in current directory
    extract     Extract VBA from workbook to vba_source/
    inject      Inject vba_source/ into workbook
    diff        Compare workbook VBA vs. vba_source/
    lint        Run static analysis on VBA code
    run         Execute a VBA macro with dialog protection
    snapshot    Checkpoint and rollback system
    dump        Dump sheet data, screenshots, named ranges
    modify      Modify cell values, formulas, or named ranges
    debug       Open Excel + VBE visibly for interactive debugging
    search      Search VBA source files for a pattern
"""

import argparse
import json
import os
import sys
from typing import List, Optional, Any


def main(args: Optional[List[str]] = None) -> None:
    """Main entry point for the xlvba CLI."""
    if sys.platform == "win32":
        try:
            sys.stdout.reconfigure(encoding="utf-8")
            sys.stderr.reconfigure(encoding="utf-8")
        except AttributeError:
            pass  # In some testing environments stdout/stderr might not support reconfigure

    parser = argparse.ArgumentParser(
        prog="xlvba",
        description="General-purpose toolkit for headless Excel VBA automation",
    )
    parser.add_argument(
        "--version", action="version",
        version=f"%(prog)s {_get_version()}"
    )
    parser.add_argument(
        "--agents", dest="show_agents", action="store_true",
        help="Show AI agent integration help and best practices"
    )

    subparsers = parser.add_subparsers(dest="command", help="Available commands")

    # Register all subcommands
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
    _register_version(subparsers)
    _register_agents(subparsers)

    args = parser.parse_args(args)

    if getattr(args, "show_agents", False):
        _cmd_agents(args)
        sys.exit(0)

    if args.command is None:
        parser.print_help()
        sys.exit(0)

    # Dispatch to the command handler
    try:
        if hasattr(args, "func"):
            args.func(args)
        else:
            parser.print_help()
    except Exception as e:
        is_com = False
        try:
            import pywintypes
            if isinstance(e, pywintypes.com_error):
                is_com = True
        except ImportError:
            pass

        if is_com:
            sys.stderr.write(f"\n[Error] Excel COM automation failure: {e}\n")
            sys.stderr.write("Please check:\n")
            sys.stderr.write("  1. Excel is installed and can open the workbook.\n")
            sys.stderr.write("  2. Trust Center Macro Settings allow access to the VBA project object model.\n")
            sys.stderr.write("  3. No other instance of Excel is blocking the file.\n")
            sys.exit(1)
        else:
            raise


def _get_version() -> str:
    try:
        from xlvbatools import __version__
        return __version__
    except ImportError:
        return "unknown"


def _cfg_path(cfg, resolved_attr: str, raw_attr: str) -> str:
    """Use config-relative paths while tolerating lightweight test doubles."""
    resolved = getattr(cfg, resolved_attr, None)
    if isinstance(resolved, str):
        return resolved
    return os.fspath(getattr(cfg, raw_attr))


def _project(cfg, *, workbook=None, source=None):
    """Create the same configured Project used by Python consumers."""
    from dataclasses import replace
    from xlvbatools import Project, ProjectSettings

    resolved_workbook = os.path.abspath(
        os.fspath(workbook) if workbook else _cfg_path(cfg, "workbook_path", "workbook")
    )
    resolved_source = os.path.abspath(
        os.fspath(source) if source else _cfg_path(cfg, "vba_source_path", "vba_source")
    )
    configured = replace(
        cfg,
        workbook=resolved_workbook,
        vba_source=resolved_source,
        config_dir=None,
    )
    return Project(ProjectSettings._from_config(configured))


def _print_result_json(result) -> None:
    print(json.dumps(result.to_dict(), indent=2, default=str))


def _fail_result(result, fallback: str) -> None:
    message = result.error.message if result.error is not None else fallback
    print(message, file=sys.stderr)
    raise SystemExit(1)


# ── Subcommand registration ──
# Each registers its argparse sub-parser and sets the dispatch handler.


def _register_init(subparsers):
    p = subparsers.add_parser("init", help="Initialize xlvbatools.toml in current directory")
    p.add_argument("--workbook", "-w", help="Path to the .xlsm workbook")
    p.add_argument("--agents", action="store_true",
                   help="Install .agents/ templates for AI coding assistant rules, skills, and workflows")
    p.add_argument("--force", "-f", action="store_true", help="Overwrite existing xlvbatools.toml")
    p.set_defaults(func=_cmd_init)


def _register_extract(subparsers):
    p = subparsers.add_parser("extract", help="Extract VBA from workbook to vba_source/")
    p.add_argument("--workbook", "-w", help="Path to the .xlsm workbook")
    p.add_argument("--output", "-o", help="Output directory for extracted code")
    p.add_argument("--component", "-c", help="Extract a single component by name")
    p.add_argument("--list", "-l", action="store_true", help="List components only")
    p.add_argument("--json", action="store_true", help="JSON output")
    p.add_argument("--timeout", type=float, default=120.0,
                   help="Maximum seconds for the isolated Excel worker")
    p.add_argument("--verbose", "-v", action="store_true")
    p.set_defaults(func=_cmd_extract)


def _register_inject(subparsers):
    p = subparsers.add_parser("inject", help="Inject vba_source/ into workbook")
    p.add_argument("--workbook", "-w", help="Path to the .xlsm workbook")
    p.add_argument("--source", "-s", help="Source directory with VBA files")
    p.add_argument("--component", "-c", help="Inject a single component by name")
    p.add_argument("--dry-run", action="store_true")
    p.add_argument("--no-backup", action="store_true")
    p.add_argument("--json", action="store_true")
    p.add_argument("--timeout", type=float, default=120.0,
                   help="Maximum seconds for the isolated Excel worker")
    p.add_argument("--verbose", "-v", action="store_true")
    p.set_defaults(func=_cmd_inject)


def _register_diff(subparsers):
    p = subparsers.add_parser("diff", help="Compare workbook VBA vs. vba_source/")
    p.add_argument("--workbook", "-w", help="Path to the .xlsm workbook")
    p.add_argument("--source", "-s", help="Source directory with VBA files")
    p.add_argument("--component", "-c", help="Diff a single component by name")
    p.add_argument("--summary", action="store_true")
    p.add_argument("--json", action="store_true")
    p.add_argument("--timeout", type=float, default=120.0,
                   help="Maximum seconds for the isolated Excel worker")
    p.add_argument("--verbose", "-v", action="store_true")
    p.set_defaults(func=_cmd_diff)


def _register_lint(subparsers):
    p = subparsers.add_parser("lint", help="Run static analysis on VBA code")
    p.add_argument("--workbook", "-w", help="Path to the .xlsm workbook")
    p.add_argument(
        "--source", "-s",
        help="Path to a VBA source directory or one .bas/.cls/.frm file",
    )
    p.add_argument("--json", action="store_true")
    p.add_argument("--timeout", type=float, default=120.0,
                   help="Maximum seconds for workbook lint worker")
    p.add_argument("--verbose", "-v", action="store_true")
    p.set_defaults(func=_cmd_lint)


def _register_run(subparsers):
    p = subparsers.add_parser("run", help="Execute a VBA macro with dialog protection")
    p.add_argument("macro", help="Name of the macro to run")
    p.add_argument("--workbook", "-w", help="Path to the .xlsm workbook")
    p.add_argument("--json", action="store_true")
    p.add_argument("--timeout", type=float, default=120.0,
                   help="Maximum seconds for the isolated Excel session")
    p.add_argument("--verbose", "-v", action="store_true")
    p.set_defaults(func=_cmd_run)


def _register_snapshot(subparsers):
    p = subparsers.add_parser("snapshot", help="Checkpoint and rollback system")
    p.add_argument("--workbook", "-w", help="Path to the .xlsm workbook")
    p.add_argument("--source", "-s", help="Path to vba_source/ directory")
    sp = p.add_subparsers(dest="snapshot_command")

    c = sp.add_parser("create", help="Create a new snapshot")
    c.add_argument("--desc", "-d", dest="description", default="")
    c.add_argument("--milestone", "-m", action="store_true")
    c.set_defaults(func=_cmd_snapshot)

    sp.add_parser("list", help="List all snapshots").set_defaults(func=_cmd_snapshot)

    i = sp.add_parser("info", help="Show snapshot details")
    i.add_argument("identifier")
    i.set_defaults(func=_cmd_snapshot)

    r = sp.add_parser("restore", help="Restore from a snapshot")
    r.add_argument("identifier")
    r.add_argument("--no-safety", action="store_true")
    r.set_defaults(func=_cmd_snapshot)

    d = sp.add_parser("diff", help="Show changes since a snapshot")
    d.add_argument("identifier")
    d.set_defaults(func=_cmd_snapshot)

    pr = sp.add_parser("prune", help="Remove old snapshots")
    pr.add_argument("--keep", "-k", type=int, default=10)
    pr.set_defaults(func=_cmd_snapshot)

    p.set_defaults(func=_cmd_snapshot_help)


def _register_dump(subparsers):
    p = subparsers.add_parser("dump", help="Dump sheet data, screenshots, named ranges")
    p.add_argument("--workbook", "-w", help="Path to the .xlsm workbook")
    p.add_argument("--sheets", "-s", help="Comma-separated sheet names")
    p.add_argument("--screenshot", action="store_true")
    p.add_argument("--data", action="store_true")
    p.add_argument("--range", "-r", help="Specific cell range")
    p.add_argument("--json", action="store_true")
    p.add_argument("--timeout", type=float, default=60.0,
                   help="Hard timeout for the isolated inspection worker")
    p.add_argument("--include-hidden-sheets", action="store_true",
                   help="Explicitly render Hidden and VeryHidden worksheets")
    p.add_argument("--verbose", "-v", action="store_true")
    p.set_defaults(func=_cmd_dump)


def _register_modify(subparsers):
    p = subparsers.add_parser("modify", help="Modify cell values, formulas, or named ranges")
    p.add_argument("--workbook", "-w", help="Path to the .xlsm workbook")
    p.add_argument("--sheet", "-s", help="Worksheet name")
    p.add_argument("--cell", "-c", help="Cell coordinate (e.g. 'A1')")
    p.add_argument("--value", help="Value to write")
    p.add_argument("--formula", "-f", help="Formula to write")
    p.add_argument("--name", "-n", help="Named range")
    p.add_argument("--refers-to", help="Reference for named range")
    p.add_argument("--delete-name", "-d", action="store_true")
    p.add_argument("--timeout", type=float, default=120.0,
                   help="Maximum seconds for the isolated Excel worker")
    p.add_argument("--verbose", "-v", action="store_true")
    p.set_defaults(func=_cmd_modify)


def _register_debug(subparsers):
    p = subparsers.add_parser("debug", help="Open Excel + VBE visibly for interactive debugging")
    p.add_argument("--workbook", "-w", help="Path to the .xlsm workbook")
    p.add_argument("--no-vbe", action="store_true", help="Don't auto-open VBE")
    p.set_defaults(func=_cmd_debug)


def _register_search(subparsers):
    p = subparsers.add_parser("search", help="Search VBA source files for a pattern")
    p.add_argument("pattern", help="Search pattern (literal or regex)")
    p.add_argument("--source", "-s", help="Source directory to search")
    p.add_argument("--regex", "-r", action="store_true", help="Treat pattern as regex")
    p.add_argument("--context", "-C", type=int, default=0, help="Context lines")
    p.set_defaults(func=_cmd_search)


def _register_fmt(subparsers):
    p = subparsers.add_parser("fmt", help="Format VBA source code")
    p.add_argument("--source", "-s", help="Source directory or file")
    p.add_argument("--dry-run", action="store_true", help="Show diff without modifying")
    p.add_argument("--indent", type=int, default=4, help="Indent size (default 4)")
    p.set_defaults(func=_cmd_fmt)


def _register_graph(subparsers):
    p = subparsers.add_parser("graph", help="Generate VBA call dependency graph")
    p.add_argument("--source", "-s", help="Source directory")
    p.add_argument("--format", choices=["mermaid", "dot", "json"], default="mermaid",
                   help="Output format (default: mermaid)")
    p.add_argument("--output", "-o", help="Output file (default: stdout)")
    p.set_defaults(func=_cmd_graph)


def _register_agents(subparsers):
    p = subparsers.add_parser("agents", help="Show AI agent integration help and best practices")
    p.set_defaults(func=_cmd_agents)


def _register_version(subparsers):
    p = subparsers.add_parser(
        "version", help="Show package, interpreter, and source revision details"
    )
    p.add_argument("--json", action="store_true")
    p.set_defaults(func=_cmd_version)


# ── Command handlers ──

def _cmd_init(args):
    """Initialize xlvbatools.toml in the current directory."""
    from xlvbatools.cli.init_cmd import run_init
    run_init(args)


def _cmd_extract(args):
    """Extract VBA from workbook."""
    from xlvbatools.config.loader import load_config
    from xlvbatools.logging import setup_logging
    cfg = load_config()
    setup_logging(verbose=getattr(args, "verbose", False), tool_name="extract",
                  log_dir=_cfg_path(cfg, "log_dir_path", "log_dir"),
                  log_name=cfg.log_name)

    out = args.output or _cfg_path(cfg, "vba_source_path", "vba_source")
    project = _project(cfg, workbook=args.workbook, source=out)

    if getattr(args, "list", False):
        result = project.list_components(timeout=args.timeout)
        if not result.success:
            if getattr(args, "json", False):
                _print_result_json(result)
                raise SystemExit(1)
            _fail_result(result, "Component listing failed")
        components = result.data or []
        if getattr(args, "json", False):
            _print_result_json(result)
        else:
            for c in components:
                print(f"  {c['type_name']:20s} {c['name']:30s} ({c['line_count']} lines)")
        return

    result = project.extract(
        output=out, component=args.component, timeout=args.timeout,
    )
    if not result.success:
        if getattr(args, "json", False):
            _print_result_json(result)
            raise SystemExit(1)
        _fail_result(result, "Extraction failed")

    data = result.data
    if args.component:
        if getattr(args, "json", False):
            _print_result_json(result)
        else:
            print(f"Extracted: {data['name']} -> {data['file']}")
    else:
        manifest = data or {}
        count = len(manifest.get("components", []))
        if getattr(args, "json", False):
            _print_result_json(result)
        else:
            print(f"Extracted {count} components to {out}/")


def _cmd_inject(args):
    """Inject VBA into workbook."""
    from xlvbatools.config.loader import load_config
    from xlvbatools.logging import setup_logging
    cfg = load_config()
    setup_logging(verbose=getattr(args, "verbose", False), tool_name="inject",
                  log_dir=_cfg_path(cfg, "log_dir_path", "log_dir"),
                  log_name=cfg.log_name)

    src = args.source or _cfg_path(cfg, "vba_source_path", "vba_source")
    project = _project(cfg, workbook=args.workbook, source=src)
    if args.component and args.dry_run:
        print("--dry-run cannot be combined with --component", file=sys.stderr)
        sys.exit(2)
    result = project.inject(
        source=src,
        component=args.component,
        backup=not args.no_backup,
        dry_run=args.dry_run,
        timeout=args.timeout,
    )
    if not result.success:
        if getattr(args, "json", False):
            _print_result_json(result)
            raise SystemExit(1)
        _fail_result(result, "Injection failed")
    if args.component:
        if getattr(args, "json", False):
            _print_result_json(result)
        else:
            print("Injected")
    else:
        results = result.data or []
        if getattr(args, "json", False):
            _print_result_json(result)
        else:
            for r in results:
                status = r["status"].upper()
                print(f"  {status:10s} {r['name']}")


def _cmd_diff(args):
    """Diff workbook VBA vs source."""
    from xlvbatools.config.loader import load_config
    from xlvbatools.logging import setup_logging
    cfg = load_config()
    setup_logging(verbose=getattr(args, "verbose", False), tool_name="diff",
                  log_dir=_cfg_path(cfg, "log_dir_path", "log_dir"),
                  log_name=cfg.log_name)

    src = args.source or _cfg_path(cfg, "vba_source_path", "vba_source")
    project = _project(cfg, workbook=args.workbook, source=src)
    operation_result = project.diff(
        source=src, component=args.component, timeout=args.timeout,
    )
    if not operation_result.success:
        if getattr(args, "json", False):
            _print_result_json(operation_result)
            raise SystemExit(1)
        _fail_result(operation_result, "Diff failed")
    if args.component:
        result = operation_result.data
        if getattr(args, "json", False):
            _print_result_json(operation_result)
        else:
            _print_diff_result(result, args)
    else:
        results = operation_result.data or []
        if getattr(args, "json", False):
            _print_result_json(operation_result)
        else:
            for result in results:
                _print_diff_result(result, args)


def _print_diff_result(r, args):
    status = r["status"]
    if status == "identical":
        print(f"  = {r['name']}")
    elif status == "modified":
        print(f"  M {r['name']} (+{r['lines_added']}/-{r['lines_removed']})")
        if not getattr(args, "summary", False) and "unified_diff" in r:
            print(r["unified_diff"])
    else:
        print(f"  ? {r['name']} ({status})")


def _cmd_lint(args):
    """Run static analysis."""
    from xlvbatools.config.loader import load_config
    from xlvbatools.logging import setup_logging
    from xlvbatools.analysis.preflight import print_report
    cfg = load_config()
    setup_logging(verbose=getattr(args, "verbose", False), tool_name="lint",
                  log_dir=_cfg_path(cfg, "log_dir_path", "log_dir"),
                  log_name=cfg.log_name)

    source = args.source or _cfg_path(cfg, "vba_source_path", "vba_source")
    project = _project(cfg, workbook=args.workbook, source=source)
    try:
        if args.workbook or (not args.source and not os.path.exists(source)):
            result = project.lint_workbook(timeout=args.timeout)
        else:
            result = project.lint_source(source)
    except (FileNotFoundError, OSError, ValueError) as error:
        print(f"ERROR: {error}", file=sys.stderr)
        sys.exit(2)

    issues = result.data or ()

    if getattr(args, "json", False):
        _print_result_json(result)
    else:
        print(print_report(issues))

    sys.exit(0 if result.success else 1)


def _cmd_run(args):
    """Run a VBA macro."""
    from xlvbatools.config.loader import load_config
    from xlvbatools.logging import setup_logging
    cfg = load_config()
    setup_logging(verbose=getattr(args, "verbose", False), tool_name="run",
                  log_dir=_cfg_path(cfg, "log_dir_path", "log_dir"),
                  log_name=cfg.log_name)

    project = _project(cfg, workbook=args.workbook)
    result = project.run(args.macro, timeout=args.timeout)

    if getattr(args, "json", False):
        _print_result_json(result)
    else:
        status = "OK" if result.success else "FAILED"
        elapsed = result.elapsed_seconds or 0
        print(f"{status} ({elapsed:.2f}s)")
        if result.error:
            print(f"  Error: {result.error.message}")
        for ev in result.diagnostics.dialog_events:
            print(f"  [{ev.get('type')}] {ev.get('text', '')}")

    sys.exit(0 if result.success else 1)


def _cmd_snapshot(args):
    """Handle snapshot subcommands."""
    from xlvbatools.config.loader import load_config
    from xlvbatools.logging import setup_logging
    cfg = load_config()
    setup_logging(
        tool_name="snapshot",
        log_dir=_cfg_path(cfg, "log_dir_path", "log_dir"),
        log_name=cfg.log_name,
    )

    wb = getattr(args, "workbook", None) or _cfg_path(
        cfg, "workbook_path", "workbook"
    )
    src = getattr(args, "source", None) or _cfg_path(
        cfg, "vba_source_path", "vba_source"
    )

    from xlvbatools.snapshot.manager import SnapshotManager
    mgr = SnapshotManager(
        wb, src, _cfg_path(cfg, "snapshots_path", "snapshots_dir"),
        rolling_limit=cfg.snapshots.rolling_limit,
    )

    sub = args.snapshot_command
    if sub == "create":
        sid = mgr.create(description=args.description, milestone=args.milestone)
        print(f"Snapshot created: {sid}")
    elif sub == "list":
        for s in mgr.list():
            milestone = " [MILESTONE]" if s.get("milestone") else ""
            print(f"  {s['snapshot_id']}  {s.get('description', '')}{milestone}")
    elif sub == "info":
        info = mgr.info(args.identifier)
        if info:
            import json
            print(json.dumps(info, indent=2, default=str))
        else:
            print(f"Not found: {args.identifier}")
    elif sub == "restore":
        success = mgr.restore(args.identifier, safety_snapshot=not args.no_safety)
        print("Restored" if success else "FAILED")
        sys.exit(0 if success else 1)
    elif sub == "diff":
        print(mgr.diff(args.identifier))
    elif sub == "prune":
        pruned = mgr.prune(keep=args.keep)
        print(f"Pruned {pruned} snapshots")
    else:
        print("Usage: xlvba snapshot <create|list|info|restore|diff|prune>")


def _cmd_dump(args):
    """Dump sheet data and screenshots."""
    from xlvbatools.config.loader import load_config
    from xlvbatools.logging import setup_logging
    cfg = load_config()
    setup_logging(verbose=getattr(args, "verbose", False), tool_name="dump",
                  log_dir=_cfg_path(cfg, "log_dir_path", "log_dir"),
                  log_name=cfg.log_name)

    sheets = args.sheets.split(",") if args.sheets else None

    if not sheets:
        print("Specify --sheets for workbook inspection")
        sys.exit(1)
    project = _project(cfg, workbook=args.workbook)
    include_screenshots = getattr(args, "screenshot", False)
    include_data = getattr(args, "data", False) or not include_screenshots
    out_json = "dump.json" if include_data and getattr(args, "json", False) else None
    out_md = "dump.md" if include_data and not getattr(args, "json", False) else None
    result = project.inspect(
        sheets, output_dir="screenshots", cell_range=args.range,
        include_data=include_data, include_screenshots=include_screenshots,
        output_json=out_json, output_markdown=out_md, timeout=args.timeout,
        include_hidden_sheets=args.include_hidden_sheets,
    )
    if not result.success:
        _print_result_json(result)
        sys.exit(1)
    if getattr(args, "json", False):
        _print_result_json(result)
    else:
        for name, path in result.data.screenshots.items():
            print(f"  {name}: {path}")
        if include_data:
            print(f"Dumped to: {out_md}")


def _cmd_modify(args):
    """Modify cell values or named ranges."""
    from xlvbatools.config.loader import load_config
    from xlvbatools.logging import setup_logging
    cfg = load_config()
    setup_logging(verbose=getattr(args, "verbose", False), tool_name="modify",
                  log_dir=_cfg_path(cfg, "log_dir_path", "log_dir"),
                  log_name=cfg.log_name)

    # Parse value with type detection
    value = args.value
    if value is not None:
        try:
            value = float(value)
            if value == int(value):
                value = int(value)
        except ValueError:
            pass

    project = _project(cfg, workbook=args.workbook)
    result = project.modify(
        sheet=args.sheet or "Sheet1",
        cell=args.cell,
        value=value,
        formula=args.formula,
        name=args.name,
        refers_to=args.refers_to,
        delete_name=args.delete_name,
        timeout=args.timeout,
    )
    print("OK" if result.success else "FAILED")
    if result.error:
        print(f"  Error: {result.error.message}")
    sys.exit(0 if result.success else 1)


def _cmd_debug(args):
    """Launch interactive debug session."""
    from xlvbatools.config.loader import load_config
    cfg = load_config()
    wb = args.workbook or _cfg_path(cfg, "workbook_path", "workbook")

    from xlvbatools.workbook.debugger import launch_debug_session
    launch_debug_session(wb, open_vbe=not args.no_vbe)


def _cmd_search(args):
    """Search VBA source files."""
    from xlvbatools.config.loader import load_config
    from xlvbatools.logging import setup_logging
    cfg = load_config()
    setup_logging(
        tool_name="search",
        log_dir=_cfg_path(cfg, "log_dir_path", "log_dir"),
        log_name=cfg.log_name,
    )

    src = args.source or _cfg_path(cfg, "vba_source_path", "vba_source")
    ctx = getattr(args, "context", 0)
    from xlvbatools.vba.search import search_vba
    matches = search_vba(src, args.pattern, regex=args.regex, context_lines=ctx)

    if not matches:
        print(f"No matches for: {args.pattern}")
        sys.exit(0)

    for m in matches:
        if m.context_before:
            for cl in m.context_before:
                print(f"  {m.file}-         {cl}")
        print(f"  {m.file}:{m.line_num}: {m.line}")
        if m.context_after:
            for cl in m.context_after:
                print(f"  {m.file}-         {cl}")
            print("  --")
    print(f"\n{len(matches)} match(es)")


def _cmd_fmt(args):
    """Format VBA source files."""
    from xlvbatools.config.loader import load_config
    cfg = load_config()
    src = args.source or _cfg_path(cfg, "vba_source_path", "vba_source")

    if os.path.isfile(src):
        from xlvbatools.vba.formatter import format_file
        result = format_file(src, dry_run=args.dry_run, indent_size=args.indent)
        if result.get("changed"):
            if args.dry_run:
                print(result.get("diff", ""))
            else:
                print(f"Formatted: {src} ({result['lines_changed']} lines changed)")
        else:
            print(f"No changes: {src}")
    elif os.path.isdir(src):
        from xlvbatools.vba.formatter import format_directory
        results = format_directory(src, dry_run=args.dry_run, indent_size=args.indent)
        changed = [r for r in results if r.get("changed")]
        for r in changed:
            if args.dry_run:
                print(r.get("diff", ""))
            else:
                print(f"  Formatted: {r['file']} ({r['lines_changed']} lines)")
        print(f"\n{len(changed)}/{len(results)} file(s) {'would be ' if args.dry_run else ''}changed")
    else:
        print(f"Not found: {src}")
        sys.exit(1)


def _cmd_graph(args):
    """Generate VBA call dependency graph."""
    from xlvbatools.config.loader import load_config
    cfg = load_config()
    src = args.source or _cfg_path(cfg, "vba_source_path", "vba_source")

    from xlvbatools.vba.dependency import build_call_graph, render_mermaid, render_dot
    graph = build_call_graph(src)

    fmt = getattr(args, "format", "mermaid")
    if fmt == "mermaid":
        output = render_mermaid(graph)
    elif fmt == "dot":
        output = render_dot(graph)
    else:
        import json
        output = json.dumps(graph.to_dict(), indent=2)

    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(output)
        print(f"Graph written to: {args.output}")
    else:
        print(output)

    print(f"\n{len(graph.procedures)} procedures, {len(graph.edges)} call edges")


def _cmd_stub(args):
    """Placeholder for commands not yet implemented."""
    cmd = args.command
    if hasattr(args, "snapshot_command") and args.snapshot_command:
        cmd = f"snapshot {args.snapshot_command}"
    print(f"xlvba {cmd}: not yet implemented (coming in a future phase)")
    sys.exit(0)


def _cmd_snapshot_help(args):
    if not hasattr(args, "snapshot_command") or args.snapshot_command is None:
        print("Usage: xlvba snapshot <create|list|info|restore|diff|prune>")
        sys.exit(0)


def _cmd_version(args):
    """Show exact installed package and source provenance."""
    from xlvbatools.version import get_version_info

    info = get_version_info()
    if getattr(args, "json", False):
        print(json.dumps(info.to_dict(), indent=2))
        return
    print(f"xlvba {info.version}")
    print(f"Python: {info.python_executable}")
    print(f"Package: {info.package_path}")
    print(f"Source: {info.source_url or 'unknown'}")
    print(f"Commit: {info.commit_id or 'unknown'}")
    if info.requested_revision:
        print(f"Requested revision: {info.requested_revision}")


def _cmd_agents(args):
    """Print AI agent integration instructions and best practices."""
    help_text = """xlvbatools v1 agent integration

Application automation uses xlvbatools.Project or the xlvba CLI. Raw COM
sessions and worker files are private implementation details.

Install guidance into a new project:
    xlvba init --workbook workbook/MyProject.xlsm --agents

Template installation is non-destructive: an existing .agents/ directory is
left unchanged. Keep customized guidance under source control.

Read in this order:
    .agents/AGENTS.md
    .agents/skills/xlvba-toolchain/SKILL.md
    .agents/rules/python-rules.md or vba-rules.md
    .agents/workflows/vba-edit.md or vba-debug.md

Safety requirements:
  - never terminate Excel globally;
  - require operation success and clean owned-process shutdown;
  - render hidden worksheets only when explicitly requested;
  - use the repository .venv and pin exact released versions or revisions.
"""
    print(help_text)


if __name__ == "__main__":
    main()
