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
        "--agents", action="store_true",
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
    _register_agents(subparsers)

    args = parser.parse_args(args)

    if getattr(args, "agents", False):
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
    p.add_argument("--verbose", "-v", action="store_true")
    p.set_defaults(func=_cmd_inject)


def _register_diff(subparsers):
    p = subparsers.add_parser("diff", help="Compare workbook VBA vs. vba_source/")
    p.add_argument("--workbook", "-w", help="Path to the .xlsm workbook")
    p.add_argument("--source", "-s", help="Source directory with VBA files")
    p.add_argument("--component", "-c", help="Diff a single component by name")
    p.add_argument("--summary", action="store_true")
    p.add_argument("--json", action="store_true")
    p.add_argument("--verbose", "-v", action="store_true")
    p.set_defaults(func=_cmd_diff)


def _register_lint(subparsers):
    p = subparsers.add_parser("lint", help="Run static analysis on VBA code")
    p.add_argument("--workbook", "-w", help="Path to the .xlsm workbook")
    p.add_argument("--source", "-s", help="Path to vba_source/ directory")
    p.add_argument("--json", action="store_true")
    p.add_argument("--verbose", "-v", action="store_true")
    p.set_defaults(func=_cmd_lint)


def _register_run(subparsers):
    p = subparsers.add_parser("run", help="Execute a VBA macro with dialog protection")
    p.add_argument("macro", help="Name of the macro to run")
    p.add_argument("--workbook", "-w", help="Path to the .xlsm workbook")
    p.add_argument("--json", action="store_true")
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
                  log_dir=cfg.log_dir, log_name=cfg.log_name)

    wb = args.workbook or cfg.workbook
    out = args.output or cfg.vba_source

    if getattr(args, "list", False):
        from xlvbatools.vba.extractor import list_components
        components = list_components(wb)
        if getattr(args, "json", False):
            import json
            print(json.dumps(components, indent=2))
        else:
            for c in components:
                print(f"  {c['type_name']:20s} {c['name']:30s} ({c['line_count']} lines)")
        return

    if args.component:
        from xlvbatools.vba.extractor import extract_component
        result = extract_component(wb, args.component, out)
        if result:
            print(f"Extracted: {result['name']} -> {result['file']}")
        else:
            print(f"Component not found: {args.component}")
            sys.exit(1)
    else:
        from xlvbatools.vba.extractor import extract_all
        manifest = extract_all(wb, out)
        count = len(manifest.get("components", []))
        print(f"Extracted {count} components to {out}/")
        if getattr(args, "json", False):
            import json
            print(json.dumps(manifest, indent=2))


def _cmd_inject(args):
    """Inject VBA into workbook."""
    from xlvbatools.config.loader import load_config
    from xlvbatools.logging import setup_logging
    cfg = load_config()
    setup_logging(verbose=getattr(args, "verbose", False), tool_name="inject",
                  log_dir=cfg.log_dir, log_name=cfg.log_name)

    wb = args.workbook or cfg.workbook
    src = args.source or cfg.vba_source

    if args.component:
        from xlvbatools.vba.injector import inject_component
        success = inject_component(wb, src, args.component, backup=not args.no_backup)
        print("Injected" if success else "FAILED")
        sys.exit(0 if success else 1)
    else:
        from xlvbatools.vba.injector import inject_all
        results = inject_all(wb, src, backup=not args.no_backup, dry_run=args.dry_run,
                             backup_limit=cfg.backups.limit)
        if getattr(args, "json", False):
            import json
            print(json.dumps(results, indent=2))
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
                  log_dir=cfg.log_dir, log_name=cfg.log_name)

    wb = args.workbook or cfg.workbook
    src = args.source or cfg.vba_source

    if args.component:
        from xlvbatools.vba.differ import diff_component
        result = diff_component(wb, src, args.component)
        if result is None:
            print(f"Component not found: {args.component}")
            sys.exit(1)
        _print_diff_result(result, args)
    else:
        from xlvbatools.vba.differ import diff_all
        results = diff_all(wb, src)
        if getattr(args, "json", False):
            import json
            print(json.dumps(results, indent=2))
        else:
            for r in results:
                _print_diff_result(r, args)


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
                  log_dir=cfg.log_dir, log_name=cfg.log_name)

    disabled = cfg.lint.disabled_rules

    if args.source:
        from xlvbatools.analysis.preflight import lint_files
        issues = lint_files(args.source, disabled_rules=disabled)
    elif args.workbook:
        from xlvbatools.analysis.preflight import lint_workbook
        issues = lint_workbook(args.workbook, disabled_rules=disabled)
    else:
        # Try source dir first (no COM needed), then workbook
        src = cfg.vba_source
        if os.path.isdir(src):
            from xlvbatools.analysis.preflight import lint_files
            issues = lint_files(src, disabled_rules=disabled)
        else:
            from xlvbatools.analysis.preflight import lint_workbook
            issues = lint_workbook(cfg.workbook, disabled_rules=disabled)

    if getattr(args, "json", False):
        import json
        print(json.dumps([i.to_dict() for i in issues], indent=2))
    else:
        print(print_report(issues))

    errors = [i for i in issues if i.severity == "ERROR"]
    sys.exit(1 if errors else 0)


def _cmd_run(args):
    """Run a VBA macro."""
    from xlvbatools.config.loader import load_config
    from xlvbatools.logging import setup_logging
    cfg = load_config()
    setup_logging(verbose=getattr(args, "verbose", False), tool_name="run",
                  log_dir=cfg.log_dir, log_name=cfg.log_name)

    wb = args.workbook or cfg.workbook
    from xlvbatools.macro.runner import run_macro
    result = run_macro(wb, args.macro)

    if getattr(args, "json", False):
        import json
        print(json.dumps(result, indent=2, default=str))
    else:
        status = "OK" if result.get("success") else "FAILED"
        elapsed = result.get("elapsed_seconds", 0)
        print(f"{status} ({elapsed:.2f}s)")
        if result.get("error"):
            print(f"  Error: {result['error']}")
        for ev in result.get("dialog_events", []):
            print(f"  [{ev.get('type')}] {ev.get('text', '')}")

    sys.exit(0 if result.get("success") else 1)


def _cmd_snapshot(args):
    """Handle snapshot subcommands."""
    from xlvbatools.config.loader import load_config
    from xlvbatools.logging import setup_logging
    cfg = load_config()
    setup_logging(tool_name="snapshot", log_dir=cfg.log_dir, log_name=cfg.log_name)

    wb = getattr(args, "workbook", None) or cfg.workbook
    src = getattr(args, "source", None) or cfg.vba_source

    from xlvbatools.snapshot.manager import SnapshotManager
    mgr = SnapshotManager(
        wb, src, cfg.snapshots_dir,
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
                  log_dir=cfg.log_dir, log_name=cfg.log_name)

    wb = args.workbook or cfg.workbook
    sheets = args.sheets.split(",") if args.sheets else None

    if getattr(args, "screenshot", False) and sheets:
        from xlvbatools.workbook.dumper import export_screenshots
        results = export_screenshots(wb, sheets, "screenshots/", custom_range=args.range)
        for name, path in results.items():
            print(f"  {name}: {path}")

    if getattr(args, "data", False) or not getattr(args, "screenshot", False):
        if not sheets:
            print("Specify --sheets for data dump")
            sys.exit(1)
        from xlvbatools.workbook.dumper import dump_sheet_data
        out_json = "dump.json" if getattr(args, "json", False) else None
        out_md = "dump.md" if not getattr(args, "json", False) else None
        dump_sheet_data(wb, sheets, output_json=out_json, output_md=out_md,
                        custom_range=args.range)
        print(f"Dumped to: {out_json or out_md}")


def _cmd_modify(args):
    """Modify cell values or named ranges."""
    from xlvbatools.config.loader import load_config
    from xlvbatools.logging import setup_logging
    cfg = load_config()
    setup_logging(verbose=getattr(args, "verbose", False), tool_name="modify",
                  log_dir=cfg.log_dir, log_name=cfg.log_name)

    wb = args.workbook or cfg.workbook

    # Parse value with type detection
    value = args.value
    if value is not None:
        try:
            value = float(value)
            if value == int(value):
                value = int(value)
        except ValueError:
            pass

    from xlvbatools.workbook.modifier import modify_cell
    success = modify_cell(
        wb,
        sheet=args.sheet or "Sheet1",
        cell=args.cell,
        value=value,
        formula=args.formula,
        name=args.name,
        refers_to=args.refers_to,
        delete_name=args.delete_name,
    )
    print("OK" if success else "FAILED")
    sys.exit(0 if success else 1)


def _cmd_debug(args):
    """Launch interactive debug session."""
    from xlvbatools.config.loader import load_config
    cfg = load_config()
    wb = args.workbook or cfg.workbook

    from xlvbatools.workbook.debugger import launch_debug_session
    launch_debug_session(wb, open_vbe=not args.no_vbe)


def _cmd_search(args):
    """Search VBA source files."""
    from xlvbatools.config.loader import load_config
    from xlvbatools.logging import setup_logging
    cfg = load_config()
    setup_logging(tool_name="search", log_dir=cfg.log_dir, log_name=cfg.log_name)

    src = args.source or cfg.vba_source
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
    src = args.source or cfg.vba_source

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
    src = args.source or cfg.vba_source

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


def _cmd_agents(args):
    """Print AI agent integration instructions and best practices."""
    help_text = """================================================================================
                    xlvbatools AI Agent Integration Guide
================================================================================

This toolkit supports the July 2026 open-standard for agent customizations
(agentskills.io and Google Antigravity specifications). By placing instructions
in a central `.agents/` directory, you allow coding assistants to discover rules,
capabilities, and workflows without polluting your workspace or bloating context.

--------------------------------------------------------------------------------
1. Standard Directory Layout (.agents/)
--------------------------------------------------------------------------------
When initialized, the following structure will be installed in the project root:

    .agents/
    ├── AGENTS.md           # Mission Control & main developer rules
    ├── rules/              # Task-specific conditional guidelines
    │   ├── vba-rules.md    # VBA coding standards & encoding constraints
    │   └── python-rules.md # Python packaging & test guidelines
    ├── skills/             # Portable capability packs (agentskills.io)
    │   └── xlvba-toolchain/
    │       └── SKILL.md    # xlvba CLI reference & python API usage
    └── workflows/          # Repeatable step-by-step pipeline procedures
        ├── vba-edit.md     # The modify-lint-inject-verify pipeline
        └── vba-debug.md    # Dialog watchdogs & COM error debugging

--------------------------------------------------------------------------------
2. How to Bootstrap/Install
--------------------------------------------------------------------------------
To install or refresh these agent templates in your current workspace, run:

    xlvba init --agents

If xlvbatools.toml already exists, you can run:

    xlvba init --agents --force

--------------------------------------------------------------------------------
3. Information for AI Agents (Antigravity, Cursor, Copilot, etc.)
--------------------------------------------------------------------------------
If you are an AI assistant:
- READ .agents/AGENTS.md first to understand the workspace overview and rules.
- REFER to skills/xlvba-toolchain/SKILL.md to understand the Python API and CLI.
- USE .agents/workflows/vba-edit.md as your standard pipeline for modifying VBA files.
- USE .agents/workflows/vba-debug.md when encountering Excel COM hangs, modal dialogs, or VBE programmatic errors.
================================================================================
"""
    print(help_text)


if __name__ == "__main__":
    main()
