"""Command implementations for the xlvba CLI."""

import json
import os
import sys

from xlvbatools import Artifact
from xlvbatools.cli.presentation import (
    fail_result as _fail_result,
    is_machine as _is_machine,
    is_table as _is_table,
    local_result as _local_result,
    print_result_json as _print_result_json,
    print_table as _print_table,
)


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


# ── Command handlers ──

def _cmd_init(args):
    """Initialize xlvbatools.toml in the current directory."""
    from xlvbatools.cli.init_cmd import run_init
    try:
        output = run_init(args)
    except (FileExistsError, OSError, RuntimeError) as error:
        result = _local_result("init", success=False, error=error, code="init_failed")
        _fail_result(args, result, "Initialization failed")
    result = _local_result("init", output)
    if _is_machine(args):
        _print_result_json(result)
    elif _is_table(args):
        _print_table(
            ("field", "value"),
            (
                ("config", output.config_path),
                ("workbook", output.workbook),
                ("agents", output.agents_status),
            ),
        )
    else:
        print(f"Created: {output.config_path}")
        for directory in output.directories:
            print(f"Created: {directory}")
        print("Project initialized.")


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
            _fail_result(args, result, "Component listing failed")
        components = result.data or []
        if _is_machine(args):
            _print_result_json(result)
        elif _is_table(args):
            _print_table(
                ("type", "name", "lines"),
                ((item.type_name, item.name, item.line_count) for item in components),
            )
        else:
            for c in components:
                print(
                    f"  {c.type_name:20s} {c.name:30s} "
                    f"({c.line_count} lines)"
                )
        return

    result = project.extract(
        output=out, component=args.component, timeout=args.timeout,
    )
    if not result.success:
        _fail_result(args, result, "Extraction failed")

    data = result.data
    if _is_machine(args):
        _print_result_json(result)
        return
    if args.component:
        extracted = data.components[0]
        if _is_table(args):
            _print_table(
                ("name", "type", "file", "lines"),
                ((extracted.name, extracted.type_name, extracted.file, extracted.line_count),),
            )
        else:
            print(f"Extracted: {extracted.name} -> {extracted.file}")
    else:
        count = len(data.components)
        if _is_table(args):
            _print_table(
                ("name", "type", "file", "lines"),
                (
                    (item.name, item.type_name, item.file, item.line_count)
                    for item in data.components
                ),
            )
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
        result = _local_result(
            "inject",
            success=False,
            error="--dry-run cannot be combined with --component",
            code="invalid_arguments",
        )
        _fail_result(args, result, "Invalid injection arguments", exit_code=2)
    result = project.inject(
        source=src,
        component=args.component,
        backup=not args.no_backup,
        dry_run=args.dry_run,
        timeout=args.timeout,
    )
    if not result.success:
        _fail_result(args, result, "Injection failed")
    if _is_machine(args):
        _print_result_json(result)
        return
    if args.component:
        if _is_table(args):
            change = result.data.changes[0]
            _print_table(
                ("status", "name", "file", "reason"),
                ((change.status, change.name, change.file, change.reason or change.error),),
            )
        else:
            print("Injected")
    else:
        changes = result.data.changes
        if _is_table(args):
            _print_table(
                ("status", "name", "file", "reason"),
                (
                    (change.status, change.name, change.file, change.reason or change.error)
                    for change in changes
                ),
            )
        else:
            for change in changes:
                status = change.status.upper()
                print(f"  {status:10s} {change.name}")


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
        _fail_result(args, operation_result, "Diff failed")
    if _is_machine(args):
        _print_result_json(operation_result)
        return
    results = operation_result.data or []
    if _is_table(args):
        _print_table(
            ("status", "name", "added", "removed"),
            (
                (item.status, item.name, item.lines_added, item.lines_removed)
                for item in results
            ),
        )
        return
    if args.component:
        result = operation_result.data[0]
        _print_diff_result(result, args)
    else:
        for result in results:
            _print_diff_result(result, args)


def _print_diff_result(r, args):
    status = r.status
    if status == "identical":
        print(f"  = {r.name}")
    elif status == "modified":
        print(f"  M {r.name} (+{r.lines_added}/-{r.lines_removed})")
        if not getattr(args, "summary", False) and r.unified_diff is not None:
            print(r.unified_diff)
    else:
        print(f"  ? {r.name} ({status})")


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
        result = _local_result(
            "lint_source",
            success=False,
            error=error,
            code="invalid_lint_target",
        )
        _fail_result(args, result, "Lint target is invalid", exit_code=2)

    issues = result.data or ()

    if _is_machine(args):
        _print_result_json(result)
    elif _is_table(args):
        _print_table(
            ("severity", "rule", "file", "line", "message"),
            (
                (
                    issue.severity,
                    issue.rule_id,
                    issue.module,
                    issue.line_num,
                    issue.message,
                )
                for issue in issues
            ),
        )
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

    if _is_machine(args):
        _print_result_json(result)
    elif _is_table(args):
        payload = result.to_dict()
        macro = payload.get("data") or {}
        _print_table(
            ("field", "value"),
            (
                ("success", result.success),
                ("phase", result.phase),
                ("elapsed_seconds", result.elapsed_seconds),
                ("macro", macro.get("macro") if isinstance(macro, dict) else args.macro),
                ("run_id", macro.get("run_id") if isinstance(macro, dict) else None),
                ("excel_pid", result.diagnostics.excel_pid),
            ),
        )
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

    from xlvbatools import SnapshotError, SnapshotNotFoundError, SnapshotService
    snapshots = SnapshotService(
        wb, src, _cfg_path(cfg, "snapshots_path", "snapshots_dir"),
        rolling_limit=cfg.snapshots.rolling_limit,
    )

    sub = args.snapshot_command
    try:
        if sub == "create":
            data = snapshots.create(
                description=args.description,
                milestone=args.milestone,
            )
        elif sub == "list":
            data = snapshots.list()
        elif sub == "info":
            data = snapshots.info(args.identifier)
            if data is None:
                raise SnapshotNotFoundError(f"Snapshot not found: {args.identifier}")
        elif sub == "restore":
            data = snapshots.restore(
                args.identifier,
                safety_snapshot=not args.no_safety,
            )
        elif sub == "diff":
            data = snapshots.diff(args.identifier)
        elif sub == "prune":
            data = {"pruned": snapshots.prune(keep=args.keep), "keep": args.keep}
        else:
            raise ValueError("snapshot subcommand is required")
    except (SnapshotError, ValueError) as error:
        result = _local_result(
            f"snapshot_{sub or 'unknown'}",
            success=False,
            error=error,
            code=(
                "snapshot_not_found"
                if isinstance(error, SnapshotNotFoundError)
                else "snapshot_failed"
            ),
        )
        _fail_result(args, result, "Snapshot operation failed")

    result = _local_result(f"snapshot_{sub}", data)
    if _is_machine(args):
        _print_result_json(result)
    elif _is_table(args):
        if sub == "list":
            _print_table(
                ("id", "timestamp", "description", "milestone"),
                (
                    (
                        record.snapshot_id,
                        record.timestamp,
                        record.description,
                        record.milestone,
                    )
                    for record in data
                ),
            )
        elif hasattr(data, "to_dict"):
            _print_table(("field", "value"), data.to_dict().items())
        elif isinstance(data, dict):
            _print_table(("field", "value"), data.items())
        else:
            _print_table(("field", "value"), (("result", data),))
    elif sub == "create":
        print(f"Snapshot created: {data.snapshot_id}")
    elif sub == "list":
        for record in data:
            milestone = " [MILESTONE]" if record.milestone else ""
            print(f"  {record.snapshot_id}  {record.description}{milestone}")
    elif sub == "info":
        print(json.dumps(data.to_dict(), indent=2))
    elif sub == "restore":
        print(f"Restored: {data.snapshot_id}")
    elif sub == "diff":
        print(data)
    elif sub == "prune":
        print(f"Pruned {data['pruned']} snapshots")


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
        result = _local_result(
            "inspect",
            success=False,
            error="Specify --sheets for workbook inspection",
            code="invalid_arguments",
        )
        _fail_result(args, result, "Missing worksheet selection")
    project = _project(cfg, workbook=args.workbook)
    include_screenshots = getattr(args, "screenshot", False)
    include_data = getattr(args, "data", False) or not include_screenshots
    out_json = args.write_json if include_data else None
    out_md = args.write_markdown if include_data else None
    result = project.inspect(
        sheets, output_dir="screenshots", cell_range=args.range,
        include_data=include_data, include_screenshots=include_screenshots,
        output_json=out_json, output_markdown=out_md, timeout=args.timeout,
        include_hidden_sheets=args.include_hidden_sheets,
    )
    if not result.success:
        _fail_result(args, result, "Workbook inspection failed")
    if _is_machine(args):
        _print_result_json(result)
    elif _is_table(args):
        _print_table(
            ("sheet", "screenshot"),
            result.data.screenshots.items(),
        )
    else:
        for name, path in result.data.screenshots.items():
            print(f"  {name}: {path}")
        if out_json:
            print(f"JSON written to: {out_json}")
        if out_md:
            print(f"Markdown written to: {out_md}")


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
    if _is_machine(args):
        _print_result_json(result)
    elif _is_table(args):
        payload = result.to_dict().get("data")
        if isinstance(payload, dict):
            _print_table(("field", "value"), payload.items())
        else:
            _print_table(
                ("field", "value"),
                (("success", result.success), ("result", payload)),
            )
    else:
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
    if not os.path.isdir(src):
        result = _local_result(
            "search",
            success=False,
            error=f"Source directory not found: {os.path.abspath(src)}",
            code="source_not_found",
        )
        _fail_result(args, result, "Search source not found", exit_code=2)
    from xlvbatools.vba.search import search_vba
    matches = search_vba(src, args.pattern, regex=args.regex, context_lines=ctx)

    result = _local_result(
        "search",
        tuple(matches),
        metadata={
            "pattern": args.pattern,
            "regex": args.regex,
            "match_count": len(matches),
            "source": os.path.abspath(src),
        },
    )
    if _is_machine(args):
        _print_result_json(result)
        return
    if _is_table(args):
        _print_table(
            ("file", "line", "module", "text"),
            ((item.file, item.line_num, item.module, item.line) for item in matches),
        )
        return
    if not matches:
        print(f"No matches for: {args.pattern}")
        return

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
        item = format_file(src, dry_run=args.dry_run, indent_size=args.indent)
        item["file"] = os.path.abspath(src)
        results = (item,)
    elif os.path.isdir(src):
        from xlvbatools.vba.formatter import format_directory
        results = tuple(
            format_directory(src, dry_run=args.dry_run, indent_size=args.indent)
        )
    else:
        result = _local_result(
            "format",
            success=False,
            error=f"Source not found: {os.path.abspath(src)}",
            code="source_not_found",
        )
        _fail_result(args, result, "Format source not found", exit_code=2)

    changed = tuple(item for item in results if item.get("changed"))
    result = _local_result(
        "format",
        results,
        metadata={
            "source": os.path.abspath(src),
            "dry_run": args.dry_run,
            "changed_count": len(changed),
            "file_count": len(results),
        },
    )
    if _is_machine(args):
        _print_result_json(result)
        return
    if _is_table(args):
        _print_table(
            ("file", "changed", "lines_changed"),
            (
                (item.get("file", ""), item.get("changed", False), item.get("lines_changed", 0))
                for item in results
            ),
        )
        return
    if len(results) == 1 and os.path.isfile(src):
        item = results[0]
        if item.get("changed"):
            if args.dry_run:
                print(item.get("diff", ""))
            else:
                print(f"Formatted: {src} ({item['lines_changed']} lines changed)")
        else:
            print(f"No changes: {src}")
    else:
        for item in changed:
            if args.dry_run:
                print(item.get("diff", ""))
            else:
                print(f"  Formatted: {item['file']} ({item['lines_changed']} lines)")
        print(f"\n{len(changed)}/{len(results)} file(s) {'would be ' if args.dry_run else ''}changed")


def _cmd_graph(args):
    """Generate VBA call dependency graph."""
    from xlvbatools.config.loader import load_config
    cfg = load_config()
    src = args.source or _cfg_path(cfg, "vba_source_path", "vba_source")
    if not os.path.isdir(src):
        result = _local_result(
            "graph",
            success=False,
            error=f"Source directory not found: {os.path.abspath(src)}",
            code="source_not_found",
        )
        _fail_result(args, result, "Graph source not found", exit_code=2)

    from xlvbatools.vba.dependency import build_call_graph, render_mermaid, render_dot
    graph = build_call_graph(src)

    graph_format = args.graph_format
    graph_data = graph.to_dict()
    if graph_format == "mermaid":
        content = render_mermaid(graph)
        data = {"graph_format": graph_format, "content": content, "graph": graph_data}
    elif graph_format == "dot":
        content = render_dot(graph)
        data = {"graph_format": graph_format, "content": content, "graph": graph_data}
    else:
        content = json.dumps(graph_data, indent=2)
        data = {"graph_format": graph_format, "graph": graph_data}

    artifacts = ()
    if args.output:
        with open(args.output, "w", encoding="utf-8") as f:
            f.write(content)
        artifacts = (
            Artifact(
                kind="graph",
                path=os.path.abspath(args.output),
                media_type=(
                    "application/json"
                    if graph_format == "json"
                    else "text/vnd.graphviz"
                    if graph_format == "dot"
                    else "text/plain"
                ),
            ),
        )

    result = _local_result(
        "graph",
        data,
        metadata={
            "source": os.path.abspath(src),
            "node_count": len(graph.procedures),
            "edge_count": len(graph.edges),
        },
        artifacts=artifacts,
    )
    if _is_machine(args):
        _print_result_json(result)
    elif _is_table(args):
        _print_table(
            ("module", "procedure", "kind", "access", "line"),
            (
                (item.module, item.name, item.kind, item.access, item.line_num)
                for item in graph.procedures.values()
            ),
        )
    else:
        if args.output:
            print(f"Graph written to: {os.path.abspath(args.output)}")
        else:
            print(content)
        print(f"\n{len(graph.procedures)} procedures, {len(graph.edges)} call edges")


def _cmd_snapshot_help(args):
    if not hasattr(args, "snapshot_command") or args.snapshot_command is None:
        print("Usage: xlvba snapshot <create|list|info|restore|diff|prune>")
        sys.exit(0)


def _cmd_version(args):
    """Show exact installed package and source provenance."""
    from xlvbatools.version import get_version_info

    info = get_version_info()
    result = _local_result("version", info)
    if _is_machine(args):
        _print_result_json(result)
    elif _is_table(args):
        _print_table(("field", "value"), info.to_dict().items())
    else:
        print(f"xlvba {info.version}")
        print(f"Python: {info.python_executable}")
        print(f"Package: {info.package_path}")
        print(f"Source: {info.source_url or 'unknown'}")
        print(f"Commit: {info.commit_id or 'unknown'}")
        if info.requested_revision:
            print(f"Requested revision: {info.requested_revision}")


def _cmd_help(args):
    """Return the stable command-discovery catalog."""
    from xlvbatools.cli.discovery import discovery_payload

    payload = discovery_payload(getattr(args, "target", None))
    result = _local_result("help", payload)
    if _is_machine(args):
        _print_result_json(result)
    elif _is_table(args):
        command = payload.get("command")
        commands = [command] if command is not None else payload["commands"]
        _print_table(
            ("command", "purpose", "excel", "interactive"),
            (
                (
                    item["name"],
                    item["summary"],
                    "yes" if item["excel_required"] else "no",
                    "yes" if item["interactive"] else "no",
                )
                for item in commands
            ),
        )
    else:
        command = payload.get("command")
        if command is not None:
            print(f"{command['name']}: {command['summary']}")
            print(f"Usage: {command['usage']}")
            print("Examples:")
            for example in command["examples"]:
                print(f"  {example}")
            return
        print("xlvbatools command discovery")
        for item in payload["commands"]:
            print(f"  {item['name']:10s} {item['summary']}")
        print("\nUse 'xlvba help COMMAND' for examples and 'COMMAND --help' for options.")


def _cmd_agents(args):
    """Print AI agent integration instructions and best practices."""
    if getattr(args, "agents_command", None) == "install":
        from xlvbatools.cli.init_cmd import install_agents_template

        try:
            output = install_agents_template(args.destination, force=args.force)
        except (OSError, RuntimeError) as error:
            result = _local_result(
                "agents_install", success=False, error=error, code="agent_install_failed"
            )
            _fail_result(args, result, "Agent template installation failed")
        result = _local_result("agents_install", output)
        if _is_machine(args):
            _print_result_json(result)
        elif _is_table(args):
            _print_table(
                ("action", "file"),
                (
                    *(("installed", path) for path in output.installed),
                    *(("overwritten", path) for path in output.overwritten),
                    *(("skipped", path) for path in output.skipped),
                ),
            )
        else:
            print(f"Agent guidance {output.status}: {output.destination}")
            print(
                f"Installed {len(output.installed)}, overwritten "
                f"{len(output.overwritten)}, skipped {len(output.skipped)} files."
            )
            print("Read .agents/AGENTS.md first, then commit any project customizations.")
        return

    guidance = {
        "public_boundary": ["xlvbatools.Project", "xlvba CLI"],
        "private_boundaries": ["raw COM sessions", "worker transport files"],
        "read_order": [
            ".agents/AGENTS.md",
            ".agents/skills/xlvba-toolchain/SKILL.md",
            ".agents/rules/python-rules.md or vba-rules.md",
            ".agents/workflows/vba-edit.md or vba-debug.md",
        ],
        "safety_requirements": [
            "never terminate Excel globally",
            "require operation success and clean owned-process shutdown",
            "render hidden worksheets only when explicitly requested",
            "use the repository .venv and pin exact versions or revisions",
        ],
        "default_output": "json",
        "presentation_flags": [
            "--output-format text",
            "--output-format table",
            "--text",
            "--table",
        ],
        "template_installation": {
            "existing_project": "xlvba agents install",
            "new_project": "xlvba init --agents",
            "destination": ".agents/",
            "force_update": "xlvba agents install --force",
        },
        "discovery": ["xlvba --help", "xlvba help", "xlvba help COMMAND"],
    }
    result = _local_result("agents", guidance)
    if _is_machine(args):
        _print_result_json(result)
        return
    help_text = """xlvbatools v1 agent integration

Application automation uses xlvbatools.Project or the xlvba CLI. Raw COM
sessions and worker files are private implementation details.

Install guidance into an existing project:
    xlvba agents install

Or initialize a new project and install it in one step:
    xlvba init --workbook workbook/MyProject.xlsm --agents

The destination is .agents/ (plural). Installation copies missing packaged
files and preserves existing files. Use --force only to overwrite packaged
file paths; project-specific extra files are never deleted.

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
    if _is_table(args):
        _print_table(
            ("category", "value"),
            (
                (category, value)
                for category, values in guidance.items()
                for value in (values if isinstance(values, list) else [values])
            ),
        )
    else:
        print(help_text)
