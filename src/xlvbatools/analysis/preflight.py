"""
VBA Preflight Static Analysis
===============================
Top-level API for running static analysis on VBA source files or workbooks.

Usage:
    from xlvbatools.analysis import lint_files, lint_workbook

    issues = lint_files("vba_source/")
    issues = lint_workbook("workbook.xlsm")
"""

import logging
import os
import re
from typing import List

from xlvbatools.analysis.issue import VBAIssue
from xlvbatools.analysis.rules import run_all_rules

logger = logging.getLogger(__name__)


def lint_files(
    source_dir: str,
    disabled_rules: List[str] | None = None,
    extensions: tuple = (".bas", ".cls"),
) -> List[VBAIssue]:
    """
    Run static analysis on extracted VBA source files (no COM needed).

    Parameters
    ----------
    source_dir : str
        Path to the vba_source/ directory.
    disabled_rules : list of str, optional
        Rule IDs to skip (e.g. ["PF001", "PF003"]).
    extensions : tuple
        File extensions to check.

    Returns
    -------
    list of VBAIssue
        All issues found, sorted by file then line number.
    """
    src_dir = os.path.abspath(source_dir)
    if not os.path.isdir(src_dir):
        logger.error(f"Source directory not found: {src_dir}")
        return []

    all_issues = []

    # First pass: collect all Public module-level variable names across all files
    # so UV001 can recognize cross-module scope.
    global_public_names = set()
    _public_re = re.compile(r"^(Public|Global)\s+", re.IGNORECASE)
    _proc_start_re = re.compile(
        r"^(Public\s+|Private\s+|Friend\s+)?(Sub|Function|Property\s+\w+)\s+",
        re.IGNORECASE,
    )
    global_public_names = set()
    seen_public_procs = {}
    class_registry = {}

    for root, _, files in os.walk(src_dir):
        for fname in sorted(files):
            if not fname.endswith(extensions):
                continue
            filepath = os.path.join(root, fname)
            rel_path = os.path.relpath(filepath, src_dir)
            file_lines = _read_file_lines(filepath)
            
            # Extract public names from header
            for fline in file_lines:
                stripped = fline.strip()
                if _proc_start_re.match(stripped):
                    break  # Stop at first procedure (module-level scope ends)
                if _public_re.match(stripped):
                    # Skip Public Sub/Function/Property/Const declarations
                    if re.match(r"^(Public\s+)?(Sub|Function|Property|Const)\s+", stripped, re.IGNORECASE):
                        # But extract Const names
                        const_m = re.match(r"^Public\s+Const\s+(\w+)", stripped, re.IGNORECASE)
                        if const_m:
                            global_public_names.add(const_m.group(1).lower())
                        continue
                    # Extract variable names from Public declarations
                    from xlvbatools.analysis.rules import _parse_vba_declarations
                    for var_name, _ in _parse_vba_declarations(stripped):
                        global_public_names.add(var_name.lower())

            # Extract class members if this is a class module
            if fname.endswith(".cls"):
                cls_name = os.path.splitext(fname)[0].lower()
                cls_members = set()
                for fline in file_lines:
                    stripped = fline.strip()
                    if stripped.startswith("'") or stripped.lower().startswith("rem "):
                        continue
                    m = re.match(
                        r"^(Public\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)",
                        stripped,
                        re.IGNORECASE,
                    )
                    if m:
                        cls_members.add(m.group(2).lower())
                    m_var = re.match(r"^Public\s+(\w+)", stripped, re.IGNORECASE)
                    if m_var:
                        cls_members.add(m_var.group(1).lower())
                class_registry[cls_name] = cls_members

            # Check duplicate public procedures
            disabled = set(disabled_rules or [])
            if "DP001" not in disabled:
                from xlvbatools.analysis.rules import _PROC_START_RE, _extract_proc_name
                for i, fline in enumerate(file_lines, 1):
                    stripped = fline.strip()
                    if _PROC_START_RE.match(stripped):
                        is_private = False
                        for prefix in ("private ", "friend "):
                            if stripped.lower().startswith(prefix):
                                is_private = True
                                break
                        if not is_private:
                            proc_name = _extract_proc_name(stripped)
                            proc_key = proc_name.lower()
                            is_sheet = "sheets" in rel_path.lower() or fname.lower().startswith("sheet")
                            if not is_sheet and proc_key:
                                if proc_key in seen_public_procs:
                                    prev_path, prev_line = seen_public_procs[proc_key]
                                    all_issues.append(VBAIssue(
                                        rule_id="DP001",
                                        severity="ERROR",
                                        module=rel_path,
                                        line_num=i,
                                        message=f"Duplicate public procedure '{proc_name}' found in both '{rel_path}' and '{prev_path}' (L{prev_line}). VBA raises a compile error (Ambiguous name detected) for duplicate public names.",
                                    ))
                                else:
                                    seen_public_procs[proc_key] = (rel_path, i)

    # Second pass: run all rules with cross-module context
    for root, _, files in os.walk(src_dir):
        for fname in sorted(files):
            if not fname.endswith(extensions):
                continue
            filepath = os.path.join(root, fname)
            rel_path = os.path.relpath(filepath, src_dir)

            lines = _read_file_lines(filepath)
            issues = run_all_rules(
                rel_path, lines, disabled_rules,
                global_names=global_public_names,
                class_registry=class_registry
            )
            all_issues.extend(issues)

    disabled = set(disabled_rules or [])
    if "DC003" not in disabled:
        from xlvbatools.vba.dependency import build_call_graph
        from xlvbatools.analysis.rules import _is_entry_point
        try:
            cg = build_call_graph(src_dir)
            called_qnames = {edge[1].lower() for edge in cg.edges}
            module_dirs = {}
            for root, _, files in os.walk(src_dir):
                for fname in files:
                    if fname.endswith(extensions):
                        mod_name = os.path.splitext(fname)[0]
                        module_dirs[mod_name] = os.path.basename(root)
            for qname, proc in cg.procedures.items():
                if qname.lower() in called_qnames:
                    continue
                # Skip known entry-point patterns (event handlers, button clicks, etc.)
                if _is_entry_point(proc.name):
                    continue
                parent_dir = module_dirs.get(proc.module, "")
                is_internal = (
                    proc.access.lower() == "private" or 
                    parent_dir in ("classes", "sheets")
                )
                if is_internal:
                    rel_file_path = ""
                    for root_dir, _, sub_files in os.walk(src_dir):
                        for sub_f in sub_files:
                            if os.path.splitext(sub_f)[0] == proc.module:
                                rel_file_path = os.path.relpath(os.path.join(root_dir, sub_f), src_dir)
                                break
                    all_issues.append(VBAIssue(
                        rule_id="DC003",
                        severity="WARNING",
                        module=rel_file_path or proc.module,
                        line_num=proc.line_num,
                        message=f"Dead procedure '{proc.name}' in module '{proc.module}'. This procedure has zero incoming calls and is not a known event handler. ACTION: Delete or implement calls to this procedure to clean up dead code.",
                        procedure=proc.name
                    ))
        except Exception as e:
            logger.warning(f"Failed to run DC003 dead code analysis: {e}")

    all_issues.sort(key=lambda i: (i.module, i.line_num))
    return all_issues


def lint_workbook(
    workbook_path: str,
    disabled_rules: List[str] | None = None,
    compile_test: bool = True,
) -> List[VBAIssue]:
    """
    Run static analysis on a live workbook via COM.

    Extracts VBA code directly from the workbook's VBProject and runs
    all rules. Optionally triggers VBE compile test for syntax errors.

    Parameters
    ----------
    workbook_path : str
        Path to the .xlsm workbook.
    disabled_rules : list of str, optional
        Rule IDs to skip.
    compile_test : bool
        Whether to also run VBE compile test (default True).

    Returns
    -------
    list of VBAIssue
    """
    from xlvbatools.core.session import ExcelSession
    from xlvbatools.vba.manifest import get_type_info

    wb_path = os.path.abspath(workbook_path)
    all_issues = []

    with ExcelSession(wb_path, visible=False, save_on_exit=False) as session:
        # Run rules against each component's code
        component_codes = {}
        for comp in session.wb.VBProject.VBComponents:
            name = comp.Name
            type_info = get_type_info(comp.Type)

            cm = comp.CodeModule
            if cm.CountOfLines == 0:
                continue

            code = cm.Lines(1, cm.CountOfLines)
            component_codes[name] = (code, type_info)
            lines = code.split("\r\n") if "\r\n" in code else code.split("\n")
            rel_path = f"{type_info['dir']}/{name}{type_info['ext']}"

            issues = run_all_rules(rel_path, lines, disabled_rules)
            all_issues.extend(issues)
        if "comp" in locals():
            del comp
        if "cm" in locals():
            del cm

        # Optional compile test
        if compile_test:
            result = session.compile_test()
            if not result["success"]:
                for err in result["errors"]:
                    all_issues.append(VBAIssue(
                        rule_id="CT001",
                        severity="ERROR",
                        module=result.get("error_module", "unknown"),
                        line_num=result.get("error_line", 0),
                        message=f"Compile error: {err.get('message', err.get('text', str(err)))}",
                    ))

        # Optional dead code analysis (DC003)
        disabled = set(disabled_rules or [])
        if "DC003" not in disabled:
            import tempfile
            import shutil
            from xlvbatools.vba.dependency import build_call_graph
            from xlvbatools.analysis.rules import _is_entry_point
            temp_dir = tempfile.mkdtemp()
            try:
                # Write component codes directly to temp directory without calling extract_all
                for name, (code, type_info) in component_codes.items():
                    subdir = os.path.join(temp_dir, type_info["dir"])
                    os.makedirs(subdir, exist_ok=True)
                    filepath = os.path.join(subdir, f"{name}{type_info['ext']}")
                    with open(filepath, "w", encoding="utf-8") as f:
                        f.write(code)

                cg = build_call_graph(temp_dir)
                called_qnames = {edge[1].lower() for edge in cg.edges}
                module_dirs = {}
                for root, _, files in os.walk(temp_dir):
                    for fname in files:
                        if fname.endswith((".bas", ".cls")):
                            mod_name = os.path.splitext(fname)[0]
                            module_dirs[mod_name] = os.path.basename(root)
                for qname, proc in cg.procedures.items():
                    if qname.lower() in called_qnames:
                        continue
                    # Skip known entry-point patterns (event handlers, button clicks, etc.)
                    if _is_entry_point(proc.name):
                        continue
                    parent_dir = module_dirs.get(proc.module, "")
                    is_internal = (
                        proc.access.lower() == "private" or 
                        parent_dir in ("classes", "sheets")
                    )
                    if is_internal:
                        ext = ".bas" if parent_dir == "modules" else ".cls"
                        rel_file_path = f"{parent_dir}/{proc.module}{ext}" if parent_dir else f"{proc.module}{ext}"
                        all_issues.append(VBAIssue(
                            rule_id="DC003",
                            severity="WARNING",
                            module=rel_file_path,
                            line_num=proc.line_num,
                            message=f"Dead procedure '{proc.name}' in module '{proc.module}'. This procedure has zero incoming calls and is not a known event handler. ACTION: Delete or implement calls to this procedure to clean up dead code.",
                            procedure=proc.name
                        ))
            except Exception as e:
                logger.warning(f"Failed to run DC003 dead code analysis on workbook: {e}")
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)

    all_issues.sort(key=lambda i: (i.module, i.line_num))
    return all_issues


def print_report(issues: List[VBAIssue]) -> str:
    """Format issues as a human-readable report string."""
    if not issues:
        return "PASS: No issues found"

    errors = [i for i in issues if i.severity == "ERROR"]
    warnings = [i for i in issues if i.severity == "WARNING"]

    lines = []
    for issue in issues:
        lines.append(str(issue))

    if errors:
        lines.append(f"\nFAIL: {len(errors)} errors, {len(warnings)} warnings")
    else:
        lines.append(f"\nPASS: {len(warnings)} warnings")

    return "\n".join(lines)


def _read_file_lines(filepath: str) -> list[str]:
    """Read a file's lines, handling encoding gracefully."""
    from xlvbatools.vba._io import read_vba_lines
    return read_vba_lines(filepath)
