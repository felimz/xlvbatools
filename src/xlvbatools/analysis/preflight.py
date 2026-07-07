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

    for root, _, files in os.walk(src_dir):
        for fname in sorted(files):
            if not fname.endswith(extensions):
                continue
            filepath = os.path.join(root, fname)
            rel_path = os.path.relpath(filepath, src_dir)

            lines = _read_file_lines(filepath)
            issues = run_all_rules(rel_path, lines, disabled_rules)
            all_issues.extend(issues)

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
        for comp in session.wb.VBProject.VBComponents:
            name = comp.Name
            type_info = get_type_info(comp.Type)

            cm = comp.CodeModule
            if cm.CountOfLines == 0:
                continue

            code = cm.Lines(1, cm.CountOfLines)
            lines = code.split("\r\n") if "\r\n" in code else code.split("\n")
            rel_path = f"{type_info['dir']}/{name}{type_info['ext']}"

            issues = run_all_rules(rel_path, lines, disabled_rules)
            all_issues.extend(issues)

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
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            return f.readlines()
    except UnicodeDecodeError:
        with open(filepath, "r", encoding="windows-1252") as f:
            return f.readlines()
