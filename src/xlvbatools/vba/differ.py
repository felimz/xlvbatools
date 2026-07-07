"""
VBA Differ
===========
Compares VBA code between a live workbook and the on-disk vba_source/ files.
Produces unified diffs with change counting.

Usage:
    from xlvbatools.vba import diff_all, diff_component

    diffs = diff_all("workbook.xlsm", "vba_source/")
    result = diff_component("workbook.xlsm", "vba_source/", "modMain")
"""

import difflib
import logging
import os
import re
import tempfile

from xlvbatools.core.session import ExcelSession
from xlvbatools.vba.manifest import get_type_info

logger = logging.getLogger(__name__)

# VBE header lines that should be stripped for comparison
_VBE_HEADER_RE = re.compile(
    r"^(Attribute VB_|VERSION \d|BEGIN|END|  MultiUse =)", re.IGNORECASE
)


def diff_all(workbook_path: str, source_dir: str) -> list[dict]:
    """
    Compare all VBA components between the workbook and source files.

    Returns a list of dicts, one per component:
        name, status ("identical"|"modified"|"missing_source"|"missing_workbook"),
        lines_added, lines_removed, unified_diff
    """
    wb_path = os.path.abspath(workbook_path)
    src_dir = os.path.abspath(source_dir)
    results = []

    with ExcelSession(wb_path, visible=False, save_on_exit=False) as session:
        vb_project = session.wb.VBProject

        # Build lookup of source files
        source_files = _scan_source_files(src_dir)

        # Compare each workbook component against source
        wb_names = set()
        for comp in vb_project.VBComponents:
            name = comp.Name
            wb_names.add(name.lower())

            wb_lines = _get_component_code(comp)
            src_path = source_files.get(name.lower())

            if src_path is None:
                results.append({
                    "name": name,
                    "status": "missing_source",
                    "lines_added": 0,
                    "lines_removed": 0,
                })
                continue

            src_lines = _read_source_file(src_path, comp.Type)
            result = _compare(name, wb_lines, src_lines, src_path)
            results.append(result)

        # Check for source files not in workbook
        for src_name, src_path in source_files.items():
            if src_name not in wb_names:
                results.append({
                    "name": src_name,
                    "status": "missing_workbook",
                    "lines_added": 0,
                    "lines_removed": 0,
                })

    return sorted(results, key=lambda r: r["name"].lower())


def diff_component(
    workbook_path: str,
    source_dir: str,
    component_name: str,
) -> dict | None:
    """
    Compare a single VBA component between workbook and source.

    Returns a result dict, or None if the component is not found in either location.
    """
    wb_path = os.path.abspath(workbook_path)
    src_dir = os.path.abspath(source_dir)

    with ExcelSession(wb_path, visible=False, save_on_exit=False) as session:
        vb_project = session.wb.VBProject

        # Find in workbook
        comp = None
        for c in vb_project.VBComponents:
            if c.Name.lower() == component_name.lower():
                comp = c
                break

        if comp is None:
            return None

        wb_lines = _get_component_code(comp)

        # Find source file
        source_files = _scan_source_files(src_dir)
        src_path = source_files.get(component_name.lower())

        if src_path is None:
            return {
                "name": component_name,
                "status": "missing_source",
                "lines_added": 0,
                "lines_removed": 0,
            }

        src_lines = _read_source_file(src_path, comp.Type)
        return _compare(component_name, wb_lines, src_lines, src_path)


def _compare(name: str, wb_lines: list[str], src_lines: list[str], src_path: str) -> dict:
    """Generate a unified diff between workbook and source lines."""
    if wb_lines == src_lines:
        return {
            "name": name,
            "status": "identical",
            "lines_added": 0,
            "lines_removed": 0,
        }

    diff = list(difflib.unified_diff(
        src_lines, wb_lines,
        fromfile=f"source/{os.path.basename(src_path)}",
        tofile=f"workbook/{name}",
        lineterm="",
    ))

    added = sum(1 for line in diff if line.startswith("+") and not line.startswith("+++"))
    removed = sum(1 for line in diff if line.startswith("-") and not line.startswith("---"))

    return {
        "name": name,
        "status": "modified",
        "lines_added": added,
        "lines_removed": removed,
        "unified_diff": "\n".join(diff),
    }


def _get_component_code(comp) -> list[str]:
    """Read code from a VBE component, stripping VBE header lines."""
    cm = comp.CodeModule
    total = cm.CountOfLines
    if total == 0:
        return []

    code = cm.Lines(1, total)
    lines = code.split("\r\n") if "\r\n" in code else code.split("\n")

    # Strip VBE header lines for clean comparison
    stripped = [line for line in lines if not _VBE_HEADER_RE.match(line)]

    # Remove leading/trailing blank lines
    while stripped and not stripped[0].strip():
        stripped.pop(0)
    while stripped and not stripped[-1].strip():
        stripped.pop()

    return stripped


def _read_source_file(filepath: str, vbe_type: int) -> list[str]:
    """Read a source file, stripping VBE headers for clean comparison."""
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            lines = f.read().split("\n")
    except UnicodeDecodeError:
        with open(filepath, "r", encoding="windows-1252") as f:
            lines = f.read().split("\n")

    # Strip \r from line endings
    lines = [line.rstrip("\r") for line in lines]

    # Strip VBE headers
    stripped = [line for line in lines if not _VBE_HEADER_RE.match(line)]

    # Remove leading/trailing blank lines
    while stripped and not stripped[0].strip():
        stripped.pop(0)
    while stripped and not stripped[-1].strip():
        stripped.pop()

    return stripped


def _scan_source_files(source_dir: str) -> dict[str, str]:
    """Build a lowercase name -> filepath mapping from the source directory."""
    files = {}
    for subdir in ("modules", "classes", "sheets"):
        dir_path = os.path.join(source_dir, subdir)
        if not os.path.isdir(dir_path):
            continue
        for fname in os.listdir(dir_path):
            if fname.endswith((".bas", ".cls")):
                name = os.path.splitext(fname)[0]
                files[name.lower()] = os.path.join(dir_path, fname)
    return files
