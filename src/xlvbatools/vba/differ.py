"""
VBA Differ
===========
Compares VBA code between a live workbook and the on-disk vba_source/ files.
Produces unified diffs with change counting.

Usage:
    from xlvbatools import Project

    result = Project.from_config().diff()

    diffs = diff_all("workbook.xlsm", "vba_source/")
    result = diff_component("workbook.xlsm", "vba_source/", "modMain")
"""

import difflib
import logging
import os
from contextlib import nullcontext

from xlvbatools.core.session import ExcelSession
from xlvbatools.vba.constants import VBE_HEADER_STRIP_RE as _VBE_HEADER_RE

logger = logging.getLogger(__name__)


COMPARISON_MODES = frozenset({"vba", "text"})


def diff_all(
    workbook_path: str,
    source_dir: str,
    *,
    comparison: str = "vba",
    _session=None,
) -> list[dict]:
    """
    Compare all VBA components between the workbook and source files.

    Returns a list of dicts, one per component:
        name, status ("identical"|"equivalent"|"modified"|"missing_source"|
        "missing_workbook"), comparison, lines_added, lines_removed, unified_diff
    """
    comparison = _validate_comparison(comparison)
    wb_path = os.path.abspath(workbook_path)
    src_dir = os.path.abspath(source_dir)
    results = []

    session_context = (
        nullcontext(_session) if _session is not None
        else ExcelSession(wb_path, visible=False, save_on_exit=False)
    )
    with session_context as session:
        vb_project = session.vb_project

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
                    "comparison": comparison,
                    "lines_added": 0,
                    "lines_removed": 0,
                })
                continue

            src_lines = _read_source_file(src_path, comp.Type)
            result = _compare(
                name, wb_lines, src_lines, src_path, comparison=comparison,
            )
            results.append(result)

        # Check for source files not in workbook
        for src_name, src_path in source_files.items():
            if src_name not in wb_names:
                results.append({
                    "name": src_name,
                    "status": "missing_workbook",
                    "comparison": comparison,
                    "lines_added": 0,
                    "lines_removed": 0,
                })

    return sorted(results, key=lambda r: r["name"].lower())


def diff_component(
    workbook_path: str,
    source_dir: str,
    component_name: str,
    *,
    comparison: str = "vba",
    _session=None,
) -> dict | None:
    """
    Compare a single VBA component between workbook and source.

    Returns a result dict, or None if the component is not found in either location.
    """
    comparison = _validate_comparison(comparison)
    wb_path = os.path.abspath(workbook_path)
    src_dir = os.path.abspath(source_dir)

    session_context = (
        nullcontext(_session) if _session is not None
        else ExcelSession(wb_path, visible=False, save_on_exit=False)
    )
    with session_context as session:
        vb_project = session.vb_project

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
                "comparison": comparison,
                "lines_added": 0,
                "lines_removed": 0,
            }

        src_lines = _read_source_file(src_path, comp.Type)
        return _compare(
            component_name,
            wb_lines,
            src_lines,
            src_path,
            comparison=comparison,
        )


def _compare(
    name: str,
    wb_lines: list[str],
    src_lines: list[str],
    src_path: str,
    *,
    comparison: str = "vba",
) -> dict:
    """Generate a unified diff between workbook and source lines."""
    comparison = _validate_comparison(comparison)
    if wb_lines == src_lines:
        return {
            "name": name,
            "status": "identical",
            "comparison": comparison,
            "lines_added": 0,
            "lines_removed": 0,
        }

    if comparison == "vba" and _normalize_vba_lines(wb_lines) == _normalize_vba_lines(
        src_lines
    ):
        return {
            "name": name,
            "status": "equivalent",
            "comparison": comparison,
            "equivalence": "vba_token_equivalent",
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
        "comparison": comparison,
        "lines_added": added,
        "lines_removed": removed,
        "unified_diff": "\n".join(diff),
    }


def _validate_comparison(value: str) -> str:
    """Validate and normalize a public comparison mode."""
    normalized = str(value).strip().casefold()
    if normalized not in COMPARISON_MODES:
        choices = ", ".join(sorted(COMPARISON_MODES))
        raise ValueError(f"comparison must be one of: {choices}")
    return normalized


def _normalize_vba_lines(
    lines: list[str],
) -> list[tuple[tuple[str, str], ...]]:
    """Tokenize lines while preserving line structure, literals, and comments."""
    return [_normalize_vba_line(line) for line in lines]


def _normalize_vba_line(line: str) -> tuple[tuple[str, str], ...]:
    """Return VBA-aware tokens, excluding insignificant code whitespace."""
    normalized: list[tuple[str, str]] = []
    index = 0
    statement_start = True
    while index < len(line):
        char = line[index]

        if char == '"':
            literal_start = index
            index += 1
            while index < len(line):
                if line[index] == '"':
                    if index + 1 < len(line) and line[index + 1] == '"':
                        index += 2
                        continue
                    index += 1
                    break
                index += 1
            normalized.append(("string", line[literal_start:index]))
            statement_start = False
            continue

        if char == "'":
            normalized.append(("comment", line[index:]))
            break

        if char == "[":
            token_start = index
            index += 1
            while index < len(line) and line[index] != "]":
                index += 1
            if index < len(line):
                index += 1
            normalized.append(("bracket", line[token_start:index]))
            statement_start = False
            continue

        if char == "#" and "#" in line[index + 1:]:
            token_start = index
            index = line.index("#", index + 1) + 1
            normalized.append(("date", line[token_start:index]))
            statement_start = False
            continue

        if char.isspace():
            index += 1
            continue

        if char.isalpha() or char == "_":
            token_start = index
            index += 1
            while index < len(line) and (
                line[index].isalnum() or line[index] == "_"
            ):
                index += 1
            token = line[token_start:index]
            normalized.append(("identifier", token.casefold()))
            if statement_start and token.casefold() == "rem":
                normalized.append(("comment", line[index:]))
                break
            statement_start = False
            continue

        if char.isdigit() and statement_start:
            token_start = index
            while index < len(line) and line[index].isdigit():
                index += 1
            normalized.append(("number", line[token_start:index]))
            # A leading numeric label leaves the statement body at its start.
            statement_start = index < len(line) and line[index].isspace()
            continue

        normalized.append(("punctuation", char))
        index += 1
        statement_start = char == ":"

    return tuple(normalized)


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
    from xlvbatools.vba._io import read_vba_text
    content = read_vba_text(filepath)
    lines = content.split("\n")

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
