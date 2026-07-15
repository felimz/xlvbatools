"""
VBA Code Formatter
====================
Non-destructive VBA code formatter that normalizes indentation, blank lines,
and comment alignment. Preserves all logic -- only whitespace changes.

Usage:
    from xlvbatools.vba.formatter import format_vba, format_file

    formatted = format_vba(source_code)
    format_file("vba_source/modules/modMain.bas", dry_run=False)
"""

import logging
import os
import re

from xlvbatools.vba.constants import VBE_HEADER_FORMAT_RE as _VBE_HEADER

logger = logging.getLogger(__name__)

# Indentation config
INDENT_SIZE = 4
INDENT = " " * INDENT_SIZE

# Patterns that increase indent (next line)
_INDENT_INCREASE = re.compile(
    r"^(Public\s+|Private\s+|Friend\s+)?"
    r"(Sub|Function|Property\s+(Get|Let|Set))\s+",
    re.IGNORECASE,
)
_IF_THEN_BLOCK = re.compile(r"^If\b.*\bThen\s*$", re.IGNORECASE)
_ELSE = re.compile(r"^(Else|ElseIf\b)", re.IGNORECASE)
_SELECT_CASE = re.compile(r"^Select\s+Case\b", re.IGNORECASE)
_CASE = re.compile(r"^Case\b", re.IGNORECASE)
_FOR_LOOP = re.compile(r"^For\s+", re.IGNORECASE)
_DO_LOOP = re.compile(r"^Do\b", re.IGNORECASE)
_WHILE_LOOP = re.compile(r"^While\b", re.IGNORECASE)
_WITH_BLOCK = re.compile(r"^With\b", re.IGNORECASE)
_TYPE_BLOCK = re.compile(r"^(Public\s+|Private\s+)?Type\s+", re.IGNORECASE)
_ENUM_BLOCK = re.compile(r"^(Public\s+|Private\s+)?Enum\s+", re.IGNORECASE)

# Patterns that decrease indent (current line)
_INDENT_DECREASE = re.compile(
    r"^(End\s+(Sub|Function|Property|If|Select|With|Type|Enum)|"
    r"Next\b|Loop\b|Wend\b)",
    re.IGNORECASE,
)

# Module-level declarations (no indent)
_MODULE_LEVEL = re.compile(
    r"^(Option\s+|Public\s+Const\s+|Private\s+Const\s+|"
    r"Public\s+Declare\s+|Private\s+Declare\s+|"
    r"Public\s+Enum\s+|Private\s+Enum\s+|"
    r"Public\s+Type\s+|Private\s+Type\s+|"
    r"Dim\s+|Public\s+|Private\s+|#)",
    re.IGNORECASE,
)


def format_vba(source: str, indent_size: int = 4) -> str:
    """
    Format VBA source code with consistent indentation.

    Parameters
    ----------
    source : str
        Raw VBA source code.
    indent_size : int
        Number of spaces per indent level (default 4).

    Returns
    -------
    str
        Formatted VBA source code.
    """
    indent_str = " " * indent_size
    lines = source.split("\n")
    output = []
    indent_level = 0
    in_proc = False
    prev_blank = False

    for raw_line in lines:
        line = raw_line.rstrip()
        stripped = line.strip()

        # Preserve blank lines (but collapse multiple blanks to one)
        if not stripped:
            if not prev_blank:
                output.append("")
                prev_blank = True
            continue
        prev_blank = False

        # VBE header lines: preserve as-is
        if _VBE_HEADER.match(stripped):
            output.append(line)
            continue

        # Check for indent decrease BEFORE applying indent
        if _INDENT_DECREASE.match(stripped):
            indent_level = max(0, indent_level - 1)

        # Else/ElseIf: decrease one level temporarily
        if _ELSE.match(stripped):
            indent_level = max(0, indent_level - 1)

        # Case: align with Select (one less than body)
        if _CASE.match(stripped) and not _SELECT_CASE.match(stripped):
            current_indent = max(0, indent_level - 1)
        else:
            current_indent = indent_level

        # Apply indentation
        if indent_level == 0 and not in_proc:
            # Module-level: no indent
            formatted_line = stripped
        else:
            formatted_line = (indent_str * current_indent) + stripped

        output.append(formatted_line)

        # Check for indent increase AFTER applying indent
        if _INDENT_INCREASE.match(stripped):
            indent_level += 1
            in_proc = True

        if _IF_THEN_BLOCK.match(stripped):
            indent_level += 1
        elif _SELECT_CASE.match(stripped):
            indent_level += 1
        elif _FOR_LOOP.match(stripped):
            indent_level += 1
        elif _DO_LOOP.match(stripped):
            indent_level += 1
        elif _WHILE_LOOP.match(stripped):
            indent_level += 1
        elif _WITH_BLOCK.match(stripped):
            indent_level += 1
        elif _TYPE_BLOCK.match(stripped):
            indent_level += 1
        elif _ENUM_BLOCK.match(stripped):
            indent_level += 1

        # Else/ElseIf: restore indent for body
        if _ELSE.match(stripped):
            indent_level += 1

        # Track proc exit
        if re.match(r"^End\s+(Sub|Function|Property)", stripped, re.IGNORECASE):
            in_proc = False
            indent_level = 0

    # Ensure trailing newline
    result = "\n".join(output)
    if not result.endswith("\n"):
        result += "\n"
    return result


def format_file(
    filepath: str,
    dry_run: bool = False,
    indent_size: int = 4,
) -> dict:
    """
    Format a VBA source file in place.

    Parameters
    ----------
    filepath : str
        Path to the .bas or .cls file.
    dry_run : bool
        If True, return the diff without modifying the file.
    indent_size : int
        Spaces per indent level.

    Returns
    -------
    dict
        Keys: changed (bool), lines_changed (int), diff (str if dry_run)
    """
    import difflib

    from xlvbatools.vba._io import read_vba_text
    original = read_vba_text(filepath)

    formatted = format_vba(original, indent_size=indent_size)

    if original == formatted:
        return {"changed": False, "lines_changed": 0}

    orig_lines = original.splitlines(keepends=True)
    fmt_lines = formatted.splitlines(keepends=True)

    diff_lines = list(difflib.unified_diff(
        orig_lines, fmt_lines,
        fromfile=os.path.basename(filepath),
        tofile=os.path.basename(filepath) + " (formatted)",
    ))
    lines_changed = sum(
        1
        for line in diff_lines
        if line.startswith("+") and not line.startswith("+++")
    )

    if dry_run:
        return {
            "changed": True,
            "lines_changed": lines_changed,
            "diff": "".join(diff_lines),
        }

    with open(filepath, "w", encoding="utf-8", newline="\r\n") as f:
        f.write(formatted)

    logger.info(f"Formatted: {filepath} ({lines_changed} lines changed)")
    return {"changed": True, "lines_changed": lines_changed}


def format_directory(
    source_dir: str,
    dry_run: bool = False,
    indent_size: int = 4,
    extensions: tuple = (".bas", ".cls"),
) -> list[dict]:
    """Format all VBA files in a directory."""
    results = []
    for root, _, files in os.walk(source_dir):
        for fname in sorted(files):
            if not fname.endswith(extensions):
                continue
            filepath = os.path.join(root, fname)
            result = format_file(filepath, dry_run=dry_run, indent_size=indent_size)
            result["file"] = os.path.relpath(filepath, source_dir)
            results.append(result)
    return results
