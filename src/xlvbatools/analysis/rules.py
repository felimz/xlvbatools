"""
VBA Static Analysis Rules
===========================
Configurable rule definitions for the VBA linter. Each rule is a function
that receives lines + context and yields VBAIssue instances.

Built-in rules:
    DS001  Dim after executable statement (VBA requires all Dim at top)
    LC001  Orphaned line continuation (missing space before _)
    SB001  Unbalanced Sub/Function...End blocks
    PF001  MsgBox call (should use Debug.Print or logging in headless mode)
    PF002  Implicit Variant (Dim without As clause)
    PF003  ActiveSheet/ActiveCell usage (fragile, prefer explicit references)
    OE001  Option Explicit missing
    OP001  Executable code outside any procedure
"""

import re
import logging
from typing import List, Callable

from xlvbatools.analysis.issue import VBAIssue

logger = logging.getLogger(__name__)

# VBA reserved keywords that should not be used as variable names
VBA_RESERVED = {
    "and", "as", "boolean", "byref", "byte", "byval", "call", "case", "class",
    "const", "currency", "date", "debug", "declare", "dim", "do", "double",
    "each", "else", "elseif", "empty", "end", "enum", "erase", "error",
    "event", "exit", "false", "for", "friend", "function", "get", "global",
    "goto", "if", "implements", "in", "integer", "is", "let", "lib", "like",
    "long", "loop", "lset", "me", "mod", "new", "next", "not", "nothing",
    "null", "object", "on", "option", "optional", "or", "paramarray",
    "preserve", "print", "private", "property", "public", "raiseevent",
    "redim", "rem", "resume", "rset", "select", "set", "single", "static",
    "step", "stop", "string", "sub", "then", "to", "true", "type",
    "typeof", "until", "variant", "wend", "while", "with", "xor",
}

# Proc start patterns
_PROC_START_RE = re.compile(
    r"^(Public\s+|Private\s+|Friend\s+)?(Sub|Function|Property\s+(Get|Let|Set))\s+",
    re.IGNORECASE,
)
_PROC_END_RE = re.compile(r"^End\s+(Sub|Function|Property)", re.IGNORECASE)

# Dim pattern
_DIM_RE = re.compile(r"^\s*Dim\s+", re.IGNORECASE)
_DIM_AS_RE = re.compile(r"\bAs\s+\w+", re.IGNORECASE)

# Other patterns
_OPTION_EXPLICIT_RE = re.compile(r"^Option\s+Explicit", re.IGNORECASE)
_MSGBOX_RE = re.compile(r"\bMsgBox\b", re.IGNORECASE)
_ACTIVESHEET_RE = re.compile(r"\b(ActiveSheet|ActiveCell|ActiveWorkbook)\b")
_LINE_CONT_RE = re.compile(r"[^ ]_\s*$")


def check_dim_after_exec(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """DS001: Dim statement after executable code in a procedure."""
    issues = []
    in_proc = False
    proc_name = ""
    dim_after_exec = False

    for i, raw_line in enumerate(lines, start=1):
        line = raw_line.strip()

        if _PROC_START_RE.match(line):
            in_proc = True
            proc_name = _extract_proc_name(line)
            dim_after_exec = False

        if _PROC_END_RE.match(line):
            in_proc = False
            proc_name = ""
            continue

        if in_proc and _DIM_RE.match(line) and dim_after_exec:
            issues.append(VBAIssue(
                rule_id="DS001",
                severity="ERROR",
                module=rel_path,
                line_num=i,
                message=f"Dim after executable code in {proc_name} "
                        f"(VBA requires all Dim at top of procedure)",
            ))

        if in_proc and line and not _DIM_RE.match(line) and not line.startswith("'"):
            if not line.startswith("ReDim ") and not _PROC_END_RE.match(line):
                if not _PROC_START_RE.match(line):
                    dim_after_exec = True

    return issues


def check_line_continuation(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """LC001: Orphaned line continuation (no space before _)."""
    issues = []
    for i, raw_line in enumerate(lines, start=1):
        line = raw_line.rstrip()
        if line.endswith("_") and not line.endswith(" _"):
            # Skip if the underscore is inside a string
            if not _inside_string(line):
                issues.append(VBAIssue(
                    rule_id="LC001",
                    severity="WARNING",
                    module=rel_path,
                    line_num=i,
                    message="Possible orphaned line continuation (no space before _)",
                ))
    return issues


def check_unbalanced_blocks(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """SB001: Unbalanced Sub/Function...End blocks."""
    issues = []
    stack = []

    for i, raw_line in enumerate(lines, start=1):
        line = raw_line.strip()
        if _PROC_START_RE.match(line):
            name = _extract_proc_name(line)
            stack.append((i, name))
        elif _PROC_END_RE.match(line):
            if stack:
                stack.pop()
            else:
                issues.append(VBAIssue(
                    rule_id="SB001",
                    severity="ERROR",
                    module=rel_path,
                    line_num=i,
                    message=f"End Sub/Function without matching start",
                ))

    for start_line, name in stack:
        issues.append(VBAIssue(
            rule_id="SB001",
            severity="ERROR",
            module=rel_path,
            line_num=start_line,
            message=f"Sub/Function '{name}' has no matching End statement",
        ))

    return issues


def check_msgbox(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """PF001: MsgBox calls (may cause hangs in headless mode)."""
    issues = []
    for i, raw_line in enumerate(lines, start=1):
        line = raw_line.strip()
        if line.startswith("'"):
            continue
        if _MSGBOX_RE.search(line):
            issues.append(VBAIssue(
                rule_id="PF001",
                severity="WARNING",
                module=rel_path,
                line_num=i,
                message="MsgBox call detected (will hang in headless COM mode)",
            ))
    return issues


def check_implicit_variant(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """PF002: Dim without explicit type (creates implicit Variant)."""
    issues = []
    for i, raw_line in enumerate(lines, start=1):
        line = raw_line.strip()
        if line.startswith("'"):
            continue
        if _DIM_RE.match(line) and not _DIM_AS_RE.search(line):
            issues.append(VBAIssue(
                rule_id="PF002",
                severity="WARNING",
                module=rel_path,
                line_num=i,
                message="Implicit Variant: Dim without 'As <Type>' clause",
            ))
    return issues


def check_active_refs(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """PF003: ActiveSheet/ActiveCell/ActiveWorkbook usage."""
    issues = []
    for i, raw_line in enumerate(lines, start=1):
        line = raw_line.strip()
        if line.startswith("'"):
            continue
        match = _ACTIVESHEET_RE.search(line)
        if match:
            issues.append(VBAIssue(
                rule_id="PF003",
                severity="WARNING",
                module=rel_path,
                line_num=i,
                message=f"'{match.group()}' usage (fragile, prefer explicit references)",
            ))
    return issues


def check_option_explicit(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """OE001: Option Explicit missing."""
    has_option_explicit = any(_OPTION_EXPLICIT_RE.match(line.strip()) for line in lines)
    has_dim = any(_DIM_RE.match(line.strip()) for line in lines)

    if not has_option_explicit and has_dim:
        return [VBAIssue(
            rule_id="OE001",
            severity="WARNING",
            module=rel_path,
            line_num=0,
            message="Option Explicit not found (variables may be implicitly declared)",
        )]
    return []


# ── Rule Registry ──

ALL_RULES: dict[str, Callable] = {
    "DS001": check_dim_after_exec,
    "LC001": check_line_continuation,
    "SB001": check_unbalanced_blocks,
    "PF001": check_msgbox,
    "PF002": check_implicit_variant,
    "PF003": check_active_refs,
    "OE001": check_option_explicit,
}


def run_all_rules(
    rel_path: str,
    lines: List[str],
    disabled_rules: List[str] | None = None,
) -> List[VBAIssue]:
    """Run all enabled rules against a file's lines."""
    disabled = set(disabled_rules or [])
    issues = []
    for rule_id, check_fn in ALL_RULES.items():
        if rule_id in disabled:
            continue
        issues.extend(check_fn(rel_path, lines))
    return issues


# ── Helpers ──

def _extract_proc_name(line: str) -> str:
    """Extract procedure name from a Sub/Function declaration line."""
    # Strip access modifiers
    for prefix in ("Public ", "Private ", "Friend "):
        if line.startswith(prefix):
            line = line[len(prefix):]
            break
    # Strip Sub/Function/Property keyword
    for kw in ("Sub ", "Function ", "Property Get ", "Property Let ", "Property Set "):
        if line.startswith(kw):
            line = line[len(kw):]
            break
    # Extract name (before parenthesis or end of line)
    name = line.split("(")[0].strip() if "(" in line else line.strip()
    return name


def _inside_string(line: str) -> bool:
    """Rough check if the trailing underscore is inside a string literal."""
    in_string = False
    for ch in line:
        if ch == '"':
            in_string = not in_string
    return in_string
