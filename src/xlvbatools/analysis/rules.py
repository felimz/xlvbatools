"""
VBA Static Analysis Rules
===========================
Configurable rule definitions for the VBA linter. Each rule is a function
that receives lines + context and yields VBAIssue instances.

Built-in rules:
    DS001  Dim after executable statement (VBA requires all Dim at top)
    CS001  Const after executable statement (same principle as DS001)
    LC001  Orphaned line continuation (missing space before _)
    SB001  Unbalanced Sub/Function...End blocks
    PF001  MsgBox call (should use Debug.Print or logging in headless mode)
    PF002  Implicit Variant (Dim/Const without As clause)
    PF003  ActiveSheet/ActiveCell usage (fragile, prefer explicit references)
    OE001  Option Explicit missing (with impact assessment)
    UV001  Undeclared variable usage (compile error with Option Explicit)
    BK001  Block-level variable declaration (hoisting vulnerability)
    SD002  Multiple variable declarations on a single line
    PF004  Avoid Integer data type (use Long to avoid silent promotion)
    SD005  Parameter passing modifier missing (explicit ByVal/ByRef required)
    SD006  Access modifier missing on procedure (implicitly Public)
    SD010  Line exceeds maximum length limit
    SD014  Avoid obsolete Call keyword for procedure invocation
    SC001  Hardcoded secret or credential literal warning
    SC002  Absolute file path usage
    CL001  Avoid obsolete Hungarian type suffixes (%, &, $, etc.)
    CL002  Fragile Selection/Select code patterns
    SF001  Silent error suppression (On Error Resume Next with no check)
    EH001  Missing error handler in Public procedure
    FD001  FileDialog without UserControl guard
    DC001  Unused local variable declaration
    DC002  Empty Sub or Function body
    DC003  Dead procedure with zero incoming calls (entry-point aware)
    SD015  Multiple consecutive blank lines
    SD016  Double-spaced code blocks (alternating blank lines)
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
_INTEGER_DECL_RE = re.compile(r"\bAs\s+Integer\b", re.IGNORECASE)
_CALL_RE = re.compile(r"^\s*Call\s+\w+", re.IGNORECASE)


# Block patterns
_BLOCK_START_PATTERNS = [
    (re.compile(r"^\s*If\s+.+?\s+Then\s*('|$)", re.IGNORECASE), "If"),
    (re.compile(r"^\s*For\s+", re.IGNORECASE), "For"),
    (re.compile(r"^\s*Do\b", re.IGNORECASE), "Do"),
    (re.compile(r"^\s*While\b", re.IGNORECASE), "While"),
    (re.compile(r"^\s*With\s+", re.IGNORECASE), "With"),
    (re.compile(r"^\s*Select\s+Case\b", re.IGNORECASE), "Select"),
]

_BLOCK_END_PATTERNS = [
    (re.compile(r"^\s*End\s+If\b", re.IGNORECASE), "If"),
    (re.compile(r"^\s*Next\b", re.IGNORECASE), "For"),
    (re.compile(r"^\s*Loop\b", re.IGNORECASE), "Do"),
    (re.compile(r"^\s*Wend\b", re.IGNORECASE), "While"),
    (re.compile(r"^\s*End\s+With\b", re.IGNORECASE), "With"),
    (re.compile(r"^\s*End\s+Select\b", re.IGNORECASE), "Select"),
]



def _get_logical_lines(lines: List[str]) -> List[tuple[int, str]]:
    """Group physical lines ending with ' _' into logical lines.

    Returns a list of tuples: (original_line_num, joined_line_text)
    """
    logical_lines = []
    current_parts = []
    start_line_num = None

    for i, raw_line in enumerate(lines, start=1):
        line_rstrip = raw_line.rstrip()

        # Check if the line ends with continuation character ' _'
        is_continuation = line_rstrip.endswith(" _")

        # Strip the trailing continuation character if present
        clean_line = line_rstrip[:-2] if is_continuation else line_rstrip

        if start_line_num is None:
            start_line_num = i

        current_parts.append(clean_line)

        if not is_continuation:
            joined_text = "".join(current_parts)
            logical_lines.append((start_line_num, joined_text))
            current_parts = []
            start_line_num = None

    if current_parts:
        logical_lines.append((start_line_num or len(lines), "".join(current_parts)))

    return logical_lines


def _parse_vba_declarations(line: str) -> List[tuple[str, str | None]]:
    """Parse a VBA declaration line (Dim, Private, Public, Const, Static, ReDim)
    and return a list of tuples: (var_name, var_type_or_none)
    """
    stripped = line.strip()
    if re.match(
        r"^(Public\s+|Private\s+|Friend\s+|Static\s+)?(Sub|Function|Property|Type|Enum)\s+",
        stripped,
        re.IGNORECASE,
    ):
        return []
    m = re.match(
        r"^(Dim|Private|Public|Static|ReDim|Const)\s+(.+)$", stripped, re.IGNORECASE
    )
    if not m:
        return []

    decl_body = m.group(2).strip()
    if decl_body.lower().startswith("const "):
        decl_body = decl_body[len("const "):].strip()
    if m.group(1).lower() == "redim" and decl_body.lower().startswith("preserve "):
        decl_body = decl_body[len("preserve "):].strip()

    # Split by comma taking paren depth into account
    parts = []
    current = []
    paren_depth = 0
    for char in decl_body:
        if char == "(":
            paren_depth += 1
        elif char == ")":
            paren_depth -= 1

        if char == "," and paren_depth == 0:
            parts.append("".join(current))
            current = []
        else:
            current.append(char)
    if current:
        parts.append("".join(current))

    results = []
    for part in parts:
        part = part.strip()
        var_match = re.match(r"^(\w+)", part)
        if var_match:
            var_name = var_match.group(1)
            # Remove array parentheses
            rest = re.sub(r"^\w+\s*\([^)]*\)", "", part).strip()
            type_match = re.search(r"\bAs\s+(\w+)", rest, re.IGNORECASE)
            var_type = type_match.group(1) if type_match else None
            results.append((var_name, var_type))
    return results


def check_dim_after_exec(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """DS001: Dim statement after executable code in a procedure."""
    issues = []
    in_proc = False
    proc_name = ""
    saw_exec = False

    logical_lines = _get_logical_lines(lines)

    for i, raw_line in logical_lines:
        line = raw_line.strip()

        if _PROC_START_RE.match(line):
            in_proc = True
            proc_name = _extract_proc_name(line)
            saw_exec = False

        if _PROC_END_RE.match(line):
            in_proc = False
            proc_name = ""
            continue

        if in_proc and _DIM_RE.match(line) and saw_exec:
            issues.append(VBAIssue(
                rule_id="DS001",
                severity="WARNING",
                module=rel_path,
                line_num=i,
                message=f"Dim after executable statement in procedure '{proc_name}'. "
                        f"VBA allocates all local variables at procedure startup regardless of where they are placed, so declaring them inline is misleading. "
                        f"ACTION: Move this declaration to the top of the '{proc_name}' procedure before any executable statements.",
            ))

        if in_proc and line and not _DIM_RE.match(line) and not line.startswith("'") and not line.lower().startswith("rem "):
            if not line.startswith("ReDim ") and not _PROC_END_RE.match(line):
                # Const is also a declaration -- don't count it as executable
                if not re.match(r"^(Public\s+|Private\s+)?Const\s+", line, re.IGNORECASE):
                    if not _PROC_START_RE.match(line):
                        saw_exec = True

    return issues


def check_const_after_exec(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """CS001: Const statement after executable code in a procedure."""
    issues = []
    in_proc = False
    proc_name = ""
    saw_exec = False
    _const_re = re.compile(r"^(Public\s+|Private\s+)?Const\s+", re.IGNORECASE)

    logical_lines = _get_logical_lines(lines)

    for i, raw_line in logical_lines:
        line = raw_line.strip()

        if _PROC_START_RE.match(line):
            in_proc = True
            proc_name = _extract_proc_name(line)
            saw_exec = False

        if _PROC_END_RE.match(line):
            in_proc = False
            proc_name = ""
            continue

        if in_proc and _const_re.match(line) and saw_exec:
            issues.append(VBAIssue(
                rule_id="CS001",
                severity="WARNING",
                module=rel_path,
                line_num=i,
                message=f"Const after executable statement in procedure '{proc_name}'. "
                        f"Like Dim, Const declarations belong at the top of a procedure for clarity. "
                        f"ACTION: Move this Const to the top of the '{proc_name}' procedure before any executable statements.",
            ))

        if in_proc and line and not _DIM_RE.match(line) and not _const_re.match(line):
            if not line.startswith("'") and not line.lower().startswith("rem "):
                if not line.startswith("ReDim ") and not _PROC_END_RE.match(line):
                    if not _PROC_START_RE.match(line):
                        saw_exec = True

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
                    message="Possible orphaned line continuation. In VBA, a continuation character ('_') must be preceded by exactly one space to form a valid multiline wrap. "
                            "ACTION: Prepend a single space before the underscore character (e.g., change 'x_' to 'x _').",
                ))
    return issues


def check_unbalanced_blocks(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """SB001: Unbalanced Sub/Function...End blocks."""
    issues = []
    stack = []
    logical_lines = _get_logical_lines(lines)

    for i, raw_line in logical_lines:
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
                    message="Mismatched closure: 'End Sub/Function' without a matching start statement. "
                            "ACTION: Ensure all procedures open and close correctly, or remove this duplicate closure.",
                ))

    for start_line, name in stack:
        issues.append(VBAIssue(
            rule_id="SB001",
            severity="ERROR",
            module=rel_path,
            line_num=start_line,
            message=f"Unbalanced procedure: Sub/Function '{name}' lacks a corresponding 'End Sub' or 'End Function'. "
                    f"ACTION: Append a matching closure statement at the end of the procedure.",
        ))

    return issues


def check_msgbox(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """PF001: MsgBox calls (may cause hangs in headless mode)."""
    issues = []
    logical_lines = _get_logical_lines(lines)
    for i, raw_line in logical_lines:
        line = raw_line.strip()
        if line.startswith("'") or line.lower().startswith("rem "):
            continue
        if _MSGBOX_RE.search(line):
            # Check if MsgBox is guarded with Application.UserControl in preceding 3 physical lines
            is_guarded = False
            for lookback in range(1, 4):
                prev_idx = i - 1 - lookback
                if prev_idx >= 0:
                    prev_line = lines[prev_idx].strip()
                    if "Application.UserControl" in prev_line:
                        is_guarded = True
                        break
            if not is_guarded:
                issues.append(VBAIssue(
                    rule_id="PF001",
                    severity="WARNING",
                    module=rel_path,
                    line_num=i,
                    message="Interactive MsgBox call detected. Raw modal dialogs freeze headless COM automation runs because there is no desktop user to click OK. "
                            "ACTION: Place an Application.UserControl guard before this line (e.g., 'If Not Application.UserControl Then Exit Sub') or write to a log file instead.",
                ))
    return issues


def check_implicit_variant(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """PF002: Dim or Const without explicit type (creates implicit Variant)."""
    issues = []
    logical_lines = _get_logical_lines(lines)
    for i, raw_line in logical_lines:
        line = raw_line.strip()
        if line.startswith("'") or line.lower().startswith("rem "):
            continue

        is_redim = re.match(r"^\s*ReDim\b", line, re.IGNORECASE) is not None
        if is_redim:
            continue

        # Check Const without type: "Const X = 1" (missing As clause)
        is_const = re.match(r"^\s*(Public\s+|Private\s+)?Const\b", line, re.IGNORECASE) is not None
        if is_const:
            # Extract const name(s) -- pattern: Const Name [As Type] = Value
            const_match = re.match(
                r"^\s*(?:Public\s+|Private\s+)?Const\s+(\w+)\s*(As\s+\w+)?\s*=",
                line, re.IGNORECASE,
            )
            if const_match and const_match.group(2) is None:
                const_name = const_match.group(1)
                if const_name.lower() not in VBA_RESERVED:
                    issues.append(VBAIssue(
                        rule_id="PF002",
                        severity="WARNING",
                        module=rel_path,
                        line_num=i,
                        message=f"Implicit type on Const: '{const_name}' lacks an explicit 'As' type clause. "
                                f"VBA will infer a type from the literal value, but explicit typing prevents ambiguity. "
                                f"ACTION: Add an explicit type (e.g., 'Const {const_name} As Long = ...').",
                    ))
            continue

        # Check Dim without type
        for var_name, var_type in _parse_vba_declarations(line):
            if var_type is None and var_name.lower() not in VBA_RESERVED:
                issues.append(VBAIssue(
                    rule_id="PF002",
                    severity="WARNING",
                    module=rel_path,
                    line_num=i,
                    message=f"Implicit Variant declaration: Variable '{var_name}' lacks an explicit 'As' type clause and will default to Variant. "
                            f"This increases memory overhead and slows execution. ACTION: Declare the variable with an explicit type (e.g., 'Dim {var_name} As Long' or 'Dim {var_name} As Double').",
                ))
    return issues


def check_active_refs(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """PF003: ActiveSheet/ActiveCell/ActiveWorkbook usage."""
    issues = []
    logical_lines = _get_logical_lines(lines)
    for i, raw_line in logical_lines:
        line = raw_line.strip()
        if line.startswith("'") or line.lower().startswith("rem "):
            continue
        match = _ACTIVESHEET_RE.search(line)
        if match:
            issues.append(VBAIssue(
                rule_id="PF003",
                severity="WARNING",
                module=rel_path,
                line_num=i,
                message=f"Active sheet or focus dependency reference '{match.group()}' detected. Focus-dependent references can lead to flaky automation bugs if the active focus shifts during runtime. "
                        f"ACTION: Prefer explicit sheet qualification (e.g., change to 'ThisWorkbook.Sheets(\"SheetName\").Range(...)').",
            ))
    return issues


def check_block_declarations(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """BK001: Local variable declaration inside control-flow or block structures."""
    issues = []
    in_proc = False
    block_stack = []

    logical_lines = _get_logical_lines(lines)

    for i, raw_line in logical_lines:
        line = raw_line.strip()

        if _PROC_START_RE.match(line):
            in_proc = True
            block_stack = []
            continue

        if _PROC_END_RE.match(line):
            in_proc = False
            block_stack = []
            continue

        if not in_proc:
            continue

        # Check block starts
        matched_start = False
        for pattern, kind in _BLOCK_START_PATTERNS:
            if pattern.match(line):
                block_stack.append(kind)
                matched_start = True
                break

        if matched_start:
            continue

        # Check block ends
        matched_end = False
        for pattern, kind in _BLOCK_END_PATTERNS:
            if pattern.match(line):
                if block_stack and block_stack[-1] == kind:
                    block_stack.pop()
                elif kind in block_stack:
                    while block_stack and block_stack[-1] != kind:
                        block_stack.pop()
                    if block_stack:
                        block_stack.pop()
                matched_end = True
                break

        if matched_end:
            continue

        # Check for Dim inside blocks
        if _DIM_RE.match(line) and block_stack:
            for var_name, _ in _parse_vba_declarations(line):
                issues.append(VBAIssue(
                    rule_id="BK001",
                    severity="WARNING",
                    module=rel_path,
                    line_num=i,
                    message=f"Local variable '{var_name}' declared inside a '{block_stack[-1]}' block. VBA does not support block-level scope, so declaring variables near-use inside loops or conditions falsely implies block limits. "
                            f"ACTION: Move the declaration of '{var_name}' to the top of the procedure outside of all block structures.",
                ))

    return issues


def check_option_explicit(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """OE001: Option Explicit missing (with impact assessment)."""
    has_option_explicit = any(_OPTION_EXPLICIT_RE.match(line.strip()) for line in lines)

    if has_option_explicit:
        return []

    # Check if module has any variable usage (Dim, assignments, For loops)
    has_vars = False
    for line in lines:
        stripped = line.strip()
        if _DIM_RE.match(stripped):
            has_vars = True
            break
        if re.match(r"^\w+\s*=", stripped) and not re.match(r"^(Sub|Function|Property|End|If|Do|For|While|Set|Let|Const|Dim|Public|Private|Option|Attribute)\b", stripped, re.IGNORECASE):
            has_vars = True
            break
        if re.match(r"^For\s+\w+\s*=", stripped, re.IGNORECASE):
            has_vars = True
            break

    if not has_vars:
        return []

    # Impact assessment: find undeclared variables
    undeclared = _find_undeclared_variables(lines)
    impact = ""
    if undeclared:
        var_list = ", ".join(sorted(undeclared)[:10])
        impact = f" IMPACT: {len(undeclared)} undeclared variable(s) found ({var_list}) that will need Dim statements before Option Explicit can be safely added."

    return [VBAIssue(
        rule_id="OE001",
        severity="ERROR",
        module=rel_path,
        line_num=0,
        message="Option Explicit statement is missing. VBA implicitly creates variables upon first assignment, allowing typos to compile silently and cause runtime bugs. "
                f"ACTION: Add 'Option Explicit' at the absolute top of the module.{impact}",
    )]


def check_one_declaration_per_line(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """SD002: One variable declaration per line (commas in Dim are forbidden)."""
    issues = []
    logical_lines = _get_logical_lines(lines)
    for i, raw_line in logical_lines:
        line = raw_line.strip()
        if line.startswith("'") or line.lower().startswith("rem "):
            continue
        m = re.match(r"^(Dim|Private|Public|Static|ReDim)\s+(.+)$", line, re.IGNORECASE)
        if m:
            decl_body = m.group(2).strip()
            if re.match(r"^(Sub|Function|Property|Type|Enum|Const)\b", decl_body, re.IGNORECASE):
                continue
            paren_depth = 0
            comma_count = 0
            for char in decl_body:
                if char == '(':
                    paren_depth += 1
                elif char == ')':
                    paren_depth -= 1
                elif char == ',' and paren_depth == 0:
                    comma_count += 1
            if comma_count > 0:
                issues.append(VBAIssue(
                    rule_id="SD002",
                    severity="STYLE",
                    module=rel_path,
                    line_num=i,
                    message="Multiple variable declarations on one line. Commas in Dim statements are forbidden to prevent typing misconceptions (e.g., in 'Dim a, b As Long', only 'b' is a Long; 'a' is a Variant). "
                            "ACTION: Declare each variable on a separate line.",
                ))
    return issues


def check_avoid_integer(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """PF004: Avoid Integer, Prefer Long."""
    issues = []
    logical_lines = _get_logical_lines(lines)
    for i, raw_line in logical_lines:
        line = raw_line.strip()
        if line.startswith("'") or line.lower().startswith("rem "):
            continue
        if _INTEGER_DECL_RE.search(line):
            issues.append(VBAIssue(
                rule_id="PF004",
                severity="WARNING",
                module=rel_path,
                line_num=i,
                message="Obsolete 'Integer' type usage. Integer (16-bit) is converted to Long (32-bit) during execution anyway and overflows easily on rows >32,767. "
                        "ACTION: Replace 'As Integer' with 'As Long' or 'As Double'.",
            ))
    return issues


def check_explicit_param_passing(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """SD005: All procedure arguments must explicitly declare ByVal or ByRef."""
    issues = []
    logical_lines = _get_logical_lines(lines)
    for i, raw_line in logical_lines:
        line = raw_line.strip()
        if line.startswith("'") or line.lower().startswith("rem "):
            continue
        if _PROC_START_RE.match(line):
            first_paren = line.find("(")
            last_paren = line.rfind(")")
            if first_paren != -1 and last_paren != -1:
                params_str = line[first_paren + 1:last_paren].strip()
                if not params_str:
                    continue
                # Split params by comma, tracking nested parenthesis (for array dimensions)
                params = []
                current = []
                paren_depth = 0
                for char in params_str:
                    if char == "(":
                        paren_depth += 1
                    elif char == ")":
                        paren_depth -= 1
                    if char == "," and paren_depth == 0:
                        params.append("".join(current).strip())
                        current = []
                    else:
                        current.append(char)
                if current:
                    params.append("".join(current).strip())

                for param in params:
                    if not param:
                        continue
                    # Strip Optional / ParamArray prefix
                    for prefix in ("optional ", "paramarray "):
                        if param.lower().startswith(prefix):
                            param = param[len(prefix):].strip()
                            break
                    p_lower = param.lower()
                    if not (p_lower.startswith("byval ") or p_lower.startswith("byref ")):
                        p_name = param.split()[0] if param.split() else param
                        issues.append(VBAIssue(
                            rule_id="SD005",
                            severity="STYLE",
                            module=rel_path,
                            line_num=i,
                            message=f"Missing parameter modifier: Argument '{p_name}' does not specify ByVal or ByRef, defaulting implicitly to ByRef (passing a mutable pointer). "
                                    f"ACTION: Add explicit 'ByVal' (if values shouldn't change) or 'ByRef' (if they should) before the parameter name.",
                        ))
    return issues


def check_explicit_access_modifiers(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """SD006: All procedures and module-level declarations must have explicit access modifiers."""
    issues = []
    in_proc = False
    
    logical_lines = _get_logical_lines(lines)
    for i, raw_line in logical_lines:
        line = raw_line.strip()
        if line.startswith("'") or line.lower().startswith("rem ") or not line:
            continue
            
        if _PROC_START_RE.match(line):
            in_proc = True
            first_word = line.split()[0].lower()
            if first_word not in ("public", "private", "friend"):
                proc_name = _extract_proc_name(line)
                issues.append(VBAIssue(
                    rule_id="SD006",
                    severity="STYLE",
                    module=rel_path,
                    line_num=i,
                    message=f"Missing access modifier: Procedure '{proc_name}' does not specify Public or Private, defaulting implicitly to Public and polluting the namespace. "
                            f"ACTION: Add explicit 'Public' or 'Private' keyword before Sub/Function/Property.",
                ))
            continue
            
        if _PROC_END_RE.match(line):
            in_proc = False
            continue
            
        if not in_proc:
            if line.lower().startswith("dim "):
                issues.append(VBAIssue(
                    rule_id="SD006",
                    severity="STYLE",
                    module=rel_path,
                    line_num=i,
                    message="Missing access modifier: Module-level variable declaration lacks explicit 'Public' or 'Private' modifier. "
                            "ACTION: Replace 'Dim' with 'Private' or 'Public' (e.g., 'Private moduleVar As Long').",
                ))
            elif line.lower().startswith("const "):
                issues.append(VBAIssue(
                    rule_id="SD006",
                    severity="STYLE",
                    module=rel_path,
                    line_num=i,
                    message="Missing access modifier: Module-level constant lacks explicit 'Public' or 'Private' modifier. "
                            "ACTION: Add explicit 'Public' or 'Private' before Const keyword (e.g., 'Private Const c_MyConst = 1').",
                ))
    return issues


def check_line_length(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """SD010: Lines should not exceed 120 characters."""
    issues = []
    for i, raw_line in enumerate(lines, start=1):
        line = raw_line.rstrip()
        if len(line) > 120:
            issues.append(VBAIssue(
                rule_id="SD010",
                severity="STYLE",
                module=rel_path,
                line_num=i,
                message=f"Line exceeds maximum length ({len(line)} chars). Horizontal scrolling degrades code readability in VBE. "
                        f"ACTION: Use line continuation character ' _' to wrap statements onto multiple lines.",
            ))
    return issues


def check_no_call_keyword(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """SD014: Omit the Call keyword."""
    issues = []
    logical_lines = _get_logical_lines(lines)
    for i, raw_line in logical_lines:
        line = raw_line.strip()
        if line.startswith("'") or line.lower().startswith("rem "):
            continue
        if _CALL_RE.match(line):
            issues.append(VBAIssue(
                rule_id="SD014",
                severity="STYLE",
                module=rel_path,
                line_num=i,
                message="Obsolete 'Call' keyword used. Procedure calls with Call require parentheses and add visual clutter. "
                        "ACTION: Remove 'Call' and parenthesis around arguments (e.g., change 'Call MySub(arg)' to 'MySub arg').",
            ))
    return issues


def check_hardcoded_secrets(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """SC001: Hardcoded passwords, credentials, or API keys."""
    issues = []
    secret_pattern = re.compile(
        r"\b(pwd|password|apikey|secret|api_key|token)\s*=\s*\".+?\"",
        re.IGNORECASE,
    )
    for i, raw_line in enumerate(lines, start=1):
        line = _strip_trailing_comment(raw_line)
        if secret_pattern.search(line):
            issues.append(VBAIssue(
                rule_id="SC001",
                severity="WARNING",
                module=rel_path,
                line_num=i,
                message="Possible hardcoded credential or secret detected. Storing raw passwords or API keys in code violates security compliance standards. ACTION: Move secrets to external configurations or environment variables.",
            ))
    return issues


def check_absolute_paths(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """SC002: Hardcoded absolute file paths."""
    issues = []
    path_pattern = re.compile(
        r'("[a-zA-Z]:\\(?:[^"\\]+\\)*[^"\\]*")|("\\\\[a-zA-Z0-9_.-]+\\[^"]+")'
    )
    for i, raw_line in enumerate(lines, start=1):
        line = _strip_trailing_comment(raw_line)
        match = path_pattern.search(line)
        if match:
            matched_path = match.group(0)
            if "Users" in matched_path or "Desktop" in matched_path or len(matched_path) > 15:
                issues.append(VBAIssue(
                    rule_id="SC002",
                    severity="WARNING",
                    module=rel_path,
                    line_num=i,
                    message=f"Hardcoded absolute local path detected ({matched_path}). This will cause runtime failures when run on other machines. ACTION: Use relative paths, 'ThisWorkbook.Path', or environment variables.",
                ))
    return issues


def check_type_suffixes(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """CL001: Obsolete type declaration suffixes."""
    issues = []
    for i, raw_line in enumerate(lines, start=1):
        line = _strip_trailing_comment(raw_line)
        if not re.match(r"^(Dim|Private|Public|Static|ReDim|Const)\b", line, re.IGNORECASE):
            continue
        match = re.search(r"\b\w+([%&@!#$])", line)
        if match:
            char = match.group(1)
            issues.append(VBAIssue(
                rule_id="CL001",
                severity="STYLE",
                module=rel_path,
                line_num=i,
                message=f"Obsolete type declaration suffix '{char}' used. VBA allows suffixes for legacy support, but explicit type declarations are preferred. ACTION: Declare the type using 'As <Type>' (e.g. change 'x{char}' to 'x As ...').",
            ))
    return issues


def check_fragile_selection(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """CL002: Fragile selection/activation patterns."""
    issues = []
    selection_pattern = re.compile(
        r"\b(Selection|ActiveWindow|ActivePrinter)\b|\.(Select|Activate)\b",
        re.IGNORECASE,
    )
    for i, raw_line in enumerate(lines, start=1):
        line = _strip_trailing_comment(raw_line)
        if selection_pattern.search(line):
            issues.append(VBAIssue(
                rule_id="CL002",
                severity="STYLE",
                module=rel_path,
                line_num=i,
                message="Fragile Excel Selection or Activation pattern used. Relying on active/selected objects makes macros slow, hard to debug, and prone to breaking when user focus shifts. ACTION: Reference ranges and sheets explicitly and avoid .Select or .Activate.",
            ))
    return issues


def check_silent_error_suppression(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """SF001: Silent error suppression with On Error Resume Next."""
    issues = []
    for i, raw_line in enumerate(lines, start=1):
        line = _strip_trailing_comment(raw_line)
        if re.match(r"^On\s+Error\s+Resume\s+Next", line, re.IGNORECASE):
            verified = False
            for offset in range(1, 6):
                idx = i + offset - 1
                if idx >= len(lines):
                    break
                next_line = _strip_trailing_comment(lines[idx])
                if re.match(r"^On\s+Error\s+", next_line, re.IGNORECASE):
                    verified = True
                    break
                if "err." in next_line.lower() or "err" in next_line.lower():
                    verified = True
                    break
            if not verified:
                issues.append(VBAIssue(
                    rule_id="SF001",
                    severity="WARNING",
                    module=rel_path,
                    line_num=i,
                    message="Silent error suppression detected without immediate handling. Using 'On Error Resume Next' without checking Err.Number or resetting via 'On Error GoTo 0' can mask critical errors and cause silent data corruption. ACTION: Add error checking or reset handling immediately after.",
                ))
    return issues


def check_unused_local_variables(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """DC001: Unused local variables in procedures."""
    issues = []
    in_proc = False
    proc_name = ""
    proc_decls = []
    proc_lines = []

    logical_lines = _get_logical_lines(lines)

    def process_proc():
        if not proc_decls:
            return
        proc_body_text = ""
        for _, text in proc_lines:
            text = _strip_trailing_comment(text)
            text = re.sub(r'".*?"', '""', text)
            proc_body_text += " " + text

        tokens = re.findall(r"\b\w+\b", proc_body_text, re.IGNORECASE)
        counts = {}
        for t in tokens:
            counts[t.lower()] = counts.get(t.lower(), 0) + 1

        for line_num, var in proc_decls:
            if counts.get(var.lower(), 0) <= 1:
                sig_line = proc_lines[0][1] if proc_lines else ""
                if re.search(rf"\b{re.escape(var)}\b", sig_line, re.IGNORECASE) and not re.search(rf"\bDim\s+{re.escape(var)}\b", sig_line, re.IGNORECASE):
                    continue
                issues.append(VBAIssue(
                    rule_id="DC001",
                    severity="WARNING",
                    module=rel_path,
                    line_num=line_num,
                    message=f"Unused local variable '{var}' declared in procedure '{proc_name}'. ACTION: Remove this declaration to clean up dead code.",
                ))

    for i, raw_line in logical_lines:
        line = raw_line.strip()
        
        if _PROC_START_RE.match(line):
            if in_proc:
                process_proc()
            in_proc = True
            proc_name = _extract_proc_name(line)
            proc_decls = []
            proc_lines = [(i, raw_line)]
            continue

        if _PROC_END_RE.match(line):
            if in_proc:
                proc_lines.append((i, raw_line))
                process_proc()
            in_proc = False
            proc_name = ""
            proc_decls = []
            proc_lines = []
            continue

        if in_proc:
            proc_lines.append((i, raw_line))
            decls = _parse_vba_declarations(line)
            for var_name, _ in decls:
                proc_decls.append((i, var_name))

    return issues


def check_empty_procedures(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """DC002: Empty Sub/Function/Property procedures."""
    issues = []
    in_proc = False
    proc_name = ""
    proc_start_line = 0
    proc_body_lines = 0

    logical_lines = _get_logical_lines(lines)
    for i, raw_line in logical_lines:
        line = raw_line.strip()
        if _PROC_START_RE.match(line):
            in_proc = True
            proc_name = _extract_proc_name(line)
            proc_start_line = i
            proc_body_lines = 0
            continue
        if _PROC_END_RE.match(line):
            if in_proc and proc_body_lines == 0:
                issues.append(VBAIssue(
                    rule_id="DC002",
                    severity="STYLE",
                    module=rel_path,
                    line_num=proc_start_line,
                    message=f"Empty procedure '{proc_name}' detected. Empty procedures occupy space without executing logic. ACTION: Implement the procedure or delete it if it is obsolete.",
                ))
            in_proc = False
            continue
        if in_proc:
            if line and not line.startswith("'") and not line.lower().startswith("rem "):
                if not _DIM_RE.match(line):
                    proc_body_lines += 1
    return issues


def check_consecutive_blank_lines(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """SD015: Multiple consecutive blank lines."""
    issues = []
    consecutive_blanks = 0
    for i, raw_line in enumerate(lines, start=1):
        line = raw_line.strip()
        if not line:
            consecutive_blanks += 1
            if consecutive_blanks >= 2:
                issues.append(VBAIssue(
                    rule_id="SD015",
                    severity="STYLE",
                    module=rel_path,
                    line_num=i,
                    message="Multiple consecutive blank lines detected. Excess whitespace increases vertical scrolling and clutter. ACTION: Collapse consecutive blank lines to a single blank line or run 'xlvba fmt'.",
                ))
        else:
            consecutive_blanks = 0
    return issues


def check_double_spacing(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """SD016: Double-spaced code blocks (alternating code and blank lines)."""
    issues = []
    logical_lines = _get_logical_lines(lines)
    in_proc = False
    proc_name = ""
    alt_sequence = []
    last_was_blank = False

    for i, raw_line in logical_lines:
        line = raw_line.strip()
        
        if _PROC_START_RE.match(line):
            if in_proc and len(alt_sequence) >= 3:
                issues.append(VBAIssue(
                    rule_id="SD016",
                    severity="STYLE",
                    module=rel_path,
                    line_num=alt_sequence[0],
                    message=f"Double-spaced code block detected in procedure '{proc_name}'. Alternating blank lines after every line of code degrades readability. ACTION: Remove alternating blank lines or run 'xlvba fmt'.",
                ))
            in_proc = True
            proc_name = _extract_proc_name(line)
            alt_sequence = []
            last_was_blank = False
            continue

        if _PROC_END_RE.match(line):
            if in_proc and len(alt_sequence) >= 3:
                issues.append(VBAIssue(
                    rule_id="SD016",
                    severity="STYLE",
                    module=rel_path,
                    line_num=alt_sequence[0],
                    message=f"Double-spaced code block detected in procedure '{proc_name}'. Alternating blank lines after every line of code degrades readability. ACTION: Remove alternating blank lines or run 'xlvba fmt'.",
                ))
            in_proc = False
            proc_name = ""
            alt_sequence = []
            last_was_blank = False
            continue

        if in_proc:
            if line.startswith("#"):
                continue
            is_blank = not line
            if is_blank:
                if not last_was_blank:
                    alt_sequence.append(i)
                    last_was_blank = True
                else:
                    if len(alt_sequence) >= 3:
                        issues.append(VBAIssue(
                            rule_id="SD016",
                            severity="STYLE",
                            module=rel_path,
                            line_num=alt_sequence[0],
                            message=f"Double-spaced code block detected in procedure '{proc_name}'. Alternating blank lines after every line of code degrades readability. ACTION: Remove alternating blank lines or run 'xlvba fmt'.",
                        ))
                    alt_sequence = []
                    last_was_blank = True
            else:
                if last_was_blank:
                    last_was_blank = False
                else:
                    if len(alt_sequence) >= 3:
                        issues.append(VBAIssue(
                            rule_id="SD016",
                            severity="STYLE",
                            module=rel_path,
                            line_num=alt_sequence[0],
                            message=f"Double-spaced code block detected in procedure '{proc_name}'. Alternating blank lines after every line of code degrades readability. ACTION: Remove alternating blank lines or run 'xlvba fmt'.",
                        ))
                    alt_sequence = []
                    last_was_blank = False

    return issues


# ── Entry-point patterns for DC003 whitelist ──

_ENTRY_POINT_PATTERNS = [
    re.compile(r"^(Worksheet_|Workbook_|CommandButton\d*_|UserForm_)", re.IGNORECASE),
    re.compile(r"_Click$", re.IGNORECASE),
    re.compile(r"_Change$", re.IGNORECASE),
    re.compile(r"_Initialize$", re.IGNORECASE),
    re.compile(r"_Terminate$", re.IGNORECASE),
    re.compile(r"_Open$", re.IGNORECASE),
    re.compile(r"_Close$", re.IGNORECASE),
    re.compile(r"_BeforeClose$", re.IGNORECASE),
    re.compile(r"_BeforeSave$", re.IGNORECASE),
    re.compile(r"_Activate$", re.IGNORECASE),
    re.compile(r"_Deactivate$", re.IGNORECASE),
    re.compile(r"_SelectionChange$", re.IGNORECASE),
    re.compile(r"_BeforeDoubleClick$", re.IGNORECASE),
    re.compile(r"_BeforeRightClick$", re.IGNORECASE),
]

# Common entry-point names called via Application.Run
_ENTRY_POINT_NAMES = {
    "main", "auto_open", "auto_close", "onretrieve", "oncalculate",
    "onvalidate", "initialize", "cleanup", "setup", "teardown",
}


def _is_entry_point(proc_name: str) -> bool:
    """Check if a procedure name matches known entry-point patterns."""
    for pattern in _ENTRY_POINT_PATTERNS:
        if pattern.search(proc_name):
            return True
    return proc_name.lower() in _ENTRY_POINT_NAMES


def check_undeclared_variables(rel_path: str, lines: List[str], global_names: set[str] | None = None) -> List[VBAIssue]:
    """UV001: Variables used but not declared (compile error with Option Explicit).

    Args:
        global_names: Optional set of lowercase variable names known to be
                      declared as Public in other modules (cross-module scope).
    """
    issues = []
    has_option_explicit = any(_OPTION_EXPLICIT_RE.match(line.strip()) for line in lines)

    # Only meaningful when Option Explicit is present
    if not has_option_explicit:
        return []

    undeclared_by_proc = _find_undeclared_variables_detailed(lines)
    known_globals = global_names or set()
    for proc_name, var_set in undeclared_by_proc.items():
        for var_name, line_num in sorted(var_set, key=lambda x: x[1]):
            # Skip variables known to be Public in other modules
            if var_name.lower() in known_globals:
                continue
            issues.append(VBAIssue(
                rule_id="UV001",
                severity="ERROR",
                module=rel_path,
                line_num=line_num,
                message=f"Undeclared variable '{var_name}' used in procedure '{proc_name}'. "
                        f"With Option Explicit enabled, this will cause a compile error. "
                        f"ACTION: Add 'Dim {var_name} As <Type>' at the top of the procedure.",
            ))

    return issues


def check_error_handler(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """EH001: Public procedures should have an error handler."""
    issues = []
    in_proc = False
    proc_name = ""
    proc_start_line = 0
    is_public = False
    has_error_handler = False

    logical_lines = _get_logical_lines(lines)

    for i, raw_line in logical_lines:
        line = raw_line.strip()

        if _PROC_START_RE.match(line):
            # Check previous proc
            if in_proc and is_public and not has_error_handler:
                if not _is_entry_point(proc_name) or True:  # Flag all public procs
                    issues.append(VBAIssue(
                        rule_id="EH001",
                        severity="WARNING",
                        module=rel_path,
                        line_num=proc_start_line,
                        message=f"Public procedure '{proc_name}' has no error handler. "
                                f"Unhandled errors in COM automation become modal dialogs that hang headless sessions. "
                                f"ACTION: Add 'On Error GoTo ErrHandler' at the top and an error handler block at the bottom.",
                    ))

            in_proc = True
            proc_name = _extract_proc_name(line)
            proc_start_line = i
            first_word = line.split()[0].lower()
            is_public = first_word not in ("private", "friend")
            has_error_handler = False
            continue

        if _PROC_END_RE.match(line):
            if in_proc and is_public and not has_error_handler:
                issues.append(VBAIssue(
                    rule_id="EH001",
                    severity="WARNING",
                    module=rel_path,
                    line_num=proc_start_line,
                    message=f"Public procedure '{proc_name}' has no error handler. "
                            f"Unhandled errors in COM automation become modal dialogs that hang headless sessions. "
                            f"ACTION: Add 'On Error GoTo ErrHandler' at the top and an error handler block at the bottom.",
                ))
            in_proc = False
            continue

        if in_proc:
            if re.match(r"^On\s+Error\s+GoTo\b", line, re.IGNORECASE):
                has_error_handler = True
            elif re.match(r"^On\s+Error\s+Resume\s+Next", line, re.IGNORECASE):
                has_error_handler = True

    return issues


def check_filedialog_guard(rel_path: str, lines: List[str]) -> List[VBAIssue]:
    """FD001: FileDialog usage without UserControl guard."""
    issues = []
    _fd_re = re.compile(r"\bApplication\.FileDialog\b|\bFileDialog\b", re.IGNORECASE)

    logical_lines = _get_logical_lines(lines)
    for i, raw_line in logical_lines:
        line = raw_line.strip()
        if line.startswith("'") or line.lower().startswith("rem "):
            continue
        if _fd_re.search(line):
            # Check if UserControl guard exists in preceding lines
            is_guarded = False
            for lookback in range(1, 8):
                prev_idx = i - 1 - lookback
                if prev_idx >= 0:
                    prev_line = lines[prev_idx].strip()
                    if "Application.UserControl" in prev_line:
                        is_guarded = True
                        break
            if not is_guarded:
                issues.append(VBAIssue(
                    rule_id="FD001",
                    severity="WARNING",
                    module=rel_path,
                    line_num=i,
                    message="Application.FileDialog call without UserControl guard. "
                            "File dialogs freeze headless COM automation because there is no desktop user to interact. "
                            "ACTION: Add 'If Not Application.UserControl Then Exit Sub' before this line.",
                ))
                break  # Only flag once per module to avoid noise

    return issues


# ── Rule Registry ──

ALL_RULES: dict[str, Callable] = {
    "DS001": check_dim_after_exec,
    "CS001": check_const_after_exec,
    "LC001": check_line_continuation,
    "SB001": check_unbalanced_blocks,
    "PF001": check_msgbox,
    "PF002": check_implicit_variant,
    "PF003": check_active_refs,
    "OE001": check_option_explicit,
    "UV001": check_undeclared_variables,
    "BK001": check_block_declarations,
    "SD002": check_one_declaration_per_line,
    "PF004": check_avoid_integer,
    "SD005": check_explicit_param_passing,
    "SD006": check_explicit_access_modifiers,
    "SD010": check_line_length,
    "SD014": check_no_call_keyword,
    "SC001": check_hardcoded_secrets,
    "SC002": check_absolute_paths,
    "CL001": check_type_suffixes,
    "CL002": check_fragile_selection,
    "SF001": check_silent_error_suppression,
    "EH001": check_error_handler,
    "FD001": check_filedialog_guard,
    "DC001": check_unused_local_variables,
    "DC002": check_empty_procedures,
    "SD015": check_consecutive_blank_lines,
    "SD016": check_double_spacing,
}


def run_all_rules(
    rel_path: str,
    lines: List[str],
    disabled_rules: List[str] | None = None,
    global_names: set[str] | None = None,
) -> List[VBAIssue]:
    """Run all enabled rules against a file's lines.

    Args:
        global_names: Optional set of lowercase variable names known to be
                      declared as Public in other modules (for UV001 cross-module check).
    """
    disabled = set(disabled_rules or [])
    issues = []
    for rule_id, check_fn in ALL_RULES.items():
        if rule_id in disabled:
            continue
        # UV001 needs cross-module context
        if rule_id == "UV001" and global_names is not None:
            issues.extend(check_fn(rel_path, lines, global_names))
        else:
            issues.extend(check_fn(rel_path, lines))

    # Map line numbers to containing procedures
    proc_map = {}
    current_proc = None
    for i, raw_line in enumerate(lines, start=1):
        line = raw_line.strip()
        if _PROC_START_RE.match(line):
            current_proc = _extract_proc_name(line)
        proc_map[i] = current_proc
        if _PROC_END_RE.match(line):
            current_proc = None

    # Attach procedure context to issues
    for issue in issues:
        if issue.line_num and issue.line_num in proc_map:
            issue.procedure = proc_map[issue.line_num]

    return issues


# ── Helpers ──

def _extract_proc_name(line: str) -> str:
    """Extract procedure name from a Sub/Function declaration line."""
    # Strip access modifiers (case-insensitive since VBA is case-insensitive)
    line_lower = line.lower()
    for prefix in ("public ", "private ", "friend "):
        if line_lower.startswith(prefix):
            line = line[len(prefix):]
            line_lower = line.lower()
            break
    # Strip Sub/Function/Property keyword
    for kw in ("sub ", "function ", "property get ", "property let ", "property set "):
        if line_lower.startswith(kw):
            line = line[len(kw):]
            break
    # Extract name (before parenthesis or end of line)
    name = line.split("(")[0].strip() if "(" in line else line.strip()
    return name


def _inside_string(line: str) -> bool:
    """Check if the trailing underscore is inside a string literal.

    VBA uses doubled quotes ("") as escape sequences inside strings,
    so we skip over "" pairs rather than toggling in/out twice.
    """
    in_string = False
    i = 0
    while i < len(line):
        if line[i] == '"':
            if in_string:
                # Check for escaped quote ("")
                if i + 1 < len(line) and line[i + 1] == '"':
                    i += 2  # Skip the escaped pair
                    continue
                in_string = False
            else:
                in_string = True
        i += 1
    return in_string


def _strip_trailing_comment(line: str) -> str:
    """Strip comment at the end of a line, respecting string literals."""
    in_string = False
    i = 0
    while i < len(line):
        if line[i] == '"':
            if in_string and i + 1 < len(line) and line[i + 1] == '"':
                i += 2
                continue
            in_string = not in_string
        elif line[i] == "'" and not in_string:
            return line[:i].rstrip()
        elif line[i:i+4].lower() == "rem " and not in_string:
            return line[:i].rstrip()
        i += 1
    return line.rstrip()


# Common VBA built-in functions/objects NOT to treat as user variables
_VBA_BUILTINS = {
    "abs", "asc", "atn", "cbool", "cbyte", "ccur", "cdate", "cdbl", "chr",
    "cint", "clng", "cos", "csng", "cstr", "cvar", "date", "dateadd",
    "datediff", "datepart", "dateserial", "datevalue", "day", "dir", "eof",
    "err", "exp", "fix", "format", "freefile", "hex", "hour", "iif",
    "instr", "instrrev", "int", "isarray", "isdate", "isempty", "iserror",
    "ismissing", "isnull", "isnumeric", "isobject", "join", "lbound",
    "lcase", "left", "len", "log", "ltrim", "mid", "minute", "month",
    "now", "oct", "replace", "right", "rnd", "round", "rtrim",
    "second", "sgn", "sin", "space", "split", "sqr", "str", "strcomp",
    "string", "strreverse", "switch", "tab", "tan", "time", "timer",
    "timeserial", "timevalue", "trim", "typename", "ubound", "ucase",
    "val", "vartype", "weekday", "year",
    # Statements and keywords
    "debug", "print", "open", "close", "get", "put", "input", "write",
    "redim", "erase", "set", "let", "call", "exit", "on", "resume",
    "goto", "gosub", "return", "stop", "end", "nothing", "true", "false",
    "me", "new", "byval", "byref",
    # Common Excel objects
    "application", "thisworkbook", "activeworkbook", "activesheet",
    "activecell", "sheets", "worksheets", "range", "cells",
    "selection", "columns", "rows", "names",
}


def _find_undeclared_variables(lines: List[str]) -> set[str]:
    """Find variables used but not declared across all procedures.

    Returns a set of undeclared variable names.
    """
    all_undeclared = set()
    detailed = _find_undeclared_variables_detailed(lines)
    for proc_name, var_set in detailed.items():
        for var_name, _ in var_set:
            all_undeclared.add(var_name)
    return all_undeclared


def _find_undeclared_variables_detailed(lines: List[str]) -> dict[str, set[tuple[str, int]]]:
    """Find undeclared variables per procedure.

    Returns dict of proc_name -> set of (var_name, first_line_num).
    """
    result = {}
    in_proc = False
    proc_name = ""
    proc_declared = set()  # lowercase names
    proc_used = {}  # name -> first_line

    # First pass: collect module-level declarations (Public/Private/Dim outside procedures)
    module_level_names = set()
    in_any_proc = False
    _mod_decl_re = re.compile(
        r"^(Public|Private|Dim|Global)\s+", re.IGNORECASE
    )
    for raw_line in lines:
        line = raw_line.strip()
        if _PROC_START_RE.match(line):
            in_any_proc = True
        if _PROC_END_RE.match(line):
            in_any_proc = False
            continue
        if not in_any_proc and _mod_decl_re.match(line):
            # Skip Const at module level (handled separately)
            if re.match(r"^(Public\s+|Private\s+)?Const\s+", line, re.IGNORECASE):
                const_m = re.match(r"^(?:Public\s+|Private\s+)?Const\s+(\w+)", line, re.IGNORECASE)
                if const_m:
                    module_level_names.add(const_m.group(1).lower())
            else:
                for var_name, _ in _parse_vba_declarations(line):
                    module_level_names.add(var_name.lower())

    for i, raw_line in enumerate(lines, start=1):
        line = raw_line.strip()

        if _PROC_START_RE.match(line):
            # Process previous proc
            if in_proc and proc_name:
                undeclared = set()
                for name, first_line in proc_used.items():
                    if name.lower() not in proc_declared:
                        undeclared.add((name, first_line))
                if undeclared:
                    result[proc_name] = undeclared

            in_proc = True
            proc_name = _extract_proc_name(line)
            proc_declared = set(module_level_names)  # Seed with module-level declarations
            proc_used = {}

            # In VBA Functions and Property Gets, assigning to the function
            # name is the return-value idiom (e.g., "MyFunc = 42").
            # Treat the function name as implicitly declared.
            line_lower = line.lower().lstrip()
            for prefix in ("public ", "private ", "friend "):
                if line_lower.startswith(prefix):
                    line_lower = line_lower[len(prefix):]
                    break
            if line_lower.startswith("function ") or line_lower.startswith("property get "):
                proc_declared.add(proc_name.lower())

            # Extract parameter names as declared -- handle multi-line signatures
            # Join continuation lines to get the full procedure declaration
            full_sig = raw_line.rstrip()
            line_idx = i - 1  # 0-based index into lines array
            while full_sig.rstrip().endswith(" _") and line_idx + 1 < len(lines):
                line_idx += 1
                full_sig = full_sig.rstrip()[:-1] + " " + lines[line_idx].strip()
            first_paren = full_sig.find("(")
            last_paren = full_sig.rfind(")")
            if first_paren != -1 and last_paren != -1:
                params_str = full_sig[first_paren + 1:last_paren]
                for param in params_str.split(","):
                    param = param.strip()
                    for prefix in ("optional ", "paramarray ", "byval ", "byref "):
                        if param.lower().startswith(prefix):
                            param = param[len(prefix):].strip()
                    pname = param.split()[0] if param.split() else ""
                    # Strip array brackets: shiftNodes() -> shiftNodes
                    pname = pname.rstrip("()")
                    if pname:
                        proc_declared.add(pname.lower())
            continue

        if _PROC_END_RE.match(line):
            if in_proc and proc_name:
                undeclared = set()
                for name, first_line in proc_used.items():
                    if name.lower() not in proc_declared:
                        undeclared.add((name, first_line))
                if undeclared:
                    result[proc_name] = undeclared
            in_proc = False
            proc_name = ""
            proc_declared = set()
            proc_used = {}
            continue

        if not in_proc:
            continue

        # Skip comments
        if line.startswith("'") or line.lower().startswith("rem "):
            continue

        # VBA uses : as a statement separator. Split the line into sub-statements
        # but respect string literals (colons inside strings are not separators).
        sub_stmts = _split_colon_statements(line)

        for sub_stmt in sub_stmts:
            sub = sub_stmt.strip()
            if not sub or sub.startswith("'"):
                continue

            # Track declarations
            if _DIM_RE.match(sub) or re.match(r"^(Public\s+|Private\s+)?Const\s+", sub, re.IGNORECASE):
                for var_name, _ in _parse_vba_declarations(sub):
                    proc_declared.add(var_name.lower())
                # Also handle Const Name = Value
                const_m = re.match(r"^(?:Public\s+|Private\s+)?Const\s+(\w+)", sub, re.IGNORECASE)
                if const_m:
                    proc_declared.add(const_m.group(1).lower())
                continue

            # Track assignments: variable = ...
            assign_m = re.match(r"^(\w+)\s*=", sub)
            if assign_m:
                vname = assign_m.group(1)
                vlow = vname.lower()
                if (vlow not in VBA_RESERVED and vlow not in _VBA_BUILTINS
                        and vlow not in proc_declared
                        and vname not in proc_used):
                    proc_used[vname] = i

            # Track For loop variables: For x = ...
            for_m = re.match(r"^For\s+(\w+)\s*=", sub, re.IGNORECASE)
            if for_m:
                vname = for_m.group(1)
                vlow = vname.lower()
                if (vlow not in VBA_RESERVED and vlow not in _VBA_BUILTINS
                        and vlow not in proc_declared
                        and vname not in proc_used):
                    proc_used[vname] = i

    return result


def _split_colon_statements(line: str) -> list[str]:
    """Split a VBA line by colon statement separators, respecting string literals.

    In VBA, ':' separates multiple statements on one line, but colons
    inside string literals are not separators.
    """
    parts = []
    current = []
    in_string = False
    for ch in line:
        if ch == '"':
            in_string = not in_string
            current.append(ch)
        elif ch == ':' and not in_string:
            parts.append("".join(current))
            current = []
        else:
            current.append(ch)
    if current:
        parts.append("".join(current))
    return parts

