"""
VBA Dependency / Call-Graph Analyzer
=====================================
Parses Sub/Function definitions and call sites across all VBA modules
to build a call dependency graph. Outputs as Mermaid, JSON adjacency list,
or DOT format.

Works without COM -- operates on extracted .bas/.cls files.

Usage:
    from xlvbatools.vba.dependency import build_call_graph, render_mermaid

    graph = build_call_graph("vba_source/")
    print(render_mermaid(graph))
"""

import logging
import os
import re
from dataclasses import dataclass, field
from typing import List, Dict, Set

logger = logging.getLogger(__name__)

# Patterns
_PROC_DEF_RE = re.compile(
    r"^(Public\s+|Private\s+|Friend\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+"
    r"(\w+)",
    re.IGNORECASE,
)

_CALL_RE = re.compile(r"\b(\w+)\s*(?:\(|$)", re.IGNORECASE)

# Common VBA built-in functions to exclude from the graph
_BUILTINS = {
    "abs", "asc", "atn", "cbool", "cbyte", "ccur", "cdate", "cdbl", "chr",
    "cint", "clng", "cos", "csng", "cstr", "cvar", "date", "dateadd",
    "datediff", "datepart", "dateserial", "datevalue", "day", "dir", "eof",
    "err", "exp", "fix", "format", "freefile", "hex", "hour", "iif",
    "instr", "instrrev", "int", "isarray", "isdate", "isempty", "iserror",
    "ismissing", "isnull", "isnumeric", "isobject", "join", "lbound",
    "lcase", "left", "len", "log", "ltrim", "mid", "minute", "month",
    "msgbox", "now", "oct", "replace", "right", "rnd", "round", "rtrim",
    "second", "sgn", "sin", "space", "split", "sqr", "str", "strcomp",
    "string", "strreverse", "switch", "tab", "tan", "time", "timer",
    "timeserial", "timevalue", "trim", "typename", "ubound", "ucase",
    "val", "vartype", "weekday", "year",
    # Statements
    "debug", "print", "open", "close", "get", "put", "input", "write",
    "redim", "erase", "set", "let", "call", "exit", "on", "resume",
    "goto", "gosub", "return", "stop", "end",
    # Keywords
    "if", "then", "else", "elseif", "select", "case", "for", "each",
    "next", "do", "loop", "while", "wend", "with", "and", "or", "not",
    "mod", "is", "like", "nothing", "true", "false", "new", "me",
    "dim", "as", "byval", "byref", "optional", "paramarray",
}


@dataclass
class ProcedureInfo:
    """Information about a VBA procedure."""
    name: str
    module: str
    kind: str  # "Sub", "Function", "Property Get", etc.
    access: str  # "Public", "Private", "Friend"
    line_num: int
    calls: Set[str] = field(default_factory=set)


@dataclass
class CallGraph:
    """Directed call graph for a VBA project."""
    procedures: Dict[str, ProcedureInfo] = field(default_factory=dict)  # qualified name -> info
    edges: List[tuple] = field(default_factory=list)  # (caller, callee)
    modules: Set[str] = field(default_factory=set)

    def to_dict(self) -> dict:
        nodes = []
        for qname, proc in sorted(self.procedures.items()):
            nodes.append({
                "name": proc.name,
                "module": proc.module,
                "kind": proc.kind,
                "access": proc.access,
                "line": proc.line_num,
                "calls": sorted(proc.calls),
            })
        return {
            "modules": sorted(self.modules),
            "procedures": nodes,
            "edges": [{"from": a, "to": b} for a, b in self.edges],
            "node_count": len(self.procedures),
            "edge_count": len(self.edges),
        }


def build_call_graph(
    source_dir: str,
    extensions: tuple = (".bas", ".cls"),
) -> CallGraph:
    """
    Build a call dependency graph from VBA source files.

    Parameters
    ----------
    source_dir : str
        Path to the vba_source/ directory.

    Returns
    -------
    CallGraph
        The constructed call graph.
    """
    src_dir = os.path.abspath(source_dir)
    graph = CallGraph()

    # Pass 1: Collect all procedure definitions
    all_proc_names = set()  # lowercase names for matching

    for root, _, files in os.walk(src_dir):
        for fname in sorted(files):
            if not fname.endswith(extensions):
                continue
            filepath = os.path.join(root, fname)
            module = os.path.splitext(fname)[0]
            graph.modules.add(module)

            lines = _read_lines(filepath)
            for i, line in enumerate(lines, start=1):
                stripped = line.strip()
                m = _PROC_DEF_RE.match(stripped)
                if m:
                    access = (m.group(1) or "Public").strip()
                    kind = m.group(2).strip()
                    name = m.group(3)
                    qname = f"{module}.{name}"

                    graph.procedures[qname] = ProcedureInfo(
                        name=name,
                        module=module,
                        kind=kind,
                        access=access,
                        line_num=i,
                    )
                    all_proc_names.add(name.lower())

    # Pass 2: Find call sites
    for root, _, files in os.walk(src_dir):
        for fname in sorted(files):
            if not fname.endswith(extensions):
                continue
            filepath = os.path.join(root, fname)
            module = os.path.splitext(fname)[0]

            lines = _read_lines(filepath)
            current_proc = None

            for line in lines:
                stripped = line.strip()

                # Track current procedure
                m = _PROC_DEF_RE.match(stripped)
                if m:
                    current_proc = f"{module}.{m.group(3)}"
                    continue

                if re.match(r"^End\s+(Sub|Function|Property)", stripped, re.IGNORECASE):
                    current_proc = None
                    continue

                if current_proc is None:
                    continue

                # Skip comments
                if stripped.startswith("'") or stripped.lower().startswith("rem "):
                    continue

                # Find potential call sites
                # Remove string literals to avoid false positives
                clean = _remove_strings(stripped)

                # Remove inline comments starting with ' to prevent matching calls in comments
                if "'" in clean:
                    clean = clean.split("'", 1)[0]

                for match in _CALL_RE.finditer(clean):
                    callee = match.group(1)
                    if callee.lower() in _BUILTINS:
                        continue
                    if callee.lower() not in all_proc_names:
                        continue

                    # Find the callee's qualified name
                    callee_qname = _resolve_callee(callee, graph.procedures)
                    if callee_qname and callee_qname != current_proc:
                        if current_proc in graph.procedures:
                            graph.procedures[current_proc].calls.add(callee_qname)
                            graph.edges.append((current_proc, callee_qname))

    # Deduplicate edges
    graph.edges = sorted(set(graph.edges))

    logger.info(
        f"Call graph: {len(graph.procedures)} procedures, "
        f"{len(graph.edges)} edges, {len(graph.modules)} modules"
    )
    return graph


def render_mermaid(graph: CallGraph, direction: str = "TD") -> str:
    """Render the call graph as a Mermaid flowchart."""
    lines = [f"graph {direction}"]

    # Group nodes by module
    for module in sorted(graph.modules):
        procs = [p for p in graph.procedures.values() if p.module == module]
        if not procs:
            continue
        lines.append(f"    subgraph {module}")
        for proc in sorted(procs, key=lambda p: p.name):
            qname = f"{module}.{proc.name}"
            node_id = qname.replace(".", "_")
            label = f"{proc.kind} {proc.name}"
            lines.append(f'        {node_id}["{label}"]')
        lines.append("    end")

    # Edges
    for caller, callee in graph.edges:
        caller_id = caller.replace(".", "_")
        callee_id = callee.replace(".", "_")
        lines.append(f"    {caller_id} --> {callee_id}")

    return "\n".join(lines)


def render_dot(graph: CallGraph) -> str:
    """Render the call graph as a Graphviz DOT file."""
    lines = ["digraph call_graph {", "    rankdir=LR;", "    node [shape=box];"]

    for module in sorted(graph.modules):
        procs = [p for p in graph.procedures.values() if p.module == module]
        if not procs:
            continue
        lines.append(f"    subgraph cluster_{module} {{")
        lines.append(f'        label="{module}";')
        for proc in sorted(procs, key=lambda p: p.name):
            qname = f"{module}.{proc.name}"
            node_id = qname.replace(".", "_")
            lines.append(f'        {node_id} [label="{proc.name}"];')
        lines.append("    }")

    for caller, callee in graph.edges:
        caller_id = caller.replace(".", "_")
        callee_id = callee.replace(".", "_")
        lines.append(f"    {caller_id} -> {callee_id};")

    lines.append("}")
    return "\n".join(lines)


# ── Helpers ──

def _read_lines(filepath: str) -> list[str]:
    from xlvbatools.vba._io import read_vba_lines
    return read_vba_lines(filepath)


def _remove_strings(line: str) -> str:
    """Remove string literals from a VBA line."""
    return re.sub(r'"[^"]*"', '""', line)


def _resolve_callee(name: str, procedures: dict) -> str | None:
    """Resolve a callee name to its qualified name."""
    # Exact match (module.name already)
    if name in procedures:
        return name
    # Search by unqualified name
    matches = [qn for qn, p in procedures.items() if p.name.lower() == name.lower()]
    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        # Prefer public procedures
        public = [qn for qn in matches if procedures[qn].access.lower() == "public"]
        if len(public) == 1:
            return public[0]
        return matches[0]  # Ambiguous, pick first
    return None
