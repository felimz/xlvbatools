"""
VBA Source Search
==================
Full-text search across VBA source files (.bas, .cls) in a vba_source/ directory.
Works without COM -- operates directly on the extracted plain-text files.

Usage:
    from xlvbatools.vba.search import search_vba

    results = search_vba("vba_source/", "MsgBox")
    results = search_vba("vba_source/", r"Dim\\s+\\w+$", regex=True)
"""

import logging
import os
import re
from dataclasses import dataclass
from typing import List

logger = logging.getLogger(__name__)


@dataclass
class SearchMatch:
    """A single search match in a VBA source file."""
    file: str       # Relative path within source_dir
    line_num: int   # 1-indexed line number
    line: str       # Full line content (stripped)
    module: str     # Module name (derived from filename)
    context_before: List[str] | None = None  # Lines before the match
    context_after: List[str] | None = None   # Lines after the match

    def __str__(self) -> str:
        return f"{self.file}:{self.line_num}: {self.line}"


def search_vba(
    source_dir: str,
    pattern: str,
    regex: bool = False,
    case_sensitive: bool = False,
    context_lines: int = 0,
    extensions: tuple = (".bas", ".cls"),
) -> List[SearchMatch]:
    """
    Search VBA source files for a pattern.

    Parameters
    ----------
    source_dir : str
        Path to the vba_source/ directory.
    pattern : str
        Search string (literal or regex).
    regex : bool
        If True, treat pattern as a regular expression.
    case_sensitive : bool
        If True, perform case-sensitive matching (default False).
    context_lines : int
        Number of context lines to include around each match (like grep -C).
    extensions : tuple
        File extensions to search.

    Returns
    -------
    list of SearchMatch
        All matches found, sorted by file then line number.
    """
    src_dir = os.path.abspath(source_dir)
    if not os.path.isdir(src_dir):
        logger.error(f"Source directory not found: {src_dir}")
        return []

    # Compile pattern
    flags = 0 if case_sensitive else re.IGNORECASE
    if regex:
        try:
            compiled = re.compile(pattern, flags)
        except re.error as e:
            logger.error(f"Invalid regex pattern: {e}")
            return []
    else:
        # Escape literal pattern for use with re.search
        compiled = re.compile(re.escape(pattern), flags)

    matches = []

    for root, _, files in os.walk(src_dir):
        for fname in sorted(files):
            if not fname.endswith(extensions):
                continue

            filepath = os.path.join(root, fname)
            rel_path = os.path.relpath(filepath, src_dir)
            module_name = os.path.splitext(fname)[0]

            from xlvbatools.vba._io import read_vba_lines as _read_lines
            lines = _read_lines(filepath)

            for i, line in enumerate(lines, start=1):
                if compiled.search(line):
                    # Gather context lines if requested
                    ctx_before = None
                    ctx_after = None
                    if context_lines > 0:
                        start_idx = max(0, i - 1 - context_lines)
                        end_idx = min(len(lines), i + context_lines)
                        ctx_before = [l.rstrip() for l in lines[start_idx:i - 1]]
                        ctx_after = [l.rstrip() for l in lines[i:end_idx]]

                    matches.append(SearchMatch(
                        file=rel_path,
                        line_num=i,
                        line=line.rstrip(),
                        module=module_name,
                        context_before=ctx_before,
                        context_after=ctx_after,
                    ))

    return matches


def search_vba_summary(
    source_dir: str,
    pattern: str,
    **kwargs,
) -> dict:
    """
    Search and return a summary with match counts per file.

    Returns dict with keys: pattern, total_matches, files, matches.
    """
    matches = search_vba(source_dir, pattern, **kwargs)

    file_counts = {}
    for m in matches:
        file_counts[m.file] = file_counts.get(m.file, 0) + 1

    return {
        "pattern": pattern,
        "total_matches": len(matches),
        "files": file_counts,
        "matches": [
            {"file": m.file, "line": m.line_num, "text": m.line}
            for m in matches
        ],
    }
