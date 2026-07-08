"""
VBA File I/O Utilities
========================
Shared helpers for reading VBA source files with encoding fallback.

VBA files exported by VBE use the system ANSI code page (typically
windows-1252 on Western Windows installations). Files written by
xlvbatools are always UTF-8. This module provides a single implementation
of the encoding fallback logic so every call site behaves consistently.
"""

import os
from typing import List


def read_vba_text(filepath: str) -> str:
    """
    Read an entire VBA source file as a string.

    Tries UTF-8 first, falls back to windows-1252 (common ANSI code page
    for VBE-exported files).
    """
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            return f.read()
    except UnicodeDecodeError:
        with open(filepath, "r", encoding="windows-1252") as f:
            return f.read()


def read_vba_lines(filepath: str) -> List[str]:
    """
    Read a VBA source file and return its lines (with line endings).

    Tries UTF-8 first, falls back to windows-1252.
    """
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            return f.readlines()
    except UnicodeDecodeError:
        with open(filepath, "r", encoding="windows-1252") as f:
            return f.readlines()
