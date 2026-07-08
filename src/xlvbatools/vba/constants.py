"""
VBA and VBE Constants
======================
Shared constants and regex patterns used across the xlvbatools.vba packages.
"""

import re

# VBA component type constants (from VBE type enum)
TYPE_STD_MODULE = 1
TYPE_CLASS_MODULE = 2
TYPE_USERFORM = 3
TYPE_ACTIVEX_DESIGNER = 11
TYPE_DOCUMENT = 100

# Canonical mapping of component type details
VB_COMP_TYPES = {
    TYPE_STD_MODULE: {"name": "standard_module", "dir": "modules", "ext": ".bas"},
    TYPE_CLASS_MODULE: {"name": "class_module", "dir": "classes", "ext": ".cls"},
    TYPE_USERFORM: {"name": "userform", "dir": "forms", "ext": ".frm"},
    TYPE_ACTIVEX_DESIGNER: {"name": "activex_designer", "dir": "designers", "ext": ".dsr"},
    TYPE_DOCUMENT: {"name": "document_module", "dir": "sheets", "ext": ".bas"},
}

# VBE header lines that should be stripped for comparison (ignoring Option Explicit etc.)
VBE_HEADER_STRIP_RE = re.compile(
    r"^(Attribute VB_|VERSION \d|BEGIN|END|  MultiUse =)", re.IGNORECASE
)

# VBE header lines for formatting indentation (includes Option lines because they stay at column 0)
VBE_HEADER_FORMAT_RE = re.compile(
    r"^(Attribute\s+VB_|VERSION\s+\d|BEGIN|END|  MultiUse\s*=|Option\s+)", re.IGNORECASE
)
