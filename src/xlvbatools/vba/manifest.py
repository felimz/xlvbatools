"""
VBA Manifest Manager
=====================
Reads and writes the manifest.json file that tracks VBA component metadata
(name, type, file path, line count, hash) for a vba_source/ directory.
"""

import hashlib
import json
import logging
import os
from dataclasses import dataclass, field, asdict
from typing import List, Optional

logger = logging.getLogger(__name__)

# VBA component type constants (from VBE type enum)
VB_COMP_TYPES = {
    1: {"name": "standard_module", "dir": "modules", "ext": ".bas"},
    2: {"name": "class_module", "dir": "classes", "ext": ".cls"},
    3: {"name": "userform", "dir": "forms", "ext": ".frm"},
    11: {"name": "activex_designer", "dir": "designers", "ext": ".dsr"},
    100: {"name": "document_module", "dir": "sheets", "ext": ".bas"},
}


def get_type_info(vb_type: int) -> dict:
    """Get directory and extension info for a VBE component type."""
    return VB_COMP_TYPES.get(vb_type, {"name": "unknown", "dir": "other", "ext": ".txt"})


@dataclass
class ComponentInfo:
    """Metadata for a single VBA component."""
    name: str
    type_code: int
    type_name: str
    file: str
    line_count: int = 0
    sha256: str = ""

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class Manifest:
    """VBA source manifest tracking all components."""
    workbook: str = ""
    extracted_at: str = ""
    components: List[ComponentInfo] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "workbook": self.workbook,
            "extracted_at": self.extracted_at,
            "components": [c.to_dict() for c in self.components],
        }

    def save(self, manifest_path: str):
        """Write the manifest to a JSON file."""
        os.makedirs(os.path.dirname(manifest_path), exist_ok=True)
        with open(manifest_path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, indent=2)
        logger.info(f"Manifest written: {manifest_path} ({len(self.components)} components)")

    @classmethod
    def load(cls, manifest_path: str) -> "Manifest":
        """Load a manifest from a JSON file."""
        if not os.path.exists(manifest_path):
            return cls()
        with open(manifest_path, "r", encoding="utf-8") as f:
            data = json.load(f)
        components = [
            ComponentInfo(**c) for c in data.get("components", [])
        ]
        return cls(
            workbook=data.get("workbook", ""),
            extracted_at=data.get("extracted_at", ""),
            components=components,
        )

    def find(self, name: str) -> Optional[ComponentInfo]:
        """Find a component by name (case-insensitive)."""
        for c in self.components:
            if c.name.lower() == name.lower():
                return c
        return None


def compute_file_hash(filepath: str) -> str:
    """Compute SHA-256 hash of a file (first 16 hex chars)."""
    hasher = hashlib.sha256()
    with open(filepath, "rb") as f:
        while chunk := f.read(8192):
            hasher.update(chunk)
    return hasher.hexdigest()[:16]
