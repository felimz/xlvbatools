"""
VBA Extractor
==============
Extracts VBA components from an Excel workbook (.xlsm) to plain-text files
on disk, organized by component type (modules/, classes/, sheets/).

Standard and class modules are exported via VBE Export method.
Document modules (sheet code-behinds, ThisWorkbook) are read line-by-line
from the CodeModule since VBE Export is not supported for them.

Usage:
    from xlvbatools.vba import extract_all, extract_component, list_components

    manifest = extract_all("workbook.xlsm", "vba_source/")
    result = extract_component("workbook.xlsm", "modMain", "vba_source/")
    components = list_components("workbook.xlsm")
"""

import datetime
import logging
import os
import tempfile

from xlvbatools.core.session import ExcelSession
from xlvbatools.vba.manifest import (
    ComponentInfo, Manifest, get_type_info, compute_file_hash,
)

logger = logging.getLogger(__name__)

# VBE component type codes
_TYPE_STD_MODULE = 1
_TYPE_CLASS_MODULE = 2
_TYPE_DOCUMENT = 100


def list_components(workbook_path: str) -> list[dict]:
    """
    List all VBA components in a workbook without extracting.

    Returns a list of dicts with keys: name, type_code, type_name, line_count.
    """
    wb_path = os.path.abspath(workbook_path)
    components = []

    with ExcelSession(wb_path, visible=False, save_on_exit=False) as session:
        for comp in session.wb.VBProject.VBComponents:
            type_code = comp.Type
            type_info = get_type_info(type_code)
            line_count = comp.CodeModule.CountOfLines if comp.CodeModule else 0
            components.append({
                "name": comp.Name,
                "type_code": type_code,
                "type_name": type_info["name"],
                "line_count": line_count,
            })

    return sorted(components, key=lambda c: (c["type_code"], c["name"]))


def extract_component(
    workbook_path: str,
    component_name: str,
    output_dir: str,
) -> dict | None:
    """
    Extract a single VBA component by name.

    Returns a dict with component info, or None if not found.
    """
    wb_path = os.path.abspath(workbook_path)
    out_dir = os.path.abspath(output_dir)

    with ExcelSession(wb_path, visible=False, save_on_exit=False) as session:
        vb_project = session.wb.VBProject
        comp = None
        for c in vb_project.VBComponents:
            if c.Name.lower() == component_name.lower():
                comp = c
                break

        if comp is None:
            return None

        info = _extract_single(comp, out_dir)
        return info.to_dict() if info else None


def extract_all(
    workbook_path: str,
    output_dir: str,
) -> dict:
    """
    Extract all VBA components from a workbook to disk.

    Creates the directory structure:
        output_dir/
            modules/     (.bas files)
            classes/     (.cls files)
            sheets/      (.bas files for document modules)
            manifest.json

    Returns the manifest as a dict.
    """
    wb_path = os.path.abspath(workbook_path)
    out_dir = os.path.abspath(output_dir)

    # Create output directories
    for subdir in ("modules", "classes", "sheets"):
        os.makedirs(os.path.join(out_dir, subdir), exist_ok=True)

    components = []

    with ExcelSession(wb_path, visible=False, save_on_exit=False) as session:
        vb_project = session.wb.VBProject

        for comp in vb_project.VBComponents:
            info = _extract_single(comp, out_dir)
            if info:
                components.append(info)

    # Build and save manifest
    manifest = Manifest(
        workbook=os.path.basename(wb_path),
        extracted_at=datetime.datetime.now().isoformat(timespec="seconds"),
        components=sorted(components, key=lambda c: (c.type_code, c.name)),
    )
    manifest_path = os.path.join(out_dir, "manifest.json")
    manifest.save(manifest_path)

    logger.info(f"Extracted {len(components)} components from {os.path.basename(wb_path)}")
    return manifest.to_dict()


def _extract_single(comp, output_dir: str) -> ComponentInfo | None:
    """Extract a single VBE component to the appropriate subdirectory."""
    type_code = comp.Type
    type_info = get_type_info(type_code)
    name = comp.Name

    subdir = os.path.join(output_dir, type_info["dir"])
    os.makedirs(subdir, exist_ok=True)
    filepath = os.path.join(subdir, f"{name}{type_info['ext']}")

    try:
        if type_code in (_TYPE_STD_MODULE, _TYPE_CLASS_MODULE):
            # Standard and class modules: use VBE Export
            _export_via_vbe(comp, filepath)
        elif type_code == _TYPE_DOCUMENT:
            # Document modules: read CodeModule line by line
            _export_code_module(comp, filepath)
        else:
            # Userforms and other types: try Export, fall back to CodeModule
            try:
                _export_via_vbe(comp, filepath)
            except Exception:
                _export_code_module(comp, filepath)

        # Compute metadata
        line_count = _count_lines(filepath)
        sha = compute_file_hash(filepath)

        info = ComponentInfo(
            name=name,
            type_code=type_code,
            type_name=type_info["name"],
            file=os.path.relpath(filepath, output_dir),
            line_count=line_count,
            sha256=sha,
        )
        logger.debug(f"Extracted: {name} -> {info.file} ({line_count} lines)")
        return info

    except Exception as e:
        logger.error(f"Failed to extract {name}: {e}")
        return None


def _export_via_vbe(comp, filepath: str):
    """Export a component using the VBE Export method."""
    # VBE Export writes in the system ANSI code page, so we use a temp file
    # and then normalize to UTF-8
    temp_dir = tempfile.mkdtemp()
    temp_path = os.path.join(temp_dir, os.path.basename(filepath))
    try:
        comp.Export(temp_path)
        # Read with windows-1252 (ANSI), write as UTF-8
        with open(temp_path, "r", encoding="windows-1252") as f:
            content = f.read()
        with open(filepath, "w", encoding="utf-8", newline="\r\n") as f:
            f.write(content)
    finally:
        try:
            os.remove(temp_path)
            os.rmdir(temp_dir)
        except Exception:
            pass


def _export_code_module(comp, filepath: str):
    """Export a document module by reading CodeModule lines."""
    cm = comp.CodeModule
    total_lines = cm.CountOfLines
    if total_lines == 0:
        with open(filepath, "w", encoding="utf-8", newline="\r\n") as f:
            f.write("")
        return

    # Read all lines at once
    code = cm.Lines(1, total_lines)
    with open(filepath, "w", encoding="utf-8", newline="\r\n") as f:
        f.write(code)
        if not code.endswith("\n"):
            f.write("\r\n")


def _count_lines(filepath: str) -> int:
    """Count non-empty lines in a file."""
    try:
        with open(filepath, "r", encoding="utf-8") as f:
            return sum(1 for line in f if line.strip())
    except Exception:
        return 0
