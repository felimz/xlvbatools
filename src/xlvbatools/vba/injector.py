"""
VBA Injector
=============
Injects VBA source code from plain-text files on disk back into an Excel
workbook (.xlsm). Creates automatic backups before injection.

Standard and class modules are injected by removing the old component
and importing the file via VBE.
Document modules (sheet code-behinds) are injected by clearing and
replacing the CodeModule content line by line.

Usage:
    from xlvbatools.vba import inject_all, inject_component

    results = inject_all("workbook.xlsm", "vba_source/")
    success = inject_component("workbook.xlsm", "vba_source/", "modMain")
"""

import logging
import os
import shutil
import tempfile

from xlvbatools.core.session import ExcelSession
from xlvbatools.vba.manifest import (
    Manifest, get_type_info,
    TYPE_STD_MODULE as _TYPE_STD_MODULE,
    TYPE_CLASS_MODULE as _TYPE_CLASS_MODULE,
    TYPE_DOCUMENT as _TYPE_DOCUMENT,
)

logger = logging.getLogger(__name__)

# Default backup limit
_DEFAULT_BACKUP_LIMIT = 5


def inject_all(
    workbook_path: str,
    source_dir: str,
    backup: bool = True,
    dry_run: bool = False,
    backup_limit: int = _DEFAULT_BACKUP_LIMIT,
) -> list[dict]:
    """
    Inject all VBA components from source files into the workbook.

    Parameters
    ----------
    workbook_path : str
        Path to the .xlsm workbook.
    source_dir : str
        Path to the vba_source/ directory.
    backup : bool
        Create a backup before injection (default True).
    dry_run : bool
        If True, report what would be injected without making changes.
    backup_limit : int
        Max number of backup files to keep.

    Returns
    -------
    list of dict
        Injection results for each component.
    """
    wb_path = os.path.abspath(workbook_path)
    src_dir = os.path.abspath(source_dir)

    # Load manifest to know what to inject
    manifest_path = os.path.join(src_dir, "manifest.json")
    manifest = Manifest.load(manifest_path)

    if not manifest.components:
        # Fall back to scanning the directory structure
        manifest = _scan_source_dir(src_dir)

    if dry_run:
        return [
            {"name": c.name, "file": c.file, "action": "would inject", "status": "dry-run"}
            for c in manifest.components
        ]

    if backup:
        _create_backup(wb_path, backup_limit)

    results = []

    with ExcelSession(wb_path, visible=False, save_on_exit=True) as session:
        vb_project = session.wb.VBProject

        for comp_info in manifest.components:
            filepath = os.path.join(src_dir, comp_info.file)
            if not os.path.exists(filepath):
                results.append({
                    "name": comp_info.name,
                    "file": comp_info.file,
                    "status": "skipped",
                    "reason": "file not found",
                })
                continue

            try:
                _inject_single(vb_project, comp_info.name, filepath, comp_info.type_code)
                results.append({
                    "name": comp_info.name,
                    "file": comp_info.file,
                    "status": "injected",
                })
                logger.info(f"Injected: {comp_info.name}")
            except Exception as e:
                results.append({
                    "name": comp_info.name,
                    "file": comp_info.file,
                    "status": "error",
                    "error": str(e),
                })
                logger.error(f"Failed to inject {comp_info.name}: {e}")

    return results


def inject_component(
    workbook_path: str,
    source_dir: str,
    component_name: str,
    backup: bool = True,
) -> bool:
    """
    Inject a single VBA component by name.

    Returns True on success, False on failure.
    """
    wb_path = os.path.abspath(workbook_path)
    src_dir = os.path.abspath(source_dir)

    # Find the component file
    filepath, type_code = _find_component_file(src_dir, component_name)
    if filepath is None:
        logger.error(f"Component file not found for: {component_name}")
        return False

    if backup:
        _create_backup(wb_path)

    try:
        with ExcelSession(wb_path, visible=False, save_on_exit=True) as session:
            _inject_single(session.wb.VBProject, component_name, filepath, type_code)
            logger.info(f"Injected: {component_name}")
            return True
    except Exception as e:
        logger.error(f"Failed to inject {component_name}: {e}")
        return False


def _inject_single(vb_project, name: str, filepath: str, type_code: int):
    """Inject a single component into the VBE project."""
    if type_code == _TYPE_DOCUMENT:
        _inject_document_module(vb_project, name, filepath)
    else:
        _inject_importable_module(vb_project, name, filepath, type_code)


def _inject_importable_module(vb_project, name: str, filepath: str, type_code: int):
    """
    Inject a standard or class module by removing the old one and importing the file.

    VBE Import requires a temp file in windows-1252 encoding because VBE
    only reads ANSI files.
    """
    # Remove existing component if it exists
    try:
        existing = vb_project.VBComponents(name)
        vb_project.VBComponents.Remove(existing)
        logger.debug(f"Removed existing component: {name}")
    except Exception:
        pass  # Component doesn't exist yet, that's fine

    # Create temp file in ANSI encoding for VBE Import
    temp_dir = tempfile.mkdtemp()
    ext = get_type_info(type_code)["ext"]
    temp_path = os.path.join(temp_dir, f"{name}{ext}")

    try:
        # Read UTF-8 source, write as windows-1252 with CRLF
        with open(filepath, "r", encoding="utf-8") as f:
            content = f.read()
        with open(temp_path, "w", encoding="windows-1252", newline="\r\n",
                  errors="replace") as f:
            f.write(content)

        # Import via VBE
        vb_project.VBComponents.Import(temp_path)
        logger.debug(f"Imported: {name} from {filepath}")
    finally:
        try:
            os.remove(temp_path)
            os.rmdir(temp_dir)
        except Exception:
            pass


def _inject_document_module(vb_project, name: str, filepath: str):
    """
    Inject a document module (sheet code-behind) by clearing and replacing
    the CodeModule content. Document modules can't be removed/re-imported.
    """
    # Find the existing document module
    comp = None
    for c in vb_project.VBComponents:
        if c.Name == name:
            comp = c
            break

    if comp is None:
        raise ValueError(f"Document module not found in workbook: {name}")

    # Read the new source
    with open(filepath, "r", encoding="utf-8") as f:
        new_code = f.read()

    # Clear existing code
    cm = comp.CodeModule
    if cm.CountOfLines > 0:
        cm.DeleteLines(1, cm.CountOfLines)

    # Insert new code line by line
    lines = new_code.split("\n")
    for line in lines:
        line = line.rstrip("\r")
        cm.InsertLines(cm.CountOfLines + 1, line)

    logger.debug(f"Replaced document module: {name} ({len(lines)} lines)")


def _find_component_file(source_dir: str, name: str) -> tuple:
    """Find a component file by name in the source directory."""
    # Check manifest first
    manifest_path = os.path.join(source_dir, "manifest.json")
    manifest = Manifest.load(manifest_path)
    comp = manifest.find(name)
    if comp:
        filepath = os.path.join(source_dir, comp.file)
        if os.path.exists(filepath):
            return filepath, comp.type_code

    # Fall back to scanning directories
    search_dirs = {
        "modules": (_TYPE_STD_MODULE, ".bas"),
        "classes": (_TYPE_CLASS_MODULE, ".cls"),
        "sheets": (_TYPE_DOCUMENT, ".bas"),
    }
    for dirname, (type_code, ext) in search_dirs.items():
        candidate = os.path.join(source_dir, dirname, f"{name}{ext}")
        if os.path.exists(candidate):
            return candidate, type_code

    return None, None


def _scan_source_dir(source_dir: str) -> Manifest:
    """Build a manifest by scanning the source directory structure."""
    from xlvbatools.vba.manifest import ComponentInfo, compute_file_hash

    components = []
    scan_map = {
        "modules": (1, ".bas", "standard_module"),
        "classes": (2, ".cls", "class_module"),
        "sheets": (100, ".bas", "document_module"),
    }

    for dirname, (type_code, ext, type_name) in scan_map.items():
        dir_path = os.path.join(source_dir, dirname)
        if not os.path.isdir(dir_path):
            continue
        for fname in sorted(os.listdir(dir_path)):
            if not fname.endswith(ext):
                continue
            name = os.path.splitext(fname)[0]
            filepath = os.path.join(dir_path, fname)
            with open(filepath, encoding="utf-8") as fh:
                line_count = sum(1 for _ in fh)
            components.append(ComponentInfo(
                name=name,
                type_code=type_code,
                type_name=type_name,
                file=os.path.join(dirname, fname),
                line_count=line_count,
                sha256=compute_file_hash(filepath),
            ))

    return Manifest(components=components)


def _create_backup(workbook_path: str, limit: int = _DEFAULT_BACKUP_LIMIT):
    """Create a timestamped backup of the workbook and prune old backups."""
    import datetime

    backup_dir = os.path.join(os.path.dirname(workbook_path), "backups")
    os.makedirs(backup_dir, exist_ok=True)

    basename = os.path.splitext(os.path.basename(workbook_path))[0]
    ts = datetime.datetime.now().strftime("%Y%m%dT%H%M%S")
    backup_name = f"{basename}_backup_{ts}.xlsm"
    backup_path = os.path.join(backup_dir, backup_name)

    shutil.copy2(workbook_path, backup_path)
    logger.info(f"Backup created: {backup_path}")

    # Prune old backups
    backups = sorted(
        [f for f in os.listdir(backup_dir) if f.startswith(basename) and f.endswith(".xlsm")]
    )
    while len(backups) > limit:
        oldest = backups.pop(0)
        try:
            os.remove(os.path.join(backup_dir, oldest))
            logger.info(f"Pruned old backup: {oldest}")
        except Exception:
            pass
