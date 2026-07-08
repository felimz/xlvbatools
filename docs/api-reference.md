# xlvbatools — Python API Reference

This document provides a detailed reference of the core Python API of `xlvbatools` for programmatic automation, testing, and static analysis of Excel workbooks and VBA source code.

---

## Core COM Session

### `xlvbatools.core.session.ExcelSession`
The central context manager for Excel COM automation. Automatically starts a background `DialogWatchdog` thread, opens the workbook, and ensures process cleanup upon exit.

```python
from xlvbatools.core.session import ExcelSession

with ExcelSession(
    workbook_path="workbook.xlsm",
    visible=False,
    save_on_exit=True,
    kill_on_enter=True,
    init_delay=1.5,
    enable_watchdog=True,
    watchdog_poll_interval=0.25
) as session:
    excel = session.excel
    wb = session.wb
    # Do COM operations...
```

#### Parameters
* **`workbook_path` (str):** Relative or absolute path to the `.xlsm` workbook.
* **`visible` (bool, default `False`):** Shows Excel visibly if `True`.
* **`save_on_exit` (bool, default `True`):** Automatically saves the workbook upon exiting the context.
* **`kill_on_enter` (bool, default `True`):** Kills stale Excel instances before opening.
* **`init_delay` (float, default `1.5`):** delay in seconds to wait for VBA project initialization.
* **`enable_watchdog` (bool, default `True`):** Starts dialog dismissal daemon if `True`.
* **`watchdog_poll_interval` (float, default `0.25`):** Watchdog polling interval in seconds.

#### Methods
* **`run_macro(macro_name: str, timeout: float = 120.0) -> dict`:**
  Runs a VBA macro safely. Under COM error or dialog pop-ups, returns diagnostic information.
  * **Returns:** Dict containing keys `success` (bool), `elapsed_seconds` (float), `error` (str | None), and `dialog_events` (list).
* **`compile_test() -> dict`:**
  Triggers a compile test in VBE.
  * **Returns:** Dict containing keys `success` (bool), `error` (str | None), `error_context` (list[str]).

---

## VBA Operations

### `xlvbatools.vba.extractor.extract_all(workbook_path: str, output_dir: str) -> dict`
Extracts all VBA source components from a workbook onto disk.

### `xlvbatools.vba.injector.inject_all(workbook_path: str, source_dir: str, backup: bool = True, dry_run: bool = False, backup_limit: int = 5) -> list[dict]`
Injects all source modules back into the workbook. Creates backup versions of the workbook in the backups folder unless `backup=False`.

### `xlvbatools.vba.differ.diff_all(workbook_path: str, source_dir: str) -> list[dict]`
Compares current workbook components against source files on disk.

---

## Static Analysis (Linter)

### `xlvbatools.analysis.preflight.lint_files(source_dir: str, disabled_rules: list[str] | None = None) -> list[VBAIssue]`
Lints VBA source files on disk offline (requires no Excel/COM).

### `xlvbatools.analysis.preflight.lint_workbook(workbook_path: str, disabled_rules: list[str] | None = None) -> list[VBAIssue]`
Runs VBE compile checks in addition to file static analysis rules.

---

## Workbook Inspection & Modification

### `xlvbatools.workbook.dumper.dump_sheet_data(workbook_path: str, sheets: list[str], output_json: str | None = None, output_md: str | None = None, custom_range: str | None = None, dump_names: bool = True, max_md_rows: int = 500) -> dict`
Dumps cell values, formulas, formatted texts, interactive shapes, and named ranges to JSON and/or Markdown.

### `xlvbatools.workbook.dumper.export_screenshots(workbook_path: str, sheets: list[str], output_dir: str, custom_range: str | None = None) -> dict[str, str]`
Exports PNG screenshots of the specified worksheets with Pillow-composited column and row headers.

---

## Snapshot Management

### `xlvbatools.snapshot.manager.SnapshotManager`
A dual-layer timestamped checkpoint and rollback system for Excel workbook binaries and VBA text source modules.

```python
from xlvbatools.snapshot import SnapshotManager

mgr = SnapshotManager(
    workbook_path="workbook.xlsm",
    vba_source_dir="workbook/vba_source/",
    snapshots_dir="workbook/snapshots/",
    rolling_limit=10
)

# Create a snapshot
sid = mgr.create(description="before risky feature", milestone=False)

# List all snapshots
log = mgr.list()

# Restore a snapshot
mgr.restore("latest")
```

#### Parameters
* **`workbook_path` (str):** Path to the target `.xlsm` workbook.
* **`vba_source_dir` (str):** Path to the directory containing git-tracked VBA modules.
* **`snapshots_dir` (str):** Directory where snapshot backups and logs will be stored.
* **`rolling_limit` (int, default `10`):** Maximum number of rolling (non-milestone) snapshots to keep before auto-pruning.

#### Methods
* **`create(description: str = "", milestone: bool = False) -> str`:**
  Creates a new snapshot of the current workbook and VBA source files. Auto-prunes older rolling snapshots if the limit is exceeded. Returns the generated snapshot ID.
* **`list() -> list[dict]`:**
  Lists all created snapshots sorted oldest to newest.
* **`info(identifier: str) -> dict | None`:**
  Retrieves detailed metadata for a specific snapshot by ID, index (e.g., `-1`), or description search.
* **`restore(identifier: str, safety_snapshot: bool = True) -> bool`:**
  Restores both the workbook binary and VBA source modules to the snapshot's state. Creates a safety snapshot first unless `safety_snapshot=False`.
* **`diff(identifier: str) -> str`:**
  Generates a git-style line-by-line diff of the VBA changes between the snapshot and current state.
* **`prune(keep: int = 10) -> int`:**
  Manually prunes old rolling snapshots, keeping the specified number of most recent entries. Returns the number of snapshots pruned.

