# xlvbatools — Python API Reference

This document provides a detailed reference of the core Python API of `xlvbatools` for programmatic automation, testing, and static analysis of Excel workbooks and VBA source code.

---

## Project Facade and Result Contract

### `xlvbatools.XlvbaProject`

`XlvbaProject` is the recommended entry point for project-specific wrappers.
It binds the nearest `xlvbatools.toml`, resolves paths relative to that file,
and keeps operation results independent of CLI output or internal module paths.

```python
from xlvbatools import XlvbaProject

project = XlvbaProject.from_config()

inspection = project.inspect(
    ["Input"],
    cell_range="B91:K99",
    include_data=True,
    include_screenshots=True,
)
inspection.require_success()
inspection.require_clean_shutdown()

macro = project.run_macro("OnCalculate", timeout=120)
macro.require_success()
macro.require_clean_shutdown()
```

The facade exposes:

- `inspect(...)`
- `run_macro(...)`
- `lint(...)`
- `extract(...)`
- `inject(...)`
- `diff(...)`
- `modify(...)`
- `snapshot_manager()`

All COM-backed facade methods use one shared worker protocol. Each request is
executed by a fresh interpreter and a fresh owned Excel process; no COM proxy
crosses the process boundary. The parent enforces the method's `timeout`, reads
atomic progress containing the exact worker and Excel PIDs, and can stop only
those owned processes. `extract`, `inject`, `diff`, `modify`, and workbook-mode
`lint` accept a `timeout` keyword in addition to their operation-specific
arguments.

`XlvbaProject.for_workbook(...)` provides the same interface without requiring
a TOML file.

### `xlvbatools.OperationResult[T]`

Every facade method returns a schema-versioned `OperationResult` containing:

- `schema_version`
- `operation`, `success`, and `phase`
- typed `data`
- `ErrorInfo`
- produced `Artifact` records
- `Diagnostics`, including dialog events and `CleanupReport`
- operation metadata and warnings

`to_dict()` returns only JSON-compatible containers. `require_success()` raises
`OperationFailedError`; `require_clean_shutdown()` raises
`HeadlessCleanupError` unless Excel reported a graceful, non-forced exit.

Existing function-level APIs remain supported for advanced callers that
deliberately manage an in-process COM lifecycle.

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
    kill_on_enter=False,
    init_delay=1.5,
    enable_watchdog=True,
    watchdog_poll_interval=0.25,
    exit_grace_period=20.0,
    terminate_owned_process=True,
) as session:
    excel = session.excel
    wb = session.wb
    # Do COM operations...
```

#### Parameters
* **`workbook_path` (str):** Relative or absolute path to the `.xlsm` workbook.
* **`visible` (bool, default `False`):** Shows Excel visibly if `True`.
* **`save_on_exit` (bool, default `True`):** Automatically saves the workbook upon exiting the context.
* **`kill_on_enter` (bool, default `False`):** Existing Excel sessions are never touched by default. Enable targeted stale-workbook recovery only as an explicit diagnostic action.
* **`init_delay` (float, default `1.5`):** delay in seconds to wait for VBA project initialization.
* **`enable_watchdog` (bool, default `True`):** Starts dialog dismissal daemon if `True`.
* **`watchdog_poll_interval` (float, default `0.25`):** Watchdog polling interval in seconds.
* **`exit_grace_period` (float, default `20.0`):** Seconds to wait for the owned Excel PID after requesting quit. Excel/VBE shutdown can exceed ten seconds under a loaded desktop test sequence; allowing it to finish avoids destabilizing the next COM session.
* **`terminate_owned_process` (bool, default `True`):** Force-terminates only the spawned PID if it outlives the grace period.

#### Properties
* **`excel`:** Reference to the spawned raw Excel COM Application object.
* **`wb`:** Reference to the opened Workbook object.
* **`excel_pid` (int | None):** The process ID of the spawned Excel process.
* **`cleanup_result` (dict):** Records quit request, graceful exit, targeted termination, and final process state.
* **`vb_project`:** Getter for the workbook's `VBProject` object. Validates "Trust access to the VBA project object model" automatically. If disabled in the Trust Center, raises a clear, diagnostic `RuntimeError` instead of raising a raw COM exception.

#### Methods
* **`run_macro(macro_name: str, timeout: float = 120.0) -> dict`:**
  Runs a VBA macro inside an already-open session. Under COM error or dialog pop-ups, returns diagnostic information. Because the COM call executes in the caller process, this low-level method cannot interrupt an infinite loop; use `xlvbatools.macro.run_macro` for an enforced timeout.
  * **Returns:** Structured result containing `success`, `run_id`, `macro`, `phase`, `elapsed_seconds`, `primary_error`, `com_error`, `dialog_events`, and `cleanup` as applicable.
* **`compile_test() -> dict`:**
  Forces project compilation through a temporary unsaved no-op VBA probe, avoiding visible VBE command-bar menus.
  * **Returns:** Dict containing keys `success` (bool), `error` (str | None), `error_context` (list[str]).
* **`set_named_range(name, value, strict=False) -> bool`:** Sets a workbook name. Strict mode raises immediately so execution cannot continue with stale inputs.

## Shared Isolated Worker

The facade and CLI use the internal `xlvbatools.core.worker` executor for:

- workbook inspection;
- macro execution;
- VBA component listing and extraction;
- VBA injection;
- workbook/source differencing;
- live-workbook lint and compile checks; and
- cell, formula, and named-range modification.

The versioned protocol uses a request file, atomically replaced progress file,
and atomically replaced result file. Worker output is redirected to a regular
file so Excel cannot keep an inherited output pipe open. On Windows, a venv's
base interpreter is launched with the active venv preserved, avoiding a
detached launcher/child pair that would make timeout ownership ambiguous.

See [Worker Protocol](worker-protocol.md) for the schema and lifecycle rules.

### `xlvbatools.macro.run_macro`

Runs the complete Excel session through the shared worker. The worker reports
its isolated Excel PID before opening the workbook. The parent waits up to
`timeout`, terminates only that PID when the deadline expires, and then
terminates the blocked worker if necessary.

```python
from xlvbatools.macro import run_macro

result = run_macro("workbook.xlsm", "OnCalculate", timeout=120)
if result.get("timed_out"):
    print(result["excel_pid"], result["cleanup"])
```

Timeout results include `timed_out`, `timeout_seconds`, `worker_pid`,
`excel_pid`, and targeted `cleanup` details. The shared worker uses a subprocess
entry point and does not require `multiprocessing.freeze_support()`.

Successful result example:

```json
{
  "success": true,
  "run_id": "9e11b72e-9b41-43af-b351-3ab0c2158e83",
  "macro": "OnCalculate",
  "phase": "macro_execution",
  "elapsed_seconds": 1.234,
  "excel_pid": 12345,
  "dialog_events": [],
  "cleanup": {
    "pid": 12345,
    "quit_requested": true,
    "exited_gracefully": true,
    "force_terminated": false,
    "still_running": false
  }
}
```

Timeout result example:

```json
{
  "success": false,
  "run_id": "8b832a6c-2601-4fd3-aa51-805bd95ab36d",
  "macro": "LoopForever",
  "phase": "macro_execution",
  "timed_out": true,
  "timeout_seconds": 120.0,
  "excel_pid": 12345,
  "primary_error": "Execution timed out after 120.000 seconds",
  "dialog_events": [],
  "cleanup": {
    "pid": 12345,
    "quit_requested": false,
    "exited_gracefully": false,
    "force_terminated": true,
    "worker_terminated": false,
    "still_running": false
  }
}
```

See [Headless Reliability Migration](headless-reliability-migration.md) for compatibility guidance.

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

Screenshots use Excel's native `Range.CopyPicture` on the original read-only worksheet. Only the resulting bitmap is pasted into a blank chart workbook; worksheet code is never copied.

Hidden and VeryHidden worksheets are skipped unless `include_hidden_sheets=True` is explicitly supplied.

### `xlvbatools.workbook.inspect_workbook(...) -> dict`

Runs screenshots and data extraction in one macro-disabled, read-only worker session with a hard timeout and PID-scoped cleanup.

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

---

## Logging Configuration

### `xlvbatools.logging.setup_logging(verbose: bool = False, tool_name: str = "xlvbatools", log_dir: str | None = None, log_name: str = "xlvbatools") -> str`
Configures centralized stream and rotating file logging for all library operations.
* **`verbose` (bool):** Sets console handler logging level to `DEBUG` if `True`, otherwise `INFO`. Rotating file handler always captures full `DEBUG` logs.
* **`tool_name` (str):** The logger label for session metadata entries.
* **`log_dir` (str, optional):** The output directory. Defaults to `logs/` in the current working directory.
* **`log_name` (str):** Base filename (e.g. `xlvbatools` becomes `xlvbatools.log`).
* **Returns:** The absolute string path to the created log file.

#### Path Normalization
All log outputs are automatically formatted via a custom `PathNormalizingFormatter` to convert Windows-style backslashes (`\`) into forward slashes (`/`), guaranteeing cross-platform consistency for automated log parsers.
