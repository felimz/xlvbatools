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
