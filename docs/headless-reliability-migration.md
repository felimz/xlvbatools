# Headless Reliability Migration

## PID-scoped watchdogs

Auto-dismissing `DialogWatchdog` instances must now receive `target_pid`. This prevents diagnostic code from dismissing dialogs owned by unrelated applications or Excel sessions.

```python
watchdog = DialogWatchdog(target_pid=excel_pid)
```

Unscoped desktop scanning is allowed only in explicit capture-only mode:

```python
watchdog = DialogWatchdog(auto_dismiss=False)
```

Code using `ExcelSession` requires no watchdog migration. The session creates Excel, resolves `Application.Hwnd` to its PID, and starts the watchdog before workbook open.

## Process cleanup

`kill_on_enter` controls only targeted stale-workbook handling before startup. Session exit always waits for and, when configured, terminates the exact PID spawned by that session. Inspect `session.cleanup_result` for the outcome.

Do not use image-wide termination such as `taskkill /f /im EXCEL.EXE`. It can destroy unrelated user workbooks.

## Enforced timeouts

`ExcelSession.run_macro()` remains a low-level blocking COM operation. Its `timeout` argument is retained for API compatibility but cannot interrupt `Application.Run` in the caller process.

Use the high-level worker runner when a deadline must be enforced:

```python
from xlvbatools.macro import run_macro

result = run_macro("workbook.xlsm", "OnCalculate", timeout=120)
```

The timeout covers the entire isolated session. The parent terminates only the worker-reported Excel PID and then the worker if COM remains blocked.

## Strict setup

`ExcelSession.set_named_range(name, value, strict=True)` raises when the name cannot be assigned. The high-level macro runner enables strict named-range setup by default so a macro cannot continue with stale workbook values. Pass `strict_named_ranges=False` only for compatibility with workflows that intentionally tolerate missing names.

## Result contract

Callers should use `primary_error` for the best available diagnostic and `phase` for attribution. Preserve `dialog_events`, `com_error`, and `cleanup` in logs. Timeout handling should check `timed_out` rather than matching error strings.
