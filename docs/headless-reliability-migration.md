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

The default exit grace period is 20 seconds. Excel can need more than ten seconds to release VBE and COM state in a loaded desktop test sequence. Session-owned COM state is released before waiting for the PID; force-terminating Excel while teardown is still unwinding can crash or terminate a later session in the same interpreter.

Do not use image-wide termination such as `taskkill /f /im EXCEL.EXE`. It can destroy unrelated user workbooks.

### Release child COM references inside the session

Python variables that reference worksheets, ranges, names, VBE components, or other child COM objects must be cleared before leaving the `ExcelSession` context. `ExcelSession` can release its own workbook and application references, but it cannot delete variables owned by caller or pytest fixture frames.

```python
import gc

with ExcelSession("workbook.xlsm", save_on_exit=False) as session:
    sheet = session.wb.Worksheets("Input")
    cell = sheet.Range("C33")
    value = cell.Value

    cell = None
    sheet = None
    gc.collect()  # Runs while the Excel RPC server is still alive.
```

This is especially important in yield-based pytest fixtures: generator locals survive until fixture teardown, which is after the context manager requests Excel shutdown. Retained proxies can therefore emit pywin32 `0x800706ba` or `0x80010108` finalizer diagnostics even when every assertion passes and the owned process exits.

### Balance COM apartment ownership

`ExcelSession` initializes COM only when the calling thread does not already have an apartment. It uninitializes only an apartment that the session created, and only on the same thread. This preserves caller-owned pywin32 state while giving worker threads a balanced COM lifecycle.

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
