# xlvbatools — Dialog Watchdog Architecture

The `DialogWatchdog` is a background daemon thread that polls for Excel and VBE modal dialog boxes, captures their content, and automatically dismisses them. It prevents headless automation runs from hanging indefinitely.

---

## The Headless Automation Challenge
Headless Excel COM automation often hangs on modal dialogs triggered by:
* **VBA Runtime Errors** (e.g. `Run-time error '1004'`)
* **VBA Compile Errors** (e.g. `Expected: End of Statement`)
* **VBA MsgBox calls**
* **Save/Overwrite Alerts**

Because Excel executes these dialogs synchronously, the COM execution thread blocks. In a headless environment, there is no user to click "OK", causing a permanent lockup. The watchdog solves this by running asynchronously.

---

## Detection Mechanism
The watchdog runs in a background Python thread (`threading.Thread`) and loops continuously at a configurable poll interval (default 250ms).

1. **Window Enumeration:** It invokes Win32 `EnumWindows` to scan all top-level window handles on the desktop.
2. **Class Filtering:** It checks for window handles of class `#32770` (the standard Win32 dialog window class).
3. **PID Filtering:** It maps the window process ID using `GetWindowThreadProcessId` and only targets dialogs owned by the specific Excel process ID of the active `ExcelSession` (preventing accidental dismissal of dialogs from other processes).
4. **Title Keywords:** Filters out non-Excel dialogs by checking if the title contains terms like "Microsoft Excel", "Visual Basic", "Warning", etc. (or is completely untitled).

---

## Dismissal Strategy

Once a dialog is found, the watchdog makes up to three bounded capture attempts before dismissal. Each child is queried first with `WM_GETTEXTLENGTH`/`WM_GETTEXT` through `SendMessageTimeoutW`, then with `GetWindowTextW` when that path is empty. Multiline text is retained, duplicates are removed, and child-control diagnostics are included in the serialized event.

### 1. PID Scope
An auto-dismissing watchdog requires `target_pid`. Capture-only diagnostic scans may omit it by explicitly setting `auto_dismiss=False`. `ExcelSession` creates Excel first, resolves `Application.Hwnd` to its PID, and only then starts the watchdog, so workbook-open dialogs are protected without exposing unrelated processes.

### 2. Message Timeout Protection
To prevent the watchdog itself from freezing when Excel crashes or enters an un-interruptible deadlock, the watchdog uses `SendMessageTimeoutW` (with `SMTO_ABORTIFHUNG` and a tight `250ms` timeout) instead of standard `SendMessageW` to query control labels:

```python
# Timeout query for control text
res = user32.SendMessageTimeoutW(
    hwnd,
    WM_GETTEXTLENGTH,
    0,
    0,
    0x0002,  # SMTO_ABORTIFHUNG
    250,     # 250ms timeout limit
    ctypes.byref(result_len)
)
```

### 3. Click Prioritization
Buttons are selected based on the configured dismiss strategy:
* **OK Strategy (default):** Clicks "OK", "yes", "close", "end", or similar affirmative buttons.
* **Cancel Strategy:** Clicks "cancel", "no", "abort", or similar negative buttons.

If a button is selected, the watchdog sends a `BM_CLICK` message to the button handle (protected by `SendMessageTimeoutW`). If no matching button is found, it sends a fallback `WM_CLOSE` message to the dialog window itself.

---

## Watchdog Lifecycle & Timeouts
The `DialogWatchdog` is managed automatically by `ExcelSession` with the following safety defaults:
1. **Thread Startup**: Initiated asynchronously after Excel PID discovery and before workbook configuration/open. The timeout is set to **600 seconds (10 minutes)** in the session context manager.
2. **Execution Monitoring**: The watchdog logs its initial startup at `DEBUG` level and runs silently in the background unless a dialog is intercepted.
3. **Shutdown Cleanliness**: Upon exiting the `ExcelSession` context, the watchdog thread is signaled via an internal event, joined with a 2.0-second timeout, and safely shut down, returning all captured dialog events to the session log.

Each event has a monotonically increasing `sequence`, allowing individual macro and compile runs to select only newly captured events. Session exit requests `Application.Quit`, releases COM references, waits for `exit_grace_period`, and if configured terminates only the session-owned PID. The outcome is available as `session.cleanup_result`.

## Enforced Macro Timeouts

The high-level `xlvbatools.macro.run_macro()` API uses a spawned worker process because a Python thread cannot safely interrupt `Excel.Application.Run` or satisfy COM apartment isolation.

1. The worker initializes COM and creates the complete `ExcelSession`.
2. As soon as `Application.Hwnd` is resolved, it sends the Excel PID to the parent.
3. The worker applies inputs, runs the macro, performs cleanup, and sends the structured result.
4. The parent enforces the requested deadline. On timeout it terminates the reported Excel PID, waits briefly for the COM call to unblock, and terminates the worker if it remains blocked.

The timeout applies to the complete isolated session, including Excel startup, workbook open, macro execution, save, and cleanup. No image-wide `taskkill /im EXCEL.EXE` fallback is used.
