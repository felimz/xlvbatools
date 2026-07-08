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

Once a dialog is captured, the watchdog queries all child control windows via `EnumChildWindows` to find visible buttons.

### 1. Reentrant Lock Protection
The watchdog utilizes a process-safe reentrant file-locking context manager to modify its state logs, preventing concurrent write issues during multi-process execution.

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
