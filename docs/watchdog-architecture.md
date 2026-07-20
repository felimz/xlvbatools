# Dialog watchdog architecture

The dialog watchdog is an internal worker component. Applications use
`Project.run`, `Project.inspect`, and the other `Project` methods; they do
not create watchdog or COM-session objects.

## Why it exists

Excel and the VBE can synchronously display runtime errors, compile errors,
`MsgBox` prompts, file pickers, and save alerts. A blocked COM call cannot
dismiss its own modal window, so the worker runs a separate watchdog thread.

## Detection

The watchdog:

1. enumerates top-level Win32 dialog windows;
2. maps each window to a process with `GetWindowThreadProcessId`;
3. accepts only dialogs owned by the worker's exact Excel PID;
4. reads child text with bounded `WM_GETTEXTLENGTH` and `WM_GETTEXT` calls;
5. falls back to `GetWindowTextW` when needed;
6. preserves multiline text, control metadata, and event sequence.

An auto-dismissing watchdog cannot start without `target_pid`. A diagnostic
capture-only scan may omit the PID only when `auto_dismiss=False`.

## Bounded dismissal

Cross-process text queries and the first button click use
`SendMessageTimeoutW` with `SMTO_ABORTIFHUNG`. If the VBE modal loop is busy,
the watchdog queues the same button click and then uses a bounded window-close
request as its final fallback. A dialog is marked dismissed only after its
window is confirmed hidden or destroyed. Failed attempts remain eligible for
the next bounded poll, and a handle that disappears can be captured again if
VBE reuses it for a later prompt. The watchdog never interacts with a window
owned by another process.

## Lifecycle

1. The isolated worker initializes COM.
2. It creates Excel and resolves `Application.Hwnd` to the owned PID.
3. It starts the PID-scoped watchdog before workbook open.
4. Each macro or compile operation records the starting event sequence and
   consumes only newer events.
5. Worker teardown closes the target workbook, keeps a disposable sentinel
   workbook alive while releasing the application proxy, releases the sentinel,
   requests `WM_CLOSE` only on the exact PID-verified Excel window, and polls
   that owned PID for graceful exit.
6. The watchdog remains active throughout that exit window because Excel/VBE
   can post a final compiler dialog after shutdown begins. It stops only after
   the graceful-exit decision and any exact-PID fallback complete.
7. Dialog events and cleanup fields are converted into `OperationResult`
   diagnostics.

## Parent deadline

A Python thread cannot safely cancel `Excel.Application.Run`. The parent
therefore owns the deadline for the whole worker operation. If the result does
not arrive, it targets the reported Excel PID, allows the blocked COM call to
unwind, and then terminates the exact worker if required. There is no
image-wide Excel fallback.

See [Headless reliability contract](headless-reliability.md) and
[Worker protocol](worker-protocol.md).
