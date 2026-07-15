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

All cross-process queries and button clicks use `SendMessageTimeoutW` with
`SMTO_ABORTIFHUNG`. Preferred buttons depend on the configured strategy; if
no safe button is found, the watchdog uses a bounded window-close request.
The watchdog never interacts with a window owned by another process.

## Lifecycle

1. The isolated worker initializes COM.
2. It creates Excel and resolves `Application.Hwnd` to the owned PID.
3. It starts the PID-scoped watchdog before workbook open.
4. Each macro or compile operation records the starting event sequence and
   consumes only newer events.
5. Worker teardown stops and joins the watchdog before releasing Excel.
6. Dialog events and cleanup fields are converted into `OperationResult`
   diagnostics.

## Parent deadline

A Python thread cannot safely cancel `Excel.Application.Run`. The parent
therefore owns the deadline for the whole worker operation. If the result does
not arrive, it targets the reported Excel PID, allows the blocked COM call to
unwind, and then terminates the exact worker if required. There is no
image-wide Excel fallback.

See [Headless reliability contract](headless-reliability.md) and
[Worker protocol](worker-protocol.md).
