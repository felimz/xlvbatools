# Headless reliability contract

These current invariants define xlvbatools v1. Application code receives them
through `Project`; implementation subpackages are not an alternate public API.

## Isolation and ownership

- Every Excel-backed attempt starts one directly tracked worker.
- The worker starts and reports one owned Excel PID.
- No operation attaches to an existing desktop Excel instance.
- No cleanup path terminates Excel by image name.
- On timeout, the parent targets the reported Excel PID before the exact worker.
- Raw COM proxies never cross the process boundary.
- Related steps in one workflow share one owned worker and Excel PID; separate
  workflows never share either.

## COM lifetime

- Worksheets, ranges, names, VBE components, and other child proxies are
  released while their owning session is still alive.
- A worker balances only the COM apartment initialization it performed.
- Startup never reads or modifies Excel's registry configuration. Installed
  COM add-ins load according to the user's existing Excel configuration.
- Teardown closes the target workbook, releases all COM proxies while a blank
  sentinel workbook keeps the owned instance alive, releases the application
  proxy before the sentinel proxy, and only then requests normal `WM_CLOSE`
  shutdown through the exact PID-verified Excel window. A bounded `WM_QUIT`
  request may be sent only to that verified UI thread before exact-PID force
  termination is considered.
- Teardown waits for graceful Excel/VBE exit before targeted termination.
- Liveness polling and the final exact-PID fallback use Win32 process handles;
  cleanup does not spawn `tasklist` or `taskkill` commands on every poll.
- Sequential tests inspect stdout and stderr for native finalizer diagnostics;
  a zero process exit code alone is insufficient.

## Dialog safety

- Auto-dismissal starts only after Excel PID discovery and is PID-scoped.
- Cross-process window text and click operations are bounded by
  `SendMessageTimeoutW`.
- Captured events preserve multiline text and child-control diagnostics.
- A dismissal is successful only after the dialog is confirmed hidden or
  destroyed. Failed clicks are retried on a later bounded poll, including when
  VBE reuses the same window handle for a shutdown-time compiler prompt.
- The PID-scoped watchdog remains active after normal shutdown is requested
  and throughout graceful-exit polling so late VBE prompts cannot strand Excel.
- Live compile resolves the VBE Compile command specifically as an Office
  command-bar button (`Type=1`, `ID=578`). Searching by ID alone is forbidden
  because it can resolve a popup control and expose the File menu.
- The requested workbook's normalized VBProject filename must match
  `ActiveVBProject` before and after Compile; xlvbatools refuses to attribute
  another workbook or add-in's compile state to the target.
- The PID-scoped watchdog also hides the owned VBE editor frame throughout
  every noninteractive operation. Only `xlvba debug` permits it to remain
  visible.
- A disabled Compile button conclusively verifies a valid project. A captured
  compiler failure produces `CT001` with the best available dialog, module,
  line, column, and context evidence. Any other state is an unverifiable
  `CT001`; it never degrades to a warning or passing lint result.

## Deadlines and retries

The parent enforces one deadline around the complete operation, including any
retry. Python threads are not used to interrupt blocking COM calls. The
executor permits at most two total attempts.

One automatic retry is safe only before Excel ownership: child-process
creation failed before a PID existed, or the worker exited during
`worker_start`, was reaped without force, and reported no Excel PID, dialog,
timeout, or ambiguous cleanup. Every worker durably publishes `session_start`
before Excel session construction; that phase is the no-replay boundary.
Session startup, workbook, VBA, macro, protocol, dialog, timeout, and cleanup
failures are not startup-retried.

A workflow uses one deadline for all steps, saving, and cleanup. Durable
progress includes the current step ID, kind, index, count, and phase so a
parent-enforced timeout identifies the interrupted stage. A workflow is never
replayed after `session_start`, even if Excel has not yet reported its PID.

Only idempotent modification may use the same remaining attempt for a
recognized transient RPC disconnect after its owned Excel process is confirmed
stopped. Startup and modification retry policies cannot stack beyond two total
attempts. Callers do not enable or reproduce these policies themselves.

## Workbook behavior

- Source-management, lint, inspection, and modification sessions disable
  events before `Workbooks.Open`. A broken `Workbook_Open` therefore cannot run
  while xlvbatools is extracting, injecting, diffing, linting, listing,
  inspecting, or modifying a workbook. Non-executing operations force-disable
  macros; an enabled live compile test permits only its explicit post-open VBE
  Compile operation while events remain suppressed.
- Only explicit macro execution and workflow sessions opt into workbook code.
- Inspection is read-only with macros, events, and link updates disabled.
- Visible worksheets are the screenshot default.
- Hidden and VeryHidden worksheets require explicit opt-in.
- Mutating operations do not save after an operation exception.
- Strict named-range setup failures stop before macro invocation.
- Workflows stop after the first failed step, mark remaining steps `not_run`,
  and perform an explicit save only after every step succeeds.
- Screenshot rendering transfers a bitmap through a blank chart workbook; it
  does not copy worksheet VBA.
- Screenshot rendering makes only the owned Excel window renderable, forces a
  viewport scroll and range repaint, requests vector picture capture before
  bitmap fallback, validates native pixels before adding headers or grid
  overlays, and restores the prior visibility and `ScreenUpdating` state. A
  populated range that remains implausibly blank fails with
  `render_content_mismatch`; exhausted native capture fails with
  `screenshot_capture_failed`. Both retain bounded per-attempt evidence.

## Acceptance

For an Excel-backed result:

```python
result.require_success()
cleanup = result.require_clean_shutdown()
assert cleanup.is_clean
```

Operation success and clean shutdown are independent conditions. The result
preserves phase, error, request timing, dialog events, worker PID, Excel PID,
cleanup fields, and per-attempt retry evidence so wrappers never need private
process inspection.
