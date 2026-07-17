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

## COM lifetime

- Worksheets, ranges, names, VBE components, and other child proxies are
  released while their owning session is still alive.
- A worker balances only the COM apartment initialization it performed.
- Teardown waits for graceful Excel/VBE exit before targeted termination.
- Sequential tests inspect stdout and stderr for native finalizer diagnostics;
  a zero process exit code alone is insufficient.

## Dialog safety

- Auto-dismissal starts only after Excel PID discovery and is PID-scoped.
- Cross-process window text and click operations are bounded by
  `SendMessageTimeoutW`.
- Captured events preserve multiline text and child-control diagnostics.
- Compile probing keeps the VBE hidden and does not invoke its File menu.

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

Only idempotent modification may use the same remaining attempt for a
recognized transient RPC disconnect after its owned Excel process is confirmed
stopped. Startup and modification retry policies cannot stack beyond two total
attempts. Callers do not enable or reproduce these policies themselves.

## Workbook behavior

- Inspection is read-only with macros, events, and link updates disabled.
- Visible worksheets are the screenshot default.
- Hidden and VeryHidden worksheets require explicit opt-in.
- Mutating operations do not save after an operation exception.
- Strict named-range setup failures stop before macro invocation.
- Screenshot rendering transfers a bitmap through a blank chart workbook; it
  does not copy worksheet VBA.

## Acceptance

For an Excel-backed result:

```python
result.require_success()
cleanup = result.require_clean_shutdown()
assert cleanup.still_running is False
```

Operation success and clean shutdown are independent conditions. The result
preserves phase, error, request timing, dialog events, worker PID, Excel PID,
cleanup fields, and per-attempt retry evidence so wrappers never need private
process inspection.
