# Internal isolated-worker protocol

Protocol version: **2.0**.

This protocol is private to `IsolatedExecutor`. Applications use `Project`
and `OperationResult`; they must not create worker files or invoke the worker
entry point.

## Lifecycle

1. The parent validates an immutable `OperationRequest`.
2. It creates a private temporary directory and atomically writes request and
   progress JSON.
3. It starts one directly tracked interpreter attempt with output redirected
   to a regular file.
4. The worker publishes `worker_start`, then durably publishes `session_start`
   before any code may construct an Excel session. It publishes the exact
   owned Excel PID only after Excel starts.
5. The worker performs exactly one operation and atomically writes one result.
6. The parent validates protocol version, request ID, operation identity, and
   result completeness.
7. `IsolatedExecutor` immediately converts transport data into a typed
   `OperationResult`.
8. The parent reaps the exact worker and records `worker_exit` independently
   from the Excel `cleanup` report.
9. Timeout cleanup targets only the reported Excel PID and exact worker.

No COM object is serialized and no existing Excel instance is selected.

## Request

```json
{
  "protocol_version": "2.0",
  "request_id": "uuid",
  "operation": "extract",
  "arguments": {
    "workbook_path": "C:/project/book.xlsm",
    "output_dir": "C:/project/vba_source",
    "component": null
  }
}
```

Supported transport operations are `inspect`, `run_macro`,
`list_components`, `extract`, `inject`, `diff`, `lint_workbook`, and
`modify`.

## Progress

```json
{
  "protocol_version": "2.0",
  "request_id": "uuid",
  "operation": "extract",
  "worker_pid": 1234,
  "phase": "workbook_open",
  "excel_pid": 5678
}
```

Atomic replacement is retried because Windows readers can briefly deny a
rename. Readers therefore observe complete JSON documents.

## Result envelope

Operation-specific values appear only under `data`; they are not duplicated
at the transport top level. Common transport fields include:

- protocol, request, and operation identity;
- `success`, `phase`, and failure evidence;
- worker and Excel PIDs;
- elapsed time and attempt count;
- dialog events, Excel cleanup, and separate worker-exit/reaping evidence;
- normalized operation `data`.

For inspection, `data` contains `workbook_data` and `screenshots`. For a
macro, `data` contains the run ID, return value, and macro-specific fields.

The transport dictionary is discarded after conversion. Public JSON comes
from `OperationResult.to_dict()` and is governed independently by
`RESULT_SCHEMA_VERSION`.

## Retry policy

`IsolatedExecutor` owns all retry decisions and permits at most two total
attempts under one original timeout budget. The automatic startup retry is
allowed only when either child-process creation failed before a worker PID
existed, or a worker failed during `worker_start` and was proven exited and
reaped with no Excel PID, dialog, timeout, forced termination, or ambiguous
cleanup. `session_start` is the no-replay boundary: that phase and every later
phase are never startup-retried.

Modification may use that same one remaining attempt for a recognized
transient RPC disconnect after the first owned Excel process is confirmed
stopped. Startup and transient policies cannot stack into more than two
attempts. Macro/VBA failures, protocol failures, timeouts, dialogs, and
post-ownership failures are never replayed automatically.
