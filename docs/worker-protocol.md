# Internal isolated-worker protocol

Protocol version: **2.0**.

This protocol is private to `IsolatedExecutor`. Applications use `Project`
and `OperationResult`; they must not create worker files or invoke the worker
entry point.

## Lifecycle

1. The parent validates an immutable `OperationRequest`.
2. It creates a private temporary directory and atomically writes request and
   progress JSON.
3. It starts one directly tracked interpreter with output redirected to a
   regular file.
4. The worker publishes its PID, phase, and exact owned Excel PID.
5. The worker performs exactly one operation and atomically writes one result.
6. The parent validates protocol version, request ID, operation identity, and
   result completeness.
7. `IsolatedExecutor` immediately converts transport data into a typed
   `OperationResult`.
8. Timeout cleanup targets only the reported Excel PID and exact worker.

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
- dialog events and cleanup;
- normalized operation `data`.

For inspection, `data` contains `workbook_data` and `screenshots`. For a
macro, `data` contains the run ID, return value, and macro-specific fields.

The transport dictionary is discarded after conversion. Public JSON comes
from `OperationResult.to_dict()` and is governed independently by
`RESULT_SCHEMA_VERSION`.

## Retry policy

Only modification opts into a single fresh-worker retry, and only for a
recognized transient RPC disconnect after the first owned Excel process is
confirmed stopped. Macro execution and injection are not retried because their
effects may be non-idempotent.
