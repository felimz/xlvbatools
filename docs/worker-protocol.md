# Isolated Worker Protocol

xlvbatools uses one process-isolation protocol for every headless Excel
operation exposed by `XlvbaProject` and the CLI. The protocol version is `1.0`.

## Lifecycle

1. The parent writes a request to a private temporary directory.
2. It starts one directly tracked Python interpreter with standard output and
   error redirected to a regular file.
3. The worker atomically publishes its PID, phase, and owned Excel PID.
4. The worker performs exactly one operation and atomically writes its result.
5. The parent returns the result after the worker exits.
6. If the deadline expires, the parent terminates only the reported Excel PID,
   waits briefly for the worker, and terminates that exact worker if needed.

Workbook modification opts into one fresh-worker retry only for recognized
transient RPC-disconnect or call-rejected errors and only after the first owned
Excel process is confirmed stopped. The failed session does not save when its
mutation raises, so the retry does not commit a partial workbook change. Macro
execution and VBA injection are never automatically retried because their
effects may not be safely repeatable.

No COM proxy is serialized or returned to the parent. Existing Excel processes
are never selected by name and are never globally terminated.

## Request

```json
{
  "protocol_version": "1.0",
  "request_id": "uuid",
  "operation": "extract",
  "arguments": {
    "workbook_path": "C:/project/book.xlsm",
    "output_dir": "C:/project/vba_source",
    "component": null
  }
}
```

Paths passed to workers are absolute. Supported operations are `inspect`,
`run_macro`, `list_components`, `extract`, `inject`, `diff`, `lint_workbook`,
and `modify`.

## Progress

```json
{
  "protocol_version": "1.0",
  "request_id": "uuid",
  "operation": "extract",
  "worker_pid": 1234,
  "phase": "workbook_open",
  "excel_pid": 5678
}
```

Progress replacement is retried briefly on Windows because a concurrent reader
can momentarily deny file replacement. Readers therefore see either the prior
complete document or the next complete document, never partial JSON.

## Result

Every result carries `protocol_version`, `request_id`, `operation`, `success`,
`phase`, `worker_pid`, `excel_pid`, `elapsed_seconds`, `dialog_events`, and
`cleanup`. Operation output is stored under `data`. Failures add
`primary_error`, and timeouts add `timed_out` and `timeout_seconds`.

The facade converts this internal result into the stable public
`OperationResult` contract. Consumers should use `XlvbaProject` rather than
calling the internal executor directly.

## Compatibility Boundary

Function-level modules such as `xlvbatools.vba.extractor` remain available for
advanced in-process use. They deliberately preserve their historical return
types. Project wrappers and automation should use `XlvbaProject` so future
worker changes do not affect application code.
