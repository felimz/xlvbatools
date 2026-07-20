# Python API reference

## Supported boundary

Only names listed in `xlvbatools.__all__` are supported. The stable application
surface consists of:

- `Project` and `ProjectSettings`;
- `Operation`, `OperationRequest`, `Executor`, and `IsolatedExecutor`;
- `OperationResult` and its typed diagnostic, error, artifact, inspection,
  workflow, operation-output, and snapshot models;
- documented public exceptions, `VBAIssue`, and version helpers.

Modules under `core`, `vba`, `macro`, `workbook`, `analysis`, and
`snapshot` are private backends.

### Exact public exports

The following names are the complete supported import surface:

| Area | Public names |
|:---|:---|
| Project | `Project`, `ProjectSettings` |
| Execution | `Operation`, `OperationRequest`, `Executor`, `IsolatedExecutor` |
| Result envelope | `OperationResult`, `RESULT_SCHEMA_VERSION` |
| Result evidence | `ErrorInfo`, `Diagnostics`, `CleanupReport`, `WorkerExitReport`, `AttemptDiagnostic`, `Artifact` |
| Inspection data | `InspectionOutput` |
| VBA component data | `VBAComponent`, `ExtractionOutput`, `InjectionChange`, `InjectionOutput` |
| VBA action data | `ComponentDiff`, `MacroOutput`, `ModificationOutput` |
| Workflow request | `WorkflowStep`, `MacroStep`, `ModifyStep`, `InspectStep`, `WORKFLOW_SCHEMA_VERSION` |
| Workflow result | `WorkflowOutput`, `WorkflowStepResult`, `ModifyStepOutput`, `RangeWriteResult` |
| Snapshots | `SnapshotService`, `SnapshotRecord`, `SnapshotGitInfo` |
| Analysis | `VBAIssue`, `LINT_BASELINE_SCHEMA_VERSION` |
| Operation errors | `XlvbaError`, `ConfigurationError`, `OperationFailedError`, `HeadlessCleanupError` |
| Excel and snapshot errors | `TrustCenterError`, `SnapshotError`, `SnapshotNotFoundError` |
| Versioning | `__version__`, `VersionInfo`, `get_version_info` |

Anything not listed here and in `xlvbatools.__all__` is an implementation
detail, even if it can be imported from a source checkout.

## Project construction

Load the nearest `xlvbatools.toml`:

```python
from xlvbatools import Project

project = Project.from_config()
```

Use explicit paths:

```python
project = Project.open(
    "workbook/MyModel.xlsm",
    source="workbook/vba_source",
)
```

Construct immutable settings when a wrapper owns configuration:

```python
from pathlib import Path
from xlvbatools import Project, ProjectSettings

settings = ProjectSettings(
    workbook=Path("workbook/MyModel.xlsm"),
    source=Path("workbook/vba_source"),
    snapshots=Path("snapshots"),
    disabled_lint_rules=("PF003",),
    backup_limit=5,
    snapshot_limit=10,
)
project = Project(settings)
```

Paths are resolved at construction. Workbooks must use `.xlsm`, `.xlsb`, or
`.xls`.

## Project methods

| Method | Excel worker | Result data |
|:---|:---:|:---|
| `list_components(timeout=60)` | yes | tuple of `VBAComponent` |
| `inspect(sheets, ..., timeout=60)` | yes | `InspectionOutput` |
| `run(macro, ..., timeout=120)` | yes | `MacroOutput` |
| `workflow(steps, ..., timeout=240)` | yes, one session for all steps | `WorkflowOutput` |
| `extract(..., timeout=120)` | yes | `ExtractionOutput` |
| `inject(..., timeout=120)` | yes, except dry-run backend | `InjectionOutput` |
| `diff(..., comparison="vba", timeout=120)` | yes | tuple of `ComponentDiff` |
| `lint_workbook(..., severities=None, rules=None, baseline=None, new_only=False, write_baseline=None)` | yes | tuple of `VBAIssue` |
| `modify(..., timeout=120)` | yes | `ModificationOutput` |
| `lint_source(source=None, ..., severities=None, rules=None, baseline=None, new_only=False, write_baseline=None)` | no | tuple of `VBAIssue` |
| `snapshots()` | no | `SnapshotService` |

All Excel-backed calls cross the configured `Executor`. Raw COM proxies never
leave the worker.

### Inspection

```python
result = project.inspect(
    ["Input", "Results"],
    output_dir="artifacts/screenshots",
    cell_range="A1:K100",
    include_data=True,
    include_screenshots=True,
    include_rich_text=True,
    output_json="artifacts/workbook.json",
    include_hidden_sheets=False,
    timeout=90,
)
inspection = result.require_success()
result.require_clean_shutdown()
```

`InspectionOutput.workbook_data` contains structured workbook data and
`InspectionOutput.screenshots` maps sheet names to render paths or status
messages. Generated screenshots are also listed as `Artifact` records.
Hidden and VeryHidden sheets are excluded unless explicitly enabled.
Partial rich-text runs are excluded by default because each populated cell
requires additional COM inspection. With `include_rich_text=True`, each cell
adds bounded, 1-based font runs and a `complete`, `truncated`, or `unsupported`
status. Inspection is capped at 4,096 characters and 256 runs per cell.
Native pixels are validated before headers and gridlines are added. A populated
range that remains implausibly blank returns
`error.code == "render_content_mismatch"` with per-attempt metrics and is not
published as a successful artifact.

### Macro execution

```python
result = project.run(
    "OnCalculate",
    named_ranges={"InputValue": 42},
    timeout=120,
    visible=False,
    save=False,
    strict_named_ranges=True,
)
data = result.require_success()
cleanup = result.require_clean_shutdown()
```

`MacroOutput` exposes `macro`, `run_id`, and `excel_pid`. Additional
JSON-compatible worker values are retained in its immutable `details` mapping.

The deadline covers worker startup, workbook open, input setup, macro execution,
save, and cleanup. Macro execution is never automatically retried.

The equivalent CLI controls are repeatable `--named-range NAME=VALUE`,
`--save`/`--no-save`, and opt-in `--visible`:

```powershell
xlvba run OnCalculate --named-range InputValue=42 --no-save --timeout 120
```

### One-session workflow

```python
from xlvbatools import InspectStep, MacroStep, ModifyStep

result = project.workflow(
    [
        MacroStep("retrieve", "OnRetrieve"),
        ModifyStep("inputs", "Input", {"C102:C104": [[0.1], [0.0], [-0.1]]}),
        MacroStep("calculate", "OnCalculate"),
        InspectStep("results", ("Input",), include_screenshots=False),
    ],
    timeout=240,
    save=False,
)
workflow = result.require_success()
result.require_clean_shutdown()
results = workflow.step("results")
```

All steps use one worker, one Excel process, and one workbook open. Execution
is ordered and fail-fast. `WorkflowStepResult.status` is `succeeded`, `failed`,
or `not_run`; `WorkflowOutput.failed_step_id` identifies the first failure.
Saving defaults off and, when requested, occurs once after every step succeeds.
The complete workflow has one timeout and is not replayed after
`session_start`. See [One-session workflows](workflows.md) for the complete
Python and versioned CLI contracts.

### Source and workbook lint

```python
source_result = project.lint_source(
    severities=("ERROR", "WARNING"),
    rules=("IP001", "DV001"),
    write_baseline="artifacts/lint-baseline.json",
)
workbook_result = project.lint_workbook(
    compile_test=True,
    baseline="artifacts/lint-baseline.json",
    new_only=True,
    timeout=240,
)
```

Both adapters use the same project-level symbol index. Workbook lint can add
Excel compile evidence. A lint result is unsuccessful when ERROR-severity
issues exist.
Live lint disables workbook events before opening, keeps the owned VBE hidden,
and fails closed if Excel compilation cannot be verified. `DV001` detects
duplicate declarations statically in both source and live-workbook analysis.

Severity and rule filters are inclusive and can be combined. Baselines contain
all unfiltered findings and are written atomically. Their stable fingerprints
exclude line numbers, normalize VBA casing and whitespace, and compare
duplicates as a multiset. Consequently, moving a known finding does not make
it new, while adding a second identical finding does. `new_only=True` requires
`baseline`. Only selected ERROR findings determine lint success; lifecycle,
transport, and cleanup failures are never suppressed by lint filters.

### VBA-aware diff

`Project.diff()` compares VBA tokens by default. Identifier and keyword case,
plus insignificant spacing between code tokens, produce
`ComponentDiff.status == "equivalent"`. String literals, comments, punctuation,
and line structure remain exact. Use `comparison="text"` for a raw,
case-sensitive line diff.

### Snapshots

```python
snapshots = project.snapshots()
record = snapshots.create("before refactor")
snapshots.restore(record)
```

`SnapshotService` returns immutable `SnapshotRecord` values. Snapshot metadata
is written atomically under an operating-system file lock; missing snapshots
raise `SnapshotNotFoundError` during restore or diff.

## OperationResult

`OperationResult[T]` contains:

- `schema_version`;
- `operation`, `success`, and `phase`;
- optional typed `data` and `ErrorInfo`;
- `warnings`, `artifacts`, and `metadata`;
- `request_id`, `elapsed_seconds`, and `attempt_count`;
- `Diagnostics` with dialog events, worker PID, Excel PID, COM evidence, an
  optional `CleanupReport`, an optional `WorkerExitReport`, and typed
  `AttemptDiagnostic` entries, plus the latest durable `progress` state.

```python
if not result.success:
    assert result.error is not None
    print(result.error.code, result.error.message)

payload = result.to_dict()
```

`require_success()` returns data or raises `OperationFailedError`.
`require_clean_shutdown()` returns the cleanup report or raises
`HeadlessCleanupError`. In Python, `CleanupReport.is_clean` is the derived
cleanliness check. Serialized JSON exposes the underlying cleanup fields, not
that derived property.

`WorkerExitReport` is deliberately separate from `CleanupReport`: the former
proves whether the isolated Python worker exited and was reaped, while the
latter describes only the owned Excel lifecycle. `WorkerExitReport.is_clean`
means the worker exited and was reaped without forced termination.

`Diagnostics.attempts` contains one `AttemptDiagnostic` per executor attempt.
Each entry records its number, phase, request ID, error code/message/details
(including captured worker output), worker evidence, Excel PID, Excel cleanup,
dialog count, elapsed time, and any executor-owned retry decision.
`retryable=True` and `retry_reason` mean the executor actually proceeded to
one final attempt; callers should log this evidence, not add another retry.

## Custom executors

Thin wrappers can inject a test double or another transport implementing the
`Executor` protocol:

```python
from xlvbatools import IsolatedExecutor, Project

project = Project.from_config(executor=IsolatedExecutor())
```

Advanced code may submit an enumerated operation:

```python
from xlvbatools import Operation

result = project.execute(
    Operation.LIST_COMPONENTS,
    {"workbook_path": str(project.workbook)},
    timeout=60,
)
```

`OperationRequest` requires a positive timeout and freezes a copy of its
arguments. Applications should prefer the named `Project` methods.

## CLI serialization

Every non-interactive CLI command serializes this same `OperationResult`
envelope to stdout by default, including local commands such as `search`,
`format`, `graph`, `snapshot`, and `version`. Text and table output are explicit
presentation options. See [Machine-first CLI output](cli-output.md).

## Versions

```python
from xlvbatools import (
    LINT_BASELINE_SCHEMA_VERSION,
    __version__,
    get_version_info,
)

info = get_version_info()
print(__version__)
print(info.result_schema_version)
print(info.worker_protocol_version)
print(info.workflow_schema_version)
print(LINT_BASELINE_SCHEMA_VERSION)
```

`VersionInfo.version` is the authoritative version embedded in the imported
code. `distribution_version` reports installed metadata separately and
`version_mismatch` makes a stale editable installation explicit. Released
wheels must report matching values.

See [Versioning and releases](versioning.md).
