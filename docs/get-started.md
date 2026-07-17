# Get started

xlvbatools has two supported entry points:

- use the `xlvba` command from PowerShell for scripts, CI, and agent-driven
  workflows;
- import `Project` from `xlvbatools` when Python code needs typed results or a
  reusable application wrapper.

Both entry points use the same isolated worker and return the same
`OperationResult` contract. Do not build wrappers around COM, worker files, or
implementation subpackages.

## Requirements

- Python 3.10 or newer;
- Microsoft Excel on Windows for workbook, macro, and live-VBA operations;
- no Excel installation for extracted-source lint, search, format, or graph
  operations.

Use a project-local virtual environment so the command and Python import always
refer to the same pinned installation:

```powershell
py -3.12 -m venv .venv
& .\.venv\Scripts\python.exe -m pip install --upgrade pip

# Replace X.Y.Z with the exact approved release.
& .\.venv\Scripts\python.exe -m pip install "xlvbatools==X.Y.Z"

$xlvba = ".\.venv\Scripts\xlvba.exe"
& $xlvba version --text
```

For a reviewed revision that has not been released, use its complete Git commit
instead of an unpinned branch:

```powershell
$Revision = "<full-40-character-commit>"
& .\.venv\Scripts\python.exe -m pip install `
  "xlvbatools @ git+https://github.com/felimz/xlvbatools.git@$Revision"
```

## Initialize a project

From the consumer repository root, create configuration and install the
packaged agent guidance:

```powershell
$xlvba = ".\.venv\Scripts\xlvba.exe"
& $xlvba init --workbook "workbook/Model.xlsm" --agents --text
```

This writes `xlvbatools.toml` and copies missing guidance into `.agents/`.
Existing agent files are preserved. Use `xlvba agents install --force` only
when intentionally refreshing packaged paths.

A typical configuration is:

```toml
[xlvbatools]
workbook = "workbook/Model.xlsm"
vba_source = "workbook/vba_source"
snapshots_dir = "snapshots"
log_dir = "logs"
```

Paths are relative to the configuration file, not the shell's current
directory. Commands use these paths unless an explicit flag overrides them.

## Discover commands and flags

```powershell
& $xlvba --help             # terminal-oriented command list
& $xlvba help               # versioned JSON catalog for tools and agents
& $xlvba help inject        # JSON metadata for one command
& $xlvba inject --help      # descriptions and copy-ready examples
```

Flags belong after the command they configure. For nested commands, put
snapshot-wide flags before the action and action-specific flags after it:

```powershell
& $xlvba snapshot --workbook "workbook/Model.xlsm" create `
  --desc "before refactor" --milestone --text
```

## Common PowerShell operations

The CLI emits one machine-readable JSON envelope by default. That is the
recommended output for automation and agents. Add `--text` for concise prose
or `--table` for aligned rows only when a person is consuming stdout.

```powershell
# Source-only checks do not launch Excel.
& $xlvba lint --source "workbook/vba_source"
& $xlvba search "FileCount" --source "workbook/vba_source" --context 2
& $xlvba fmt --source "workbook/vba_source" --dry-run
& $xlvba graph --source "workbook/vba_source" --graph-format mermaid --text

# Excel-backed operations should carry an explicit, bounded timeout.
& $xlvba extract --output "workbook/vba_source" --timeout 120
& $xlvba inject --source "workbook/vba_source" --dry-run --timeout 120
& $xlvba inject --source "workbook/vba_source" --timeout 120
& $xlvba diff --source "workbook/vba_source" --summary --timeout 120
& $xlvba lint --workbook "workbook/Model.xlsm" --timeout 240
& $xlvba run "OnCalculate" --workbook "workbook/Model.xlsm" `
  --named-range "InputValue=42" --named-range 'Mode="Design"' `
  --no-save --timeout 120

# Inspection requires an explicit sheet selection.
& $xlvba dump --sheets "Input,Results" --data --range "A1:K100" --timeout 90
& $xlvba dump --sheets "Input" --screenshot --range "B91:K99" --timeout 90

# Workbook changes require a target and the intended value or formula.
& $xlvba modify --sheet "Input" --cell "C33" --value 42 --timeout 120
& $xlvba modify --sheet "Input" --cell "C34" --formula "=SUM(C1:C33)" `
  --timeout 120
```

Hidden and VeryHidden sheets are excluded by default. Add
`--include-hidden-sheets` only when those sheets are intentionally in scope.
Injection creates a backup by default; use `--no-backup` only when a verified
snapshot or equivalent rollback already exists.

### Flags worth choosing explicitly

| Flag | Use it when |
|:---|:---|
| `--workbook`, `-w` | overriding the configured workbook for one command |
| `--source`, `-s` | selecting a VBA source tree or source file |
| `--timeout` | bounding any Excel-backed operation; leave room for cleanup |
| `--dry-run` | previewing injection or formatting before writing |
| `--summary` | checking whether a diff exists without printing every changed line |
| `--sheets`, `--range` | constraining inspection to the required worksheet area |
| `--data`, `--screenshot` | choosing inspection artifacts deliberately |
| `--include-hidden-sheets` | explicitly including Hidden or VeryHidden sheets |
| `--named-range NAME=VALUE` | supplying one macro input; repeat for additional names |
| `--save`, `--no-save` | choosing whether macro workbook changes are persisted |
| `--visible` | deliberately showing the isolated owned Excel instance |
| `--text`, `--table` | requesting presentation output instead of default JSON |
| `--verbose` | collecting additional diagnostic logging for a failed operation |

Do not add `--no-backup` as a routine convenience flag. It disables a safety
default.

### Macro inputs, saving, and visibility

Repeat `--named-range NAME=VALUE` for each macro input. Values that are valid
JSON become native types; other values remain strings. Names are
case-insensitive and duplicates are rejected.

```powershell
& $xlvba run "OnCalculate" `
  --named-range "Count=42" `
  --named-range "Ratio=0.707" `
  --named-range "Enabled=true" `
  --named-range 'Mode="Design"' `
  --named-range "Label=North Sea" `
  --no-save `
  --timeout 120
```

Here `Count` is an integer, `Ratio` is a number, `Enabled` is a Boolean, and
the quoted JSON value for `Mode` plus the non-JSON value for `Label` are
strings. `null` becomes Python `None`, and `Name=` supplies an empty string.

`--save` is the default and preserves the existing CLI behavior. Use
`--no-save` when the run is validation-only. `--visible` is opt-in; it makes
the isolated Excel window visible without selecting or reusing a desktop Excel
instance.

## Parse CLI results in PowerShell

Capture stdout before converting it because a failed command still emits a
structured result and returns a nonzero exit code:

```powershell
$json = & $xlvba run "OnCalculate" --timeout 120
$exitCode = $LASTEXITCODE
$result = $json | ConvertFrom-Json

if ($exitCode -ne 0 -or -not $result.success) {
    throw $result.error.message
}
if ($result.diagnostics.cleanup.still_running) {
    throw "xlvbatools reported an owned Excel process still running"
}
```

## Use the Python API

Import supported application types from the package root. Do not import from
`xlvbatools.core`, `xlvbatools.vba`, `xlvbatools.macro`, or other backend
modules.

```python
from xlvbatools import Project

project = Project.from_config()

source_lint = project.lint_source()
issues = source_lint.require_success()
print(f"Source issues: {len(issues)}")

macro_result = project.run("OnCalculate", timeout=120, save=False)
macro = macro_result.require_success()
macro_result.require_clean_shutdown()
print(macro.run_id)
```

Use explicit paths when a wrapper does not own an `xlvbatools.toml` file:

```python
from xlvbatools import Project

project = Project.open(
    "workbook/Model.xlsm",
    source="workbook/vba_source",
)
```

Inspection and modification use keyword arguments that correspond to the CLI
flags:

```python
inspection_result = project.inspect(
    ["Input", "Results"],
    cell_range="A1:K100",
    include_data=True,
    include_screenshots=True,
    include_hidden_sheets=False,
    timeout=90,
)
inspection = inspection_result.require_success()
inspection_result.require_clean_shutdown()
print(inspection.screenshots)

change_result = project.modify(
    sheet="Input",
    cell="C33",
    value=42,
    timeout=120,
)
change = change_result.require_success()
change_result.require_clean_shutdown()
print(change.applied)
```

Every operation method above returns an `OperationResult`. Call
`require_success()` to obtain its typed data or raise `OperationFailedError`.
After an Excel-backed method, also call `require_clean_shutdown()` to verify
the owned Excel lifecycle. Source-only operations do not have Excel cleanup to
verify. `project.snapshots()` is the exception: it returns a `SnapshotService`
whose methods manage local checkpoints.

Other useful root-level imports include:

```python
from xlvbatools import (
    HeadlessCleanupError,
    OperationFailedError,
    Project,
    VBAIssue,
    get_version_info,
)
```

See the [Python API reference](api-reference.md) for the complete supported
import surface and method signatures.

## Choose the right entry point

| Need | PowerShell | Python |
|:---|:---|:---|
| Script or agent orchestration | Prefer the default JSON CLI | Use when already inside Python |
| Typed operation data | Parse the JSON envelope | Use `OperationResult` and typed outputs |
| Named-range macro inputs | Repeat `--named-range NAME=VALUE` | Use `Project.run(named_ranges=...)` |
| Control macro save or visibility | Use `--save`, `--no-save`, or `--visible` | Use `Project.run(save=..., visible=...)` |
| Offline wrapper tests | Invoke CLI integration tests sparingly | Inject an `Executor` test double |

Choose the entry point that best fits the caller; these macro controls now map
directly to the same supported `Project.run()` arguments.

## Safe edit and verify cycle

```powershell
& $xlvba snapshot create --desc "before change"
& $xlvba extract --timeout 120
& $xlvba lint --source "workbook/vba_source"
& $xlvba inject --dry-run --timeout 120
& $xlvba inject --timeout 120
& $xlvba diff --summary --timeout 120
& $xlvba run "OnCalculate" --no-save --timeout 120
```

Never terminate Excel by image name. xlvbatools cleans up only the worker and
Excel PID owned by the current operation. For deeper guidance, see
[Agent integration](agent-integration.md),
[Machine-first CLI output](cli-output.md), and
[Headless reliability](headless-reliability.md).
