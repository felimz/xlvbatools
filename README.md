# xlvbatools

Reliable, isolated Excel/VBA automation for Python projects.

xlvbatools gives downstream applications one small API for workbook
inspection, macro execution, one-session workflows, VBA source
synchronization, linting, and targeted modification. Every Excel-backed call
runs in a separately tracked process; raw COM objects never cross into the
caller.

## Install

```powershell
python -m pip install xlvbatools
```

Python 3.10 or newer is supported. Excel-backed operations require Microsoft
Excel on Windows. Source linting and source-code utilities do not require
Excel.

For a complete copy-ready setup in PowerShell and Python, including common
flags and result handling, start with [Get started](docs/get-started.md).

## Python API

Use a repository configuration:

```python
from xlvbatools import Project

project = Project.from_config()

inspection = project.inspect(
    ["Input"],
    cell_range="B91:K99",
    include_data=True,
    include_screenshots=True,
    timeout=60,
)
inspection.require_success()
inspection.require_clean_shutdown()

macro = project.run("OnCalculate", timeout=120, save=False)
macro.require_success()
macro.require_clean_shutdown()
```

Or construct a project directly:

```python
project = Project.open(
    "workbook/MyModel.xlsm",
    source="workbook/vba_source",
)
```

The public workflow methods are:

- `inspect`
- `run`
- `workflow`
- `list_components`
- `extract`
- `inject`
- `diff`
- `lint_source`
- `lint_workbook`
- `modify`
- `snapshots`

Every operation returns `OperationResult`. Its data uses immutable public
models such as `MacroOutput`, `ExtractionOutput`, `InjectionOutput`,
`ComponentDiff`, and `ModificationOutput`; no worker dictionaries escape into
application wrappers. Results also contain structured errors, artifacts,
timing, dialog diagnostics, exact worker and Excel PIDs, and cleanup reports.
Automatic retry evidence is typed under `diagnostics.attempts`; the executor
allows at most two total attempts under one timeout and never startup-retries
after the durable `session_start` boundary. `to_dict()` is the versioned JSON
form.

`project.snapshots()` returns a typed `SnapshotService`. It creates immutable
`SnapshotRecord` values and persists metadata with atomic writes and
crash-safe file locking.

Only names in `xlvbatools.__all__` are public. The implementation subpackages
are private worker backends and should not be used by application wrappers.

## CLI

```powershell
xlvba init --workbook workbook/Model.xlsm --agents
xlvba extract --timeout 120
xlvba inject --dry-run --timeout 120
xlvba inject --timeout 120
xlvba diff --summary --timeout 120
xlvba lint --source vba_source
xlvba run OnCalculate --named-range InputValue=42 --no-save --timeout 120
xlvba workflow --file workflow.json --no-save --timeout 240
xlvba dump --sheets Input --screenshot --range B91:K99 --timeout 90
xlvba modify --sheet Input --cell C33 --value 42 --timeout 120
xlvba snapshot create --desc "before change"
xlvba search "FileCount"
xlvba fmt --dry-run
xlvba graph
xlvba version
```

Discover the interface without scraping presentation text:

```powershell
xlvba --help          # conventional terminal help
xlvba help            # versioned JSON command catalog for agents
xlvba help extract    # JSON discovery for one command
xlvba extract --help  # complete options and copy-ready examples
```

Install the packaged agent guidance into an existing repository with
`xlvba agents install`. For a new repository, `xlvba init --agents` initializes
the configuration and installs the guidance together. The exact destination is
`.agents/` (plural). Installation is incremental: missing packaged files are
copied, existing files and project-specific extras are preserved, and only an
explicit `xlvba agents install --force` overwrites packaged file paths. Start
with `.agents/AGENTS.md`, then follow the referenced skill, rules, and workflow
for the task. Commit project-specific customizations with the repository.

All Excel-backed CLI commands use the same `Project` and executor as the Python
API. Every non-interactive command emits one complete `OperationResult` JSON
envelope to stdout by default. Request presentation output explicitly with
`--text`, `--table`, or `--output-format text|table`.

`xlvba graph --graph-format mermaid --text` renders Mermaid text. Graph payload
format and CLI presentation format are separate choices. The intentionally
interactive `xlvba debug` command remains text-based.

Hidden and VeryHidden worksheets are excluded from inspection by default. Use
`--include-hidden-sheets` only when hidden sheets are intentionally required.

## Architecture

```text
Python wrapper ─┐
                ├─> Project ─> OperationRequest ─> IsolatedExecutor
CLI command ────┘                                      │
                                                       v
                                         one owned worker per attempt
                                                       │
                                                       v
                                              one owned Excel PID
                                                       │
                                                       v
                                                OperationResult
```

The parent tracks only the worker it created and the Excel PID reported by that
worker. Timeouts and cleanup never terminate Excel by image name and never
select an unrelated desktop instance. A single automatic replay is possible
only for a proven pre-Excel worker-start failure; callers must not add another.

Related macro, range-write, and inspection steps can use `Project.workflow()`
to share one workbook open and one owned Excel PID. A workflow is fail-fast,
uses one overall timeout, saves only once after complete success when requested,
and is never replayed after `session_start`. See
[One-session workflows](docs/workflows.md).

Static linting uses one whole-project symbol index for both extracted files and
live workbooks, so cross-module public declarations resolve consistently.

## Configuration

```toml
[xlvbatools]
workbook = "workbook/MyModel.xlsm"
vba_source = "workbook/vba_source"
snapshots_dir = "snapshots"
log_dir = "logs"

[xlvbatools.snapshots]
rolling_limit = 10

[xlvbatools.backups]
limit = 5

[xlvbatools.lint]
disabled_rules = []
```

Paths are resolved relative to `xlvbatools.toml`, not the caller's current
directory.

## Versioning

Version 1.0.0 establishes the intentional public API. See
[Versioning and releases](docs/versioning.md) and [CHANGELOG.md](CHANGELOG.md).

## Development

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -e ".[dev]"
.venv\Scripts\ruff.exe check src tests
.venv\Scripts\mypy.exe --follow-imports=skip `
  src/xlvbatools/project.py `
  src/xlvbatools/execution.py `
  src/xlvbatools/results.py `
  src/xlvbatools/outputs.py `
  src/xlvbatools/workflow.py `
  src/xlvbatools/core/workflow.py `
  src/xlvbatools/workbook/dumper.py `
  src/xlvbatools/workbook/modifier.py `
  src/xlvbatools/snapshots.py `
  src/xlvbatools/cli
.venv\Scripts\python.exe -m pytest -m unit
.venv\Scripts\python.exe -m pytest
```

See [CONTRIBUTING.md](CONTRIBUTING.md) and
[release validation](docs/release-validation.md).

Documentation:

- [Get started](docs/get-started.md)
- [Python API](docs/api-reference.md)
- [One-session workflows](docs/workflows.md)
- [Agent integration](docs/agent-integration.md)
- [Inspection and modification](docs/dumper-and-modifier.md)
- [Linting and formatting](docs/lint-and-format.md)
- [Headless reliability contract](docs/headless-reliability.md)
- [Internal worker protocol](docs/worker-protocol.md)
- [Dialog watchdog architecture](docs/watchdog-architecture.md)
- [Versioning and releases](docs/versioning.md)
- [Machine-first CLI output](docs/cli-output.md)
- [Release validation](docs/release-validation.md)

## License

MIT
