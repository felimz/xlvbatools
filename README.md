# xlvbatools

Reliable, isolated Excel/VBA automation for Python projects.

xlvbatools gives downstream applications one small API for workbook
inspection, macro execution, VBA source synchronization, linting, and targeted
modification. Every Excel-backed call runs in a separately tracked process;
raw COM objects never cross into the caller.

## Install

```powershell
python -m pip install xlvbatools
```

Python 3.10 or newer is supported. Excel-backed operations require Microsoft
Excel on Windows. Source linting and source-code utilities do not require
Excel.

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
- `list_components`
- `extract`
- `inject`
- `diff`
- `lint_source`
- `lint_workbook`
- `modify`
- `snapshots`

Every operation returns `OperationResult`. It contains typed data, a structured
error, produced artifacts, request timing, dialog diagnostics, the exact worker
and Excel PIDs, and a cleanup report. `to_dict()` is the versioned JSON form.

Only names in `xlvbatools.__all__` are public. The implementation subpackages
are private worker backends and should not be used by application wrappers.

## CLI

```powershell
xlvba init
xlvba extract
xlvba inject
xlvba diff
xlvba lint
xlvba run OnCalculate --timeout 120
xlvba dump --sheets Input --screenshot --range B91:K99
xlvba modify --sheet Input --cell C33 --value 42
xlvba snapshot create --desc "before change"
xlvba search "FileCount"
xlvba fmt --dry-run
xlvba graph --format mermaid
xlvba version --json
```

All Excel-backed CLI commands use the same `Project` and executor as the Python
API. With `--json`, commands emit the complete `OperationResult` envelope.

Hidden and VeryHidden worksheets are excluded from inspection by default. Use
`--include-hidden-sheets` only when hidden sheets are intentionally required.

## Architecture

```text
Python wrapper ─┐
                ├─> Project ─> OperationRequest ─> IsolatedExecutor
CLI command ────┘                                      │
                                                       v
                                              one owned worker
                                                       │
                                                       v
                                              one owned Excel PID
                                                       │
                                                       v
                                                OperationResult
```

The parent tracks only the worker it created and the Excel PID reported by that
worker. Timeouts and cleanup never terminate Excel by image name and never
select an unrelated desktop instance.

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
.venv\Scripts\python.exe -m pytest -m unit
.venv\Scripts\python.exe -m pytest
```

See [CONTRIBUTING.md](CONTRIBUTING.md) and
[release validation](docs/release-validation.md).

Documentation:

- [Python API](docs/api-reference.md)
- [Agent integration](docs/agent-integration.md)
- [Inspection and modification](docs/dumper-and-modifier.md)
- [Linting and formatting](docs/lint-and-format.md)
- [Headless reliability contract](docs/headless-reliability-migration.md)
- [Internal worker protocol](docs/worker-protocol.md)
- [Dialog watchdog architecture](docs/watchdog-architecture.md)
- [Versioning and releases](docs/versioning.md)

## License

MIT
