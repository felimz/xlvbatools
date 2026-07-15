---
name: xlvba-toolchain
description: >
  Use xlvbatools v1 through its Project API and CLI for isolated VBA
  extraction, injection, linting, diffing, macro execution, snapshots,
  workbook inspection, modification, formatting, and dependency analysis.
---

# xlvbatools v1 Toolchain

Use this skill for headless Excel/VBA work. Application code must use
`xlvbatools.Project`; raw COM sessions and worker transport are internal.

## Edit-verify cycle

```text
1. Snapshot:  xlvba snapshot create --desc "before-change"
2. Extract:   xlvba extract
3. Edit:      change files under vba_source/
4. Lint:      xlvba lint
5. Inject:    xlvba inject
6. Diff:      xlvba diff
7. Run:       xlvba run <MacroName> --json
8. Pass:      commit workbook/source changes together
9. Fail:      xlvba snapshot restore latest
```

## CLI quick reference

| Command | Purpose |
|:---|:---|
| `xlvba extract` | Extract VBA into the configured source tree |
| `xlvba inject` | Inject configured VBA source into the workbook |
| `xlvba diff` | Compare live workbook VBA with extracted source |
| `xlvba lint` | Analyze source, or a live workbook with `--workbook` |
| `xlvba run <macro> --json` | Run one macro in an isolated worker |
| `xlvba snapshot` | Create, list, inspect, restore, diff, or prune checkpoints |
| `xlvba dump` | Inspect sheet data and optionally render screenshots |
| `xlvba modify` | Change cells, formulas, or named ranges |
| `xlvba debug` | Open Excel and the VBE for explicit interactive debugging |
| `xlvba search <pattern>` | Search extracted VBA |
| `xlvba fmt` | Format extracted VBA |
| `xlvba graph` | Generate Mermaid, DOT, or JSON call graphs |
| `xlvba version --json` | Report package, result-schema, and protocol versions |
| `xlvba agents` | Print agent integration guidance |

## Python API

```python
from xlvbatools import Project

project = Project.from_config()

lint_result = project.lint_source()
issues = lint_result.require_success()

run_result = project.run("OnCalculate", timeout=120, save=False)
run_result.require_success()
run_result.require_clean_shutdown()
```

For explicit paths:

```python
project = Project.open(
    "workbook/MyModel.xlsm",
    source="workbook/vba_source",
)
```

## Result handling

Every operation returns an `OperationResult`. Check:

- `success`, `phase`, and `error` for the operation outcome;
- `diagnostics.dialog_events` for captured Excel/VBE dialogs;
- `diagnostics.cleanup` for the owned Excel lifecycle;
- `artifacts` for screenshots and other durable outputs;
- `request_id`, `elapsed_seconds`, and `attempt_count` for tracing.

CLI `--json` output uses the same versioned envelope. Do not parse private
worker files.

## Troubleshooting

| Scenario | Action |
|:---|:---|
| Need module names | `xlvba extract --list --json` |
| Need a procedure or symbol | `xlvba search "FunctionName"` |
| Need source-only analysis | `xlvba lint --source vba_source/ --json` |
| Need live compile/analyzer validation | `xlvba lint --workbook workbook.xlsm --json` |
| Need call relationships | `xlvba graph --format json` |
| Need sheet values | `xlvba dump --sheets Sheet1 --data --json` |
| Need a visible-sheet screenshot | `xlvba dump --sheets Sheet1 --screenshot --json` |
| Need hidden sheets intentionally | Add `--include-hidden-sheets` |
| Need to set a value | `xlvba modify --sheet Sheet1 --cell A1 --value 42` |
| Macro timed out or Excel failed | Inspect cleanup diagnostics; never kill Excel globally |
| Need rollback | `xlvba snapshot restore latest` |

## Safety rules

1. Snapshot before risky workbook or VBA mutations.
2. Lint before injection; diff and run the relevant macro after injection.
3. Never terminate `EXCEL.EXE` by image name. Only xlvbatools may clean up
   the PID owned by its current operation.
4. Never use raw COM from application wrappers.
5. Avoid unguarded `MsgBox`, `FileDialog`, focus-dependent selection, and
   silent error suppression in headless VBA paths.
6. Hidden and VeryHidden worksheets are not rendered unless explicitly
   requested.
7. Ignore Excel owner-lock files whose names begin with `~$`.
8. Use the repository `.venv` for development and validation.
9. Pin released versions or exact full Git revisions in downstream projects.
10. Use forward slashes when presenting paths in portable logs and JSON.
