---
name: xlvba-toolchain
description: >
  Use xlvbatools v1 through its Project API and CLI for isolated VBA
  extraction, injection, linting, diffing, macro execution, snapshots,
  one-session workflows, workbook inspection, modification, formatting, and
  dependency analysis.
---

# xlvbatools v1 Toolchain

Use this skill for headless Excel/VBA work. Application code must use
`xlvbatools.Project`; raw COM sessions and worker transport are internal.

Before using the commands, install a pinned xlvbatools version in the project
`.venv` and create or verify `xlvbatools.toml`. Installing this `.agents/`
template alone does not install the package or configure a workbook.

## Start here

For a new consumer repository, follow `.agents/workflows/get-started.md` before
the task-specific workflow. It verifies the project-local executable, installs
or checks `xlvbatools.toml`, demonstrates common flags, and shows the supported
root-level Python imports.

Use `xlvba help COMMAND` for machine-readable flag discovery or
`xlvba COMMAND --help` for terminal help. Put flags after the command they
configure. Give Excel-backed operations an explicit `--timeout`; use
`--dry-run` before supported mutations. Default JSON is for tools and agents,
while `--text` and `--table` are explicit presentation choices.

## Edit-verify cycle

```text
1. Snapshot:  xlvba snapshot create --desc "before-change"
2. Extract:   xlvba extract --timeout 120
3. Edit:      change files under vba_source/
4. Lint:      xlvba lint --source vba_source
5. Preview:   xlvba inject --dry-run --timeout 120
6. Inject:    xlvba inject --timeout 120
7. Diff:      xlvba diff --comparison vba --summary --timeout 120
8. Run:       xlvba run <MacroName> --timeout 120
9. Pass:      commit workbook/source changes together
10. Fail:     xlvba snapshot restore latest
```

## CLI quick reference

| Command | Purpose |
|:---|:---|
| `xlvba help [command]` | Return versioned machine-readable CLI discovery |
| `xlvba extract` | Extract VBA into the configured source tree |
| `xlvba inject` | Inject configured VBA source into the workbook |
| `xlvba diff` | Compare VBA tokens by default, or raw text explicitly |
| `xlvba lint` | Analyze source/live VBA with filters and reviewed baselines |
| `xlvba run <macro>` | Run one macro with optional inputs and lifecycle flags |
| `xlvba workflow --file <path>` | Run ordered typed steps in one Excel session |
| `xlvba snapshot` | Create, list, inspect, restore, diff, or prune checkpoints |
| `xlvba dump` | Inspect sheet data and optionally render screenshots |
| `xlvba modify` | Change cells, formulas, or named ranges |
| `xlvba debug` | Open Excel and the VBE for explicit interactive debugging |
| `xlvba search <pattern>` | Search extracted VBA |
| `xlvba fmt` | Format extracted VBA |
| `xlvba graph` | Generate Mermaid, DOT, or JSON call graphs |
| `xlvba version` | Report package, result-schema, and protocol versions |
| `xlvba agents` | Print agent integration guidance |
| `xlvba agents install` | Safely install packaged guidance into `.agents/` |

For an existing consumer repository, run `xlvba agents install`, then read
`.agents/AGENTS.md` and this skill. For a new project, `xlvba init --agents`
performs project initialization and template installation together. Existing
template paths are preserved unless `xlvba agents install --force` is used;
project-specific extra files are never deleted.

## Python API

```python
from xlvbatools import InspectStep, MacroStep, ModifyStep, Project

project = Project.from_config()

lint_result = project.lint_source()
issues = lint_result.require_success()

project.lint_source(write_baseline=".xlvba/lint-baseline.json")
new_lint = project.lint_source(
    baseline=".xlvba/lint-baseline.json", new_only=True,
)

diff_result = project.diff(comparison="vba", timeout=120)

run_result = project.run("OnCalculate", timeout=120, save=False)
macro = run_result.require_success()
print(macro.run_id)
run_result.require_clean_shutdown()

workflow_result = project.workflow(
    [
        MacroStep("retrieve", "OnRetrieve"),
        ModifyStep("inputs", "Input", {"C102:C104": [[0.1], [0.0], [-0.1]]}),
        MacroStep("calculate", "OnCalculate"),
        InspectStep("results", ("Input",), include_screenshots=False),
    ],
    timeout=240,
    save=False,
)
workflow = workflow_result.require_success()
workflow_result.require_clean_shutdown()

snapshot = project.snapshots().create("before change")
project.snapshots().restore(snapshot)
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
- `InspectionOutput.screenshot_diagnostics` for native picture-format,
  clipboard, viewport, off-screen-window, attempt, and pixel evidence;
- `request_id`, `elapsed_seconds`, and `attempt_count` for tracing.
- `diagnostics.attempts` when `attempt_count` is `2`; retain the phase,
  worker-exit evidence, and retry reason instead of adding another retry.

Operation data is modeled, not a private dictionary: use attributes on
`MacroOutput`, `ExtractionOutput`, `InjectionOutput`, `ComponentDiff`,
`ModificationOutput`, `WorkflowOutput`, and `SnapshotRecord`.

CLI output is one versioned JSON envelope by default. Use `--text` or `--table`
only when presentation output is explicitly requested. Do not parse private
worker files. `IsolatedExecutor` automatically replays only a proven failure
before Excel ownership and never after `session_start`; wrappers must not
duplicate that policy.

## Troubleshooting

| Scenario | Action |
|:---|:---|
| Need module names | `xlvba extract --list` |
| Need a procedure or symbol | `xlvba search "FunctionName"` |
| Need source-only analysis | `xlvba lint --source vba_source/` |
| Need only lint regressions | Add `--baseline <path> --new-only` |
| Need a reviewed lint snapshot | Add `--write-baseline <path>` |
| Need raw case-sensitive VBA differences | `xlvba diff --comparison text` |
| Need live compile/analyzer validation | `xlvba lint --workbook workbook.xlsm` |
| Need call relationships | `xlvba graph` |
| Need sheet values | `xlvba dump --sheets Sheet1 --data` |
| Need partial cell formatting | Add `--rich-text`; cell-run output is bounded |
| Need a visible-sheet screenshot | `xlvba dump --sheets Sheet1 --screenshot` |
| Need hidden sheets intentionally | Add `--include-hidden-sheets` |
| Need to set a value | `xlvba modify --sheet Sheet1 --cell A1 --value 42` |
| Need macro inputs without saving | Repeat `--named-range NAME=VALUE` and add `--no-save` |
| Need visible macro execution | Add `--visible`; the Excel instance remains isolated and owned |
| Related steps need one workbook open | Use typed `Project.workflow()` steps or `xlvba workflow --file ...` |
| Macro timed out or Excel failed | Inspect cleanup diagnostics; never kill Excel globally |
| Screenshot reports `render_content_mismatch` | Retain the attempt metrics and workbook data; do not accept the PNG as visual evidence |
| Screenshot reports `screenshot_capture_failed` | Retain structured attempt evidence; do not add a retry after `session_start` |
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
11. Source management and live lint must remain non-executing. Do not work
    around the library's suppressed startup events or hidden VBE; use `run`,
    `workflow`, or explicit `debug` only when workbook code is intended.
