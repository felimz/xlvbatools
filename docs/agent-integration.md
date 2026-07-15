# Agent integration

xlvbatools v1 gives coding agents the same supported surface as any other
application: the `Project` Python API and the `xlvba` CLI. Agents must not
manage COM sessions or private worker files directly.

## Bootstrap

Initialize a new repository and install the packaged agent guidance:

```powershell
xlvba init --workbook workbook/MyProject.xlsm --agents
xlvba version --json
```
Template installation is non-destructive. If `.agents/` already exists,
xlvbatools leaves it unchanged; update customized guidance through source
control rather than silently overwriting it.

## Standard edit-verify cycle

```powershell
xlvba snapshot create --desc "before change"
xlvba extract
# Edit files under vba_source/
xlvba lint --source vba_source --json
xlvba inject --json
xlvba diff --summary
xlvba run OnCalculate --timeout 120 --json
```

A passing macro is not the only acceptance condition. In JSON output, verify
both the operation outcome and the owned-process cleanup:

- `success` is true;
- `error` is null;
- `diagnostics.cleanup.still_running` is false;
- `diagnostics.cleanup.worker_terminated` is false;
- no unexpected `diagnostics.dialog_events` were captured.

Restore the checkpoint if validation fails.

## Python wrapper pattern

```python
from xlvbatools import Project

project = Project.open(
    "workbook/MyProject.xlsm",
    source="workbook/vba_source",
)

lint_result = project.lint_source()
issues = lint_result.require_success()

run_result = project.run("OnCalculate", timeout=120, save=False)
run_result.require_success()
run_result.require_clean_shutdown()
```

For unit tests, inject an `Executor` test double into `Project`. Do not mock
or expose COM objects in downstream wrappers.

## Workbook inspection

```powershell
xlvba dump --sheets "Input,Results" --data --json
xlvba dump --sheets Input --screenshot --range B91:K99 --json
```

Only visible worksheets are rendered by default. Add
`--include-hidden-sheets` only when Hidden or VeryHidden content is explicitly
required.

```python
inspection = project.inspect(
    ["Input"],
    cell_range="B91:K99",
    include_data=True,
    include_screenshots=True,
)
output = inspection.require_success()
inspection.require_clean_shutdown()
print(output.workbook_data)
print(output.screenshots)
```

## Error handling

| Evidence | Agent action |
|:---|:---|
| `compile_error` | Fix the reported module, line, and column; inject and rerun |
| `runtime_error` | Use the captured VBA error number and description |
| `msgbox` or `file_dialog` | Guard interactive VBA with `Application.UserControl` |
| `named_range_setup` | Correct required inputs before invoking the macro again |
| `timeout` | Inspect cleanup; never terminate unrelated Excel processes |
| Trust Center failure | Ask the user to enable programmatic VBA project access |

Use `xlvba debug` only when visible interactive debugging is deliberately
requested. Ordinary automation remains isolated and headless.

## Analysis behavior

Source lint and live-workbook lint use the same whole-project symbol index.
Public declarations in standard modules therefore resolve across modules.
Live lint can additionally request Excel's compile test:

```powershell
xlvba lint --workbook workbook/MyProject.xlsm --json --timeout 240
```

Treat Excel compilation as the semantic authority. Do not suppress a reported
cross-module symbol merely to make modes agree.

## Safety rules

- Never run `taskkill /im EXCEL.EXE`.
- Never select an existing desktop Excel instance for headless work.
- Never parse or construct the private worker request/result files.
- Never render hidden worksheets without explicit authorization.
- Never commit workbook and extracted VBA states that fail `xlvba diff`.
- Pin a released package version or an exact full Git revision downstream.

## Installed agent layout

```text
.agents/
├── AGENTS.md
├── rules/
│   ├── python-rules.md
│   └── vba-rules.md
├── skills/
│   └── xlvba-toolchain/
│       └── SKILL.md
└── workflows/
    ├── vba-debug.md
    └── vba-edit.md
```
