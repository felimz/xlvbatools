# Agent integration

xlvbatools v1 gives coding agents the same supported surface as any other
application: the `Project` Python API and the `xlvba` CLI. Agents must not
manage COM sessions or private worker files directly.

Install a pinned release into the repository virtual environment before using
the CLI:

```powershell
.venv\Scripts\python.exe -m pip install "xlvbatools==X.Y.Z"
```

Replace `X.Y.Z` with the exact approved release. For an unreleased reviewed
revision, pin its exact full Git commit instead.
Installing the Python package does not modify the repository's `.agents/` tree
or create `xlvbatools.toml`; those are explicit steps below.

For a copy-ready installation, configuration, PowerShell flag, and Python
import walkthrough, begin with [Get started](get-started.md).

## Discover the CLI

Use conventional help when reading in a terminal and the versioned JSON
catalog when an agent needs to discover commands programmatically:

```powershell
xlvba --help
xlvba help
xlvba help lint
xlvba lint --help
```

`xlvba help` returns one `OperationResult` envelope containing the discovery
schema version, command purposes, usage, examples, Excel requirements, the
actual parser flags, defaults, choices, nested subcommands, the output contract,
and agent-template installation commands. Option metadata is derived from the
execution parser so it cannot drift into a separate hand-maintained list.
Agents should parse this catalog instead of scraping human-formatted help text.

## Install the packaged templates

The installation directory is `.agents/` (plural). Add guidance to an existing
repository without modifying its xlvbatools configuration:

```powershell
xlvba agents install
```

For a new repository, initialize the configuration and install the same
packaged guidance in one operation:

```powershell
xlvba init --workbook workbook/MyProject.xlsm --agents
xlvba version
```

Installation copies every missing packaged file and skips existing paths, so a
partial `.agents/` tree is repaired without overwriting customizations. To
refresh packaged file paths deliberately, run `xlvba agents install --force`.
Even in force mode, xlvbatools does not delete project-specific extra files.
The JSON result lists `installed`, `skipped`, and `overwritten` paths.

After installation:

1. Read `.agents/AGENTS.md` for the repository-wide contract.
2. Follow `.agents/workflows/get-started.md` to verify the local installation
   and configuration.
3. Read `.agents/skills/xlvba-toolchain/SKILL.md` before Excel/VBA work.
4. Select the relevant Python or VBA rule file.
5. Follow `vba-edit.md` for changes or `vba-debug.md` for diagnosis.
6. Customize paths or project-specific acceptance commands and commit those
   changes with the repository.

## Standard edit-verify cycle

```powershell
xlvba snapshot create --desc "before change"
xlvba extract --timeout 120
# Edit files under vba_source/
xlvba lint --source vba_source
xlvba inject --dry-run --timeout 120
xlvba inject --timeout 120
xlvba diff --summary --timeout 120
xlvba run OnCalculate --no-save --timeout 120
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

PowerShell callers can provide the same common macro controls without a Python
wrapper:

```powershell
xlvba run OnCalculate --named-range InputValue=42 --no-save --timeout 120
```

Repeat `--named-range NAME=VALUE` as needed. Valid JSON values are typed;
other values remain strings. Add `--visible` only when the isolated owned Excel
window is intentionally required.

For unit tests, inject an `Executor` test double into `Project`. Do not mock
or expose COM objects in downstream wrappers.

## Workbook inspection

```powershell
xlvba dump --sheets "Input,Results" --data
xlvba dump --sheets Input --screenshot --range B91:K99
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
xlvba lint --workbook workbook/MyProject.xlsm --timeout 240
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
    ├── get-started.md
    ├── vba-debug.md
    └── vba-edit.md
```
