---
name: xlvba-toolchain
description: >
  How to use the xlvbatools CLI and Python API for VBA extraction, injection,
  linting, diffing, macro execution, snapshot/rollback, workbook inspection,
  code formatting, and dependency graph analysis.
---

# xlvba-toolchain Skill

This skill teaches agents how to use `xlvbatools` for headless VBA development.

## Core Workflow: Edit-Verify Cycle

```
1. Snapshot:    xlvba snapshot create --desc "before-change"
2. Extract:    xlvba extract
3. Edit:       Modify vba_source/modules/<name>.bas
4. Lint:       xlvba lint
5. Inject:     xlvba inject
6. Run:        xlvba run <MacroName> --json
7. If PASS:    git add vba_source/ && git commit
8. If FAIL:    xlvba snapshot restore latest
```

## CLI Quick Reference

| Command | Purpose |
|:---|:---|
| `xlvba extract` | Extract VBA from workbook to vba_source/ |
| `xlvba inject` | Inject vba_source/ files into workbook |
| `xlvba diff` | Compare workbook VBA vs vba_source/ |
| `xlvba lint` | Run static analysis (30+ built-in rules) |
| `xlvba run <macro>` | Execute macro with dialog protection |
| `xlvba snapshot` | Checkpoint and rollback (create/list/restore/prune) |
| `xlvba dump` | Dump sheet data and screenshots |
| `xlvba modify` | Modify cells, formulas, named ranges |
| `xlvba debug` | Open Excel + VBE visibly for debugging |
| `xlvba search <pattern>` | Search VBA source files |
| `xlvba fmt` | Format VBA code (normalize indentation) |
| `xlvba graph` | Generate call dependency graph (Mermaid/DOT/JSON) |
| `xlvba agents` | Show AI agent integration help and best practices |

## Troubleshooting Matrix

| Scenario | Command |
|:---|:---|
| Need to see what VBA modules exist | `xlvba extract --list` |
| Need to find a specific function | `xlvba search "FunctionName"` |
| Need to check for common VBA issues | `xlvba lint --source vba_source/` |
| Need to understand call relationships | `xlvba graph --format json` |
| Need to read a cell value | `xlvba dump --sheets Sheet1 --json` |
| Need to set a cell value before running | `xlvba modify --cell A1 --value 42` |
| COM session hangs or crashes | `taskkill /f /im EXCEL.EXE` then retry |
| Need to rollback after a failed change | `xlvba snapshot restore latest` |
| Need help on agent integration / templates | `xlvba agents` or `xlvba --agents` |

## Python API

```python
from xlvbatools.core.session import ExcelSession
from xlvbatools.analysis.preflight import lint_files
from xlvbatools.vba.search import search_vba
from xlvbatools.snapshot.manager import SnapshotManager
```

## Rules

1. **Always snapshot before risky changes**
2. **Run lint after every VBA edit** before injecting
3. **Never use MsgBox in VBA** -- use Debug.Print or logging
4. **All Dim statements at top** of Sub/Function blocks
5. **Use explicit types** -- `Dim x As Double`, never `Dim x`
6. **Guard interactive code** with `If Not Application.UserControl Then`
7. **Use `session.vb_project`** instead of `session.wb.VBProject` directly to benefit from Trust Center access validation.
8. **Ignore lock files** -- always filter out Excel's temporary owner lock files starting with `~$` when scanning workbooks.
9. **Targeted session closure** -- clean up Excel sessions using targeted graceful closure (`ExcelSession` with `kill_on_enter=True` usesROT/Hwnd tracking) to protect unrelated user workbooks.
10. **Path Log Normalization** -- format all file paths in logs with forward slashes (`/`) for cross-platform consistency.
