# xlvbatools Agent Integration Guide

This document describes how AI coding agents (Antigravity, Cursor, Copilot, etc.)
can use `xlvbatools` to programmatically develop, debug, and maintain VBA code
in Excel workbooks.

---

## Quick Start for Agents

### 1. Initialize a Project

```bash
xlvba init --workbook my_project.xlsm
```

This creates `xlvbatools.toml` and the `vba_source/` directory structure.

### 2. The Edit-Verify Cycle

```bash
# 1. Create a safety checkpoint
xlvba snapshot create --desc "before changes"

# 2. Extract VBA from workbook to editable files
xlvba extract

# 3. Edit the .bas/.cls files (agents edit these directly)
# ... modify vba_source/modules/modMain.bas ...

# 4. Run static analysis
xlvba lint

# 5. Inject changes back into workbook
xlvba inject

# 6. Run a macro to verify
xlvba run OnCalculate --json

# 7. If something breaks, rollback
xlvba snapshot restore latest
```

### 3. Search and Understand Code

```bash
# Find all MsgBox calls
xlvba search MsgBox

# Find patterns with regex
xlvba search "Dim\s+\w+$" --regex

# Generate a call dependency graph
xlvba graph --format mermaid
xlvba graph --format json --output graph.json
```

### 4. Inspect Workbook State

```bash
# Dump sheet values to JSON
xlvba dump --sheets "Sheet1,Results" --json

# Modify a cell value
xlvba modify --sheet Sheet1 --cell C30 --value 42

# Set a formula
xlvba modify --sheet Sheet1 --cell D5 --formula "=C5*2"

# Manage named ranges
xlvba modify --name MyRange --refers-to "=Sheet1!$A$1:$B$10"
```

---

## Agent-Specific Patterns

### Error Recovery

When `xlvba run` reports a dialog event, agents should:

1. **compile_error**: Read the error module/line, fix the VBA source, re-inject
2. **runtime_error**: Check `Err.Number` and fix the logic
3. **file_dialog / msgbox**: Add a `UserControl` guard in the VBA code:
   ```vba
   If Not Application.UserControl Then
       ' Skip interactive dialogs in headless mode
       Exit Sub
   End If
   ```

### Safe Macro Execution (Python API)

```python
from xlvbatools.core.session import ExcelSession

with ExcelSession("workbook.xlsm") as session:
    result = session.run_macro("MyMacro", timeout=60)
    if not result["success"]:
        print(f"Error: {result['error']}")
        for event in result["dialog_events"]:
            print(f"  [{event['type']}] {event['text']}")
```

### Programmatic Lint

```python
from xlvbatools.analysis.preflight import lint_files, print_report

issues = lint_files("vba_source/", disabled_rules=["PF001"])
print(print_report(issues))
```

### Snapshot Management

```python
from xlvbatools.snapshot.manager import SnapshotManager

mgr = SnapshotManager("workbook.xlsm", "vba_source/", "snapshots/")
sid = mgr.create(description="pre-refactor", milestone=True)
# ... make changes ...
mgr.restore(sid)  # rollback if needed
```

---

## Configuration

Create `xlvbatools.toml` in your project root:

```toml
[xlvbatools]
workbook = "workbook/MyProject.xlsm"
vba_source = "workbook/vba_source"
snapshots_dir = "snapshots"
log_dir = "logs"

[xlvbatools.snapshots]
rolling_limit = 10

[xlvbatools.backups]
limit = 5

[xlvbatools.lint]
disabled_rules = ["PF001", "PF003"]
protected_sheets = ["Sheet1"]
```

---

## Dialog Protection Architecture

All Excel COM operations go through `ExcelSession`, which runs a `DialogWatchdog`
background thread that automatically captures and dismisses pop-up dialogs:

1. **VBA Error Handlers** (`On Error GoTo ErrHandler`) -- errors become COM exceptions
2. **VBA UserControl Guards** -- interactive code is skipped in headless mode
3. **Python DialogWatchdog** -- catches compile errors and anything that bypasses layers 1-2

---

## Encoding Rules

- VBE standard/class modules are ANSI only (`windows-1252` on Western systems)
- Avoid Unicode characters in VBA code (they become `?` on import)
- `xlvbatools` handles UTF-8 <-> ANSI conversion automatically during extract/inject

---

## CLI Reference

| Command | Purpose |
|:---|:---|
| `xlvba init` | Initialize project with xlvbatools.toml |
| `xlvba extract` | Extract VBA from workbook |
| `xlvba inject` | Inject VBA into workbook |
| `xlvba diff` | Compare workbook vs source |
| `xlvba lint` | Static analysis (7 rules) |
| `xlvba run <macro>` | Execute macro with dialog protection |
| `xlvba snapshot` | Checkpoint/rollback system |
| `xlvba dump` | Export sheet data and screenshots |
| `xlvba modify` | Edit cells, formulas, named ranges |
| `xlvba debug` | Open Excel + VBE visibly |
| `xlvba search` | Search VBA source files |
| `xlvba fmt` | Format VBA code |
| `xlvba graph` | Generate call dependency graph |
