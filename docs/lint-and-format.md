# Linting and formatting

xlvbatools provides source-only analysis, live-workbook analysis, and a
non-destructive VBA formatter.

## Lint modes

Source lint does not require Excel:

```powershell
xlvba lint --source vba_source
```

Live lint extracts every VBProject component into one in-memory project model,
runs the same rules, and optionally asks Excel for compile evidence:

```powershell
xlvba lint --workbook workbook/MyModel.xlsm --timeout 240
```

```python
from xlvbatools import Project

project = Project.from_config()
source_result = project.lint_source()
workbook_result = project.lint_workbook(compile_test=True, timeout=240)
```

Both adapters share one immutable project symbol index. Public declarations in
standard modules, class members, and duplicate public procedures are resolved
at project scope rather than through per-file allowlists. Excel compilation is
the semantic authority for the live project. Workbook startup code is disabled
before live lint opens the file, the owned VBE is kept hidden, and an
unverifiable compile result fails closed.

Lint returns an unsuccessful `OperationResult` when ERROR-severity findings
exist. Warnings and style findings remain in `data`.

## Select findings

Severity and rule filters are inclusive and repeatable. They work identically
for source and live lint:

```powershell
xlvba lint --source vba_source `
  --severity ERROR --severity WARNING `
  --rule IP001 --rule DV001
```

Unknown severities and rule IDs fail closed instead of silently selecting zero
findings. `xlvba help lint` exposes the accepted values to agents.

```python
result = project.lint_source(
    severities=("ERROR", "WARNING"),
    rules=("IP001", "DV001"),
)
```

Only selected ERROR findings determine the result's lint success. Filtering
cannot convert an Excel startup, worker, protocol, timeout, or cleanup failure
into success. Counts for raw, selected, suppressed, known, and new findings are
retained in `OperationResult.metadata`.

## Baselines and new findings

Create a baseline from all current, unfiltered findings:

```powershell
xlvba lint --source vba_source --write-baseline .xlvba/lint-baseline.json
```

Then report only findings not represented in that baseline:

```powershell
xlvba lint --source vba_source `
  --baseline .xlvba/lint-baseline.json --new-only
```

```python
project.lint_source(write_baseline=".xlvba/lint-baseline.json")
new_result = project.lint_source(
    baseline=".xlvba/lint-baseline.json",
    new_only=True,
)
```

`--new-only` requires `--baseline`. Baseline JSON uses
`LINT_BASELINE_SCHEMA_VERSION`, is sorted deterministically, and is replaced
atomically. A fingerprint includes severity, rule, module, procedure, and
normalized message, but deliberately excludes the line number. It therefore
survives code movement and VBA casing changes. Duplicate fingerprints are
counted, not collapsed: if the baseline has one occurrence and analysis finds
two, the second occurrence is new.

Baseline creation records raw findings before severity, rule, or new-only
selection. This keeps a baseline complete even when the same invocation uses
presentation filters. Review and commit baseline changes like source changes;
do not use them to suppress unexplained analyzer regressions.

## Rules

| ID | Severity | Detects |
|:---|:---|:---|
| DS001 | ERROR | `Dim` after executable code |
| CS001 | ERROR | `Const` after executable code |
| LC001 | WARNING | Invalid or orphaned line continuation |
| SB001 | ERROR | Unbalanced procedure or block terminator |
| PF001 | WARNING | Interactive `MsgBox` in headless paths |
| PF002 | WARNING | Implicit Variant declaration |
| PF003 | WARNING | `ActiveSheet` or `ActiveCell` dependency |
| OE001 | WARNING | Missing `Option Explicit` |
| UV001 | ERROR | Undeclared identifier, resolved with project scope |
| DV001 | ERROR | Duplicate declaration in one VBA scope |
| BK001 | WARNING | Misleading block-level declaration |
| SD002 | STYLE | Multiple declarations on one line |
| PF004 | STYLE | Obsolete `Integer` type |
| SD005 | STYLE | Missing `ByVal` or `ByRef` |
| SD006 | STYLE | Missing procedure access modifier |
| SD010 | STYLE | Line longer than the configured limit |
| SD014 | STYLE | Obsolete `Call` keyword |
| SC001 | WARNING | Hard-coded secret or credential |
| SC002 | WARNING | Non-portable absolute path |
| CL001 | STYLE | Type-declaration suffix |
| CL002 | STYLE | `.Select`, `Selection`, or `ActiveWindow` dependency |
| SF001 | WARNING | Silent error suppression |
| EH001 | WARNING | Public procedure without an error handler |
| FD001 | WARNING | `FileDialog` without a `UserControl` guard |
| DC001 | WARNING | Unused local variable |
| DC002 | WARNING | Empty procedure |
| DC003 | WARNING | Entry-point-aware procedure with no incoming calls |
| SD015 | STYLE | Consecutive blank lines |
| SD016 | STYLE | Double-spaced code block |
| RK001 | WARNING | Reserved keyword used as an identifier |
| IP001 | ERROR | Executable statement outside a procedure |
| DP001 | ERROR | Duplicate public procedure across modules |
| SM001 | WARNING | Unknown class or typed-object member |
| SM002 | WARNING | Ambiguous class-member call requiring instance context |
| CT001 | ERROR | Excel/VBE compile-test failure |

Rules can be disabled through `[xlvbatools.lint].disabled_rules`. Prefer
fixing analyzer context or code defects over suppressing project-wide symbols.

## Formatter

Preview formatting:

```powershell
xlvba fmt --source vba_source --dry-run
```

Apply formatting:

```powershell
xlvba fmt --source vba_source --indent 4
```

The formatter normalizes indentation, preserves VBE attributes and `Option`
statements at column zero, collapses excessive blank lines, and removes
mechanical double-spacing. It does not inject or modify a workbook; lint and
diff again before injection.
