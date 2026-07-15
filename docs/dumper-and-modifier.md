# Workbook inspection and modification

Use `Project.inspect` or `xlvba dump` to read workbook state and render
worksheets. Use `Project.modify` or `xlvba modify` for targeted changes.
Both operations run in an isolated, owned Excel process.

## Inspection

```python
from xlvbatools import Project

project = Project.open("workbook/MyModel.xlsm")

result = project.inspect(
    ["Input", "Results"],
    cell_range="A1:K100",
    include_data=True,
    include_screenshots=True,
    output_dir="artifacts/screenshots",
    output_json="artifacts/workbook.json",
    include_hidden_sheets=False,
    timeout=90,
)
inspection = result.require_success()
result.require_clean_shutdown()
```

The equivalent CLI commands are:

```powershell
xlvba dump --sheets "Input,Results" --data
xlvba dump --sheets Input --screenshot --range A1:K100
```

### Structured cells

Workbook data stores populated cells by address:

```json
{
  "A1": {
    "row": 1,
    "col": 1,
    "value": 42.0,
    "text": "$42.00",
    "formula": null,
    "is_error": false,
    "error_type": null
  }
}
```

The model preserves Excel's formatted text, formulas, error types, shapes,
named ranges, and the inspected range bounds.

### Screenshot rendering

- The workbook is opened read-only with macros, events, and link updates
  disabled.
- Only visible worksheets are rendered by default.
- Hidden and VeryHidden sheets require `include_hidden_sheets=True` or
  `--include-hidden-sheets`.
- Excel renders the requested range; only the resulting bitmap is moved through
  a blank temporary chart workbook. Worksheet VBA is not copied or compiled.
- Data and screenshots requested together share one worker and one owned Excel
  lifecycle.
- Column letters, row numbers, gridlines, cropped bounds, merged cells, and
  non-default row/column dimensions are composited from the inspected range.
- Clipboard-sensitive copy/paste operations use bounded retries.

Screenshot files appear in both
`InspectionOutput.screenshots` and `OperationResult.artifacts`.

## Modification

```python
# Write a value.
project.modify(
    sheet="Input", cell="C3", value=12.5,
).require_success()

# Write a formula.
project.modify(
    sheet="Input", cell="C4", formula="=C3*2",
).require_success()

# Create and remove a named range.
project.modify(
    name="TaxRate", refers_to="=Input!$C$3",
).require_success()
project.modify(
    name="TaxRate", delete_name=True,
).require_success()
```

For each Excel-backed result, also call `require_clean_shutdown()` when clean
teardown is part of acceptance.

`modify` may use one fresh-worker retry only for a recognized transient RPC
disconnect after the first owned Excel process is confirmed stopped. Macro
execution and VBA injection are never automatically retried because their
effects may be non-idempotent.

## Process safety

Inspection and modification never attach to an unrelated desktop Excel
instance and never terminate Excel by image name. The parent tracks its exact
worker; the worker reports its exact Excel PID. Cleanup can affect only those
owned processes.
