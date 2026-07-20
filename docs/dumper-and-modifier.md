# Workbook inspection and modification

Use `Project.inspect` or `xlvba dump` to read workbook state and render
worksheets. Use `Project.modify` or `xlvba modify` for targeted changes.
Both operations run in an isolated, owned Excel process.
When reads and writes belong to one ordered calculation pipeline, typed
`ModifyStep` and `InspectStep` values can instead run inside
`Project.workflow()` without reopening Excel. See
[One-session workflows](workflows.md).

## Inspection

```python
from xlvbatools import Project

project = Project.open("workbook/MyModel.xlsm")

result = project.inspect(
    ["Input", "Results"],
    cell_range="A1:K100",
    include_data=True,
    include_screenshots=True,
    include_rich_text=True,
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
xlvba dump --sheets Input --data --rich-text --range A1:K100
xlvba dump --sheets Input --screenshot --range A1:K100
```

### Structured cells

Workbook data stores populated cells by address:

```json
{
  "A1": {
    "row": 1,
    "col": 1,
    "value": "BoldPlain",
    "text": "BoldPlain",
    "formula": null,
    "is_error": false,
    "error_type": null,
    "rich_text": {
      "status": "complete",
      "text_length": 9,
      "characters_inspected": 9,
      "truncated": false,
      "runs": [
        {
          "start": 1,
          "length": 4,
          "text": "Bold",
          "font": {"name": "Aptos", "size": 11.0, "bold": true}
        }
      ]
    }
  }
}
```

The model preserves Excel's formatted text, formulas, error types, shapes,
named ranges, and the inspected range bounds.

Partial rich text is opt-in through `include_rich_text=True` or `--rich-text`.
Each run uses Excel's 1-based character position and includes the font name,
size, bold, italic, underline, strikeout, subscript, superscript, color, and
color index. Collection is bounded to 4,096 characters and 256 runs per cell.
A cell reports `truncated` at either limit and `unsupported` if Excel cannot
expose its Characters model; one unsupported cell does not fail the dump.
This option implies `--data` in the CLI and requires `include_data=True` in
Python.

### Screenshot rendering

- The workbook is opened read-only with macros, events, and link updates
  disabled.
- Only visible worksheets are rendered by default.
- Hidden and VeryHidden sheets require `include_hidden_sheets=True` or
  `--include-hidden-sheets`.
- Excel renders the requested range; only the resulting bitmap is moved through
  a blank temporary chart workbook. Worksheet VBA is not copied or compiled.
- Data and screenshots requested together share one worker and one owned Excel
  lifecycle. Structured data is collected first and supplies independent
  evidence that the rendered range contains visible content.
- Column letters, row numbers, gridlines, cropped bounds, merged cells, and
  non-default row/column dimensions are composited from the inspected range.
- Clipboard-sensitive copy/paste operations use bounded retries.
- Before each capture, xlvbatools temporarily enables `ScreenUpdating`, makes
  only its owned Excel window renderable (off-screen for headless calls), and
  flushes the window paint queue. It restores the caller's original
  `ScreenUpdating` and visibility state afterward.
- The unmodified Excel bitmap is measured before headers and gridlines are
  composited. If a populated range repeatedly yields an implausibly blank
  bitmap, the operation fails with `error.code == "render_content_mismatch"`
  and per-attempt pixel metrics instead of publishing a misleading artifact.
  Set `continue_on_render_error=True` in Python only when partial inspection
  output is explicitly acceptable; the CLI remains fail-closed.

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

`modify` may use the executor's one remaining attempt for a recognized
transient RPC disconnect after the first owned Excel process is confirmed
stopped. This shares the global two-attempt ceiling with safe pre-ownership
startup recovery. Macro execution and VBA injection failures are never
automatically replayed because their effects may be non-idempotent.

## Process safety

Inspection and modification never attach to an unrelated desktop Excel
instance and never terminate Excel by image name. The parent tracks its exact
worker; the worker reports its exact Excel PID. Cleanup can affect only those
owned processes.
