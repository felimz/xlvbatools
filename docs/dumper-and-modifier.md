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

`Project.inspect(..., gridlines="native")` preserves the workbook's appearance
by default. Gridline mode and row/column headings are independent:

| Mode | Screenshot behavior |
|:---|:---|
| `native` (default) | Use Excel's saved/current visibility and gridline color. Fills and merged cells suppress gridlines as they do in Excel. No grid is added. |
| `forced` | Hide Excel's gridlines, then overlay a solid, opaque grid at every native cell boundary, including across fills and merged ranges. Color and pixel width are sampled from Excel's actual automatic UI grid. Explicit cell borders take precedence. |
| `off` | Hide Excel's gridlines and add no overlay. Explicitly formatted cell borders remain visible. |
| `custom` | Overlay every cell boundary using the requested color, opacity, width, and line style. Excel's native grid is hidden to avoid duplicate lines. |

```python
from xlvbatools import GridlineOptions, Project

project = Project.open("workbook/Model.xlsm")
result = project.inspect(
    ["Input"], cell_range="A1:K100", include_screenshots=True,
    gridlines=GridlineOptions(
        mode="custom", color="#808080", opacity=0.3, width=1, style="dashed",
    ),
    include_headers=True,
)
result.require_success()
result.require_clean_shutdown()
```

`color` is a six-digit `#RRGGBB` value, or `None` to use Excel's measured native
automatic grid color. `opacity` is 0 (invisible) through 1
(opaque), blended over the actual cell content and fill. `width` is a positive
integer in output pixels, or `None` to use Excel's measured native stroke width.
`style` is `solid`, `dashed`, or `dotted`. Styling
fields require `mode="custom"`; invalid options fail before Excel starts.
Set `include_headers=False` to omit column letters and row numbers.

```powershell
xlvba dump --sheets Input --screenshot --gridlines off
xlvba dump --sheets Input --screenshot --gridlines forced
xlvba dump --sheets Input --screenshot --gridlines custom `
  --gridline-color "#808080" --gridline-opacity 0.3 `
  --gridline-width 1 --gridline-style dashed --no-headers
```

These controls affect the captured image. Cell fills, merges, and borders are
not edited. Temporary native-gridline overrides are restored after success or
failure, including within a workflow that later saves the workbook.

Worksheet content and heading boxes/text are copied from Excel's own window
pixels at its native DPI and workbook zoom. No Arial substitute, fixed heading
font size, interpolated grid coordinates, or chart frame is added. A blank
temporary workbook mirrors exact row/column dimensions to measure grid edges
even where the source has fills, merges, or shapes. Long ranges are assembled
from unscaled native tiles. Custom strokes land on those measured cell edges;
displayed cell borders retain their native pixels.

- The workbook is opened read-only with macros, events, and link updates
  disabled.
- Only visible worksheets are rendered by default.
- Hidden and VeryHidden sheets require `include_hidden_sheets=True` or
  `--include-hidden-sheets`.
- Only the owned Excel window is captured. The geometry reference is a blank
  workbook; source worksheets and VBA are not copied or compiled.
- Data and screenshots requested together share one worker and one owned Excel
  lifecycle. Structured data is collected first and supplies independent
  evidence that the rendered range contains visible content.
- Excel renders cell formatting, merges, borders, and optional headings.
  Forced/custom grids use measured native pixel edges without rounding drift.
- Native capture uses bounded attempts and does not depend on the clipboard.
- Before each capture, xlvbatools temporarily enables `ScreenUpdating`, makes
  only its owned Excel window renderable (off-screen for headless calls), and
  scrolls each tile into the active viewport before flushing the window paint
  queue. It restores the caller's view, `ScreenUpdating`, and visibility state.
- `InspectionOutput.screenshot_diagnostics` records native DPI/zoom, exact pixel
  boundaries, tile source rectangles, sampled grid color/width, bounded attempt
  history, window state, and native image metrics. This makes rendering failures diagnosable
  without attaching to COM or replaying the operation.
- If native capture fails, the operation fails closed in phase
  `screenshot_capture` with `error.code == "screenshot_capture_failed"` and
  the same structured attempt evidence.
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
