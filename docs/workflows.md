# One-session workflows

Use `Project.workflow()` when related macro, range-write, and inspection steps
must share one live workbook state. xlvbatools validates the complete request,
starts one isolated worker, opens one owned Excel process and workbook, executes
the steps in order, then closes both. Independent workflows never share an
Excel session.

This avoids paying for a separate Excel startup for each operation while
preserving the same process-isolation boundary as every other `Project`
method.

## Python

```python
from xlvbatools import InspectStep, MacroStep, ModifyStep, Project

project = Project.open("workbook/Model.xlsm")

result = project.workflow(
    [
        MacroStep(
            id="retrieve",
            macro="OnRetrieve",
            named_ranges={"FilePath": "model/input.r3d"},
        ),
        ModifyStep(
            id="inputs",
            sheet="Input",
            values={"C102:C104": [[0.1], [0.0], [-0.1]]},
            calculate=False,
        ),
        MacroStep(id="calculate", macro="OnCalculate"),
        InspectStep(
            id="results",
            sheets=("Input",),
            cell_range="A1:K120",
            include_data=True,
            include_screenshots=False,
            include_rich_text=True,
        ),
    ],
    timeout=240,
    visible=False,
    save=False,
)

workflow = result.require_success()
result.require_clean_shutdown()
inspection = workflow.step("results").data
```

`MacroStep` sets its named ranges immediately before invoking that macro.
`ModifyStep` accepts one or more A1 assignments on one worksheet. Scalars are
valid for single cells; multi-cell ranges require a rectangular two-dimensional
sequence whose shape exactly matches the target. Calculation is off by default
so a later macro can control when the workbook recalculates.

`InspectStep` reads the current in-memory workbook state and can create data,
screenshot, JSON, or Markdown outputs without opening another workbook.
Visible worksheets are the default screenshot scope; set
`include_hidden_sheets=True` only when hidden content is intentionally needed.
Partial cell formatting is opt-in with `include_rich_text=True`; it uses the
same bounded run model as `Project.inspect()` without opening another session.

## CLI

The CLI reads the same versioned typed request from a JSON file:

```json
{
  "workflow_schema_version": "1.0",
  "steps": [
    {
      "id": "retrieve",
      "kind": "macro",
      "macro": "OnRetrieve",
      "named_ranges": {"FilePath": "model/input.r3d"}
    },
    {
      "id": "inputs",
      "kind": "modify",
      "sheet": "Input",
      "values": {"C102:C104": [[0.1], [0.0], [-0.1]]},
      "calculate": false
    },
    {"id": "calculate", "kind": "macro", "macro": "OnCalculate"},
    {
      "id": "results",
      "kind": "inspect",
      "sheets": ["Input"],
      "cell_range": "A1:K120",
      "include_data": true,
      "include_screenshots": false,
      "include_rich_text": true
    }
  ]
}
```

```powershell
& .\.venv\Scripts\xlvba.exe workflow `
  --workbook "workbook/Model.xlsm" `
  --file "workflow.json" `
  --no-save `
  --timeout 240

# Standard input is useful for generated requests.
Get-Content -Raw "workflow.json" |
  & .\.venv\Scripts\xlvba.exe workflow --file - --timeout 240
```

The default stdout is one machine-first JSON envelope. Use
`--output-format text` or `--output-format table` only for a presentation.
`xlvba help workflow` exposes the supported step fields and current workflow
schema to agents without requiring documentation scraping.

`--no-save` is the workflow default. `--save` writes the workbook exactly once,
after every step succeeds. `--visible` deliberately shows the isolated owned
Excel instance; it never selects an existing desktop session.

## Result and failure contract

`WorkflowOutput.steps` preserves request order. Each `WorkflowStepResult`
contains the step ID, kind, status, phase, elapsed time, typed data, error,
dialog evidence, and generated screenshot artifacts. Use `workflow.step(id)`
or the immutable `workflow.by_id` mapping for lookup.

If a step fails, later steps are marked `not_run`, explicit saving is skipped,
and the outer result is unsuccessful. `failed_step_id` identifies the first
failure. The parent enforces one overall timeout for the complete workflow.
The last durable step and phase are retained in `diagnostics.progress` when a
timeout or abrupt worker exit prevents a normal result.

The executor may retry a proven worker failure only before the durable
`session_start` boundary. It never replays a workflow after that boundary, so
a failed calculation cannot repeat an earlier retrieval or external file
write.

## Saving is not a transaction

A workflow does not provide database rollback. `save=False` prevents a final
workbook save, but VBA and Excel can still write external files or call other
systems. For high-fidelity tests, copy the workbook and every external input
into a temporary directory, run the workflow against those disposable files,
and retain or promote outputs only after operation success and clean shutdown.
