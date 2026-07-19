---
description: Run ordered macros, range writes, and inspection in one isolated Excel session.
---

# One-Session Excel Workflow

Use this workflow when later steps depend on workbook state produced by earlier
steps. Independent tasks should remain independent `Project` operations.

## Python

```python
from xlvbatools import InspectStep, MacroStep, ModifyStep, Project

project = Project.from_config()
result = project.workflow(
    [
        MacroStep("retrieve", "OnRetrieve", {"FilePath": "model/input.r3d"}),
        ModifyStep("inputs", "Input", {"C102:C104": [[0.1], [0.0], [-0.1]]}),
        MacroStep("calculate", "OnCalculate"),
        InspectStep("results", ("Input",), include_screenshots=False),
    ],
    timeout=240,
    save=False,
)
workflow = result.require_success()
result.require_clean_shutdown()
inspection = workflow.step("results").data
```

## CLI

Store the typed request in `workflow.json`:

```json
{
  "workflow_schema_version": "1.0",
  "steps": [
    {"id": "retrieve", "kind": "macro", "macro": "OnRetrieve"},
    {
      "id": "inputs",
      "kind": "modify",
      "sheet": "Input",
      "values": {"C102:C104": [[0.1], [0.0], [-0.1]]}
    },
    {"id": "calculate", "kind": "macro", "macro": "OnCalculate"},
    {
      "id": "results",
      "kind": "inspect",
      "sheets": ["Input"],
      "include_data": true,
      "include_screenshots": false
    }
  ]
}
```

```powershell
& .\.venv\Scripts\xlvba.exe help workflow
& .\.venv\Scripts\xlvba.exe workflow --file workflow.json `
  --no-save --timeout 240
```

Parse the default JSON envelope. Use `--text` or `--table` only when a person
requested presentation output.

## Acceptance and safety

- The complete request is validated before Excel starts.
- Steps run in order through one owned Excel process and workbook open.
- The first failed step stops execution; remaining steps are `not_run`.
- `save=False` and `--no-save` are the defaults. Saving occurs exactly once
  after all steps succeed when explicitly requested.
- One timeout covers the whole workflow. Do not add another startup retry.
- Check `diagnostics.progress` to identify the active step after a timeout.
- A workflow is not a transaction. Use disposable copies of the workbook and
  external inputs when VBA can write files or call other systems.
- Hidden worksheet screenshots remain excluded unless explicitly requested.
- Never terminate Excel globally; only xlvbatools may clean up its owned PID.
