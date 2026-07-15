---
description: Diagnose VBA failures, timeouts, and dialogs through the v1 result contract.
---

# xlvbatools VBA Debug Workflow

1. Run the macro with a bounded deadline:

   ```powershell
   xlvba run <MacroName> --timeout 120
   ```

2. Inspect the result envelope:

   - `success`, `phase`, and `error`;
   - `diagnostics.dialog_events`;
   - `diagnostics.cleanup`;
   - the `request_id` used to correlate logs.

3. Classify the failure:

   - `compile_error`: fix the reported module, line, and column;
   - `runtime_error`: inspect the VBA error number and description;
   - `msgbox` or `file_dialog`: add an `Application.UserControl` guard;
   - `named_range_setup`: fix the missing or invalid input before retrying;
   - `timeout`: inspect cleanup and the owned PID; do not terminate unrelated
     Excel processes.

4. Search and lint before reinjection:

   ```powershell
   xlvba search "MsgBox" --source vba_source
   xlvba search "ActiveSheet" --source vba_source
   xlvba lint --source vba_source
   ```

5. Inject, diff, and rerun. If headless evidence is insufficient, use
   `xlvba debug` explicitly; it is the only workflow intended to expose Excel
   and the VBE interactively.

Never run image-wide Excel termination during diagnosis.
