# xlvbatools VBA Debug Workflow

Use this workflow to debug VBA runtime errors and hangs.

## Steps

1. **Check dialog events** from the last run:
   ```bash
   xlvba run <MacroName> --json
   ```
   Look for `dialog_events` in the output.

2. **Classify the error**:
   - `compile_error`: Syntax or missing reference -- check module and line
   - `runtime_error`: Logic error -- check Err.Number and Description
   - `msgbox`: Unguarded MsgBox -- replace with Debug.Print
   - `file_dialog`: Missing UserControl guard

3. **Search for the issue**:
   ```bash
   xlvba search "MsgBox" --source vba_source/
   xlvba search "ActiveSheet"
   ```

4. **Fix the code** following the VBA Edit workflow.

5. **If stuck, debug visually**:
   ```bash
   xlvba debug
   ```
   This opens Excel and VBE visibly so you can set breakpoints.

## Common Fixes

- **MsgBox in headless mode**: Replace with `Debug.Print` or remove
- **ActiveSheet reference**: Replace with explicit `ws.Range(...)` reference
- **Dim after executable code**: Move all `Dim` statements to top of procedure
- **Missing On Error handler**: Add `On Error GoTo ErrHandler` at procedure start
