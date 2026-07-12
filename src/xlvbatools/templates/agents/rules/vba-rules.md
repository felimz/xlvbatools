---
trigger: model_decision
description: "VBA coding standards, COM automation safety rules, and encoding constraints for Visual Basic for Applications (VBA) code and Excel automation."
---

# VBA Development Rules

## VBA Coding Standards

1. **All Dim statements at top**: VBA does not support Dim inside For loops or after executable statements.
2. **Explicit types only**: `Dim x As Double`, never `Dim x` (creates Variant).
3. **No MsgBox in production code**: Use `Debug.Print` or structured error handling.
4. **Guard interactive code**: Wrap FileDialog/MsgBox with `If Not Application.UserControl Then`.
5. **Error handlers required**: Use `On Error GoTo ErrHandler` + `Err.Raise` (never MsgBox).

## COM Automation Rules

1. **Clean up only session-owned Excel** through `ExcelSession`; never terminate every `EXCEL.EXE` process.
2. **Use ExcelSession** from `xlvbatools.core.session` -- never manually call `Dispatch("Excel.Application")`
3. **Check dialog events** after every macro run
4. **Use try/finally** to ensure cleanup

## Development Workflow

After any VBA change:
1. Run `xlvba lint` to check for issues
2. Run `xlvba inject` to push changes to workbook
3. Run `xlvba extract` to keep vba_source/ in sync
4. Run `xlvba run <macro> --json` to verify

Before risky changes:
- `xlvba snapshot create --desc "description"`
- Use `--milestone` for permanent architecture checkpoints

## Encoding Constraints

- VBE standard/class modules are restricted to the system ANSI code page
- Avoid Unicode characters (math symbols, Greek letters, box-drawing) in VBA code
- Use standard ASCII for all code, comments, and strings
