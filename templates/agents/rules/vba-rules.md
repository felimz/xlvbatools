---
trigger: model_decision
description: "VBA coding, headless execution, COM safety, and encoding rules."
---

# VBA Development Rules

## VBA code

1. Put procedure-level `Dim` statements before executable statements.
2. Use explicit types; avoid implicit `Variant` declarations.
3. Do not use an unguarded `MsgBox` or `FileDialog` in headless paths.
4. Guard interactive behavior with `Application.UserControl`.
5. Use structured error handlers and re-raise failures; do not hide them with
   unbounded `On Error Resume Next`.
6. Prefer explicit workbook, worksheet, range, and object references over
   `ActiveSheet`, `Selection`, or `.Select`.

## Automation safety

1. Use the `xlvba` CLI or `xlvbatools.Project`; raw COM sessions are private.
2. Never run `taskkill /im EXCEL.EXE` or otherwise terminate Excel globally.
3. Inspect `diagnostics.dialog_events` when an operation fails.
4. Require both operation success and clean owned-process shutdown.
5. Let the isolated worker enforce deadlines; do not wrap a blocking COM call
   in a Python thread.

## Edit and verify

1. `xlvba snapshot create --desc "before change"`
2. `xlvba extract`
3. Edit files under `vba_source/`.
4. `xlvba lint`
5. `xlvba inject`
6. `xlvba diff` and confirm no unintended differences.
7. `xlvba run <MacroName> --json` and inspect the complete result envelope.
8. Restore the snapshot if verification fails.

## Encoding

- VBE standard and class modules use the system ANSI code page.
- Keep VBA source, comments, and string literals representable in that code
  page; prefer ASCII when portability is required.
- Let xlvbatools perform UTF-8/source-to-VBE conversion during extraction and
  injection.
