---
description: Safely modify, lint, inject, diff, and verify VBA through xlvbatools v1.
---

# xlvbatools VBA Edit Workflow

1. Create a checkpoint:

   ```powershell
   xlvba snapshot create --desc "before: <change>"
   ```

2. Extract the current workbook project:

   ```powershell
   xlvba extract
   ```

3. Edit files under `vba_source/`:

   - standard modules: `modules/*.bas`;
   - class modules: `classes/*.cls`;
   - document modules: `sheets/*.cls`;
   - forms: `forms/*.frm` and associated binary assets.

4. Lint source and resolve every ERROR:

   ```powershell
   xlvba lint --source vba_source --json
   ```

5. Inject and confirm round-trip equality:

   ```powershell
   xlvba inject --json
   xlvba diff --summary
   ```

6. Run the relevant macro and inspect both the outcome and cleanup:

   ```powershell
   xlvba run <MacroName> --timeout 120 --json
   ```

7. Commit the workbook and matching source together if verification passes.
   Otherwise restore the checkpoint with `xlvba snapshot restore latest`.

Document modules are updated in place because the VBE cannot remove and
re-import them. Use `--no-backup` only when a verified snapshot already
exists. Never use global Excel termination to recover a failed run.
