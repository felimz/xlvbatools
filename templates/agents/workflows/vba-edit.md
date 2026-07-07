# xlvbatools VBA Edit Workflow

Use this workflow whenever modifying VBA code in an Excel workbook.

## Steps

1. **Checkpoint** current state:
   ```bash
   xlvba snapshot create --desc "before: <description of change>"
   ```

2. **Extract** VBA from workbook:
   ```bash
   xlvba extract
   ```

3. **Edit** the VBA source files:
   - Standard modules: `vba_source/modules/*.bas`
   - Class modules: `vba_source/classes/*.cls`
   - Sheet code: `vba_source/sheets/*.cls`

4. **Lint** the changes:
   ```bash
   xlvba lint
   ```
   Fix any ERROR-severity issues before proceeding.

5. **Inject** changes back into the workbook:
   ```bash
   xlvba inject
   ```

6. **Verify** by running the target macro:
   ```bash
   xlvba run <MacroName> --json
   ```

7. **Commit or Rollback**:
   - If PASS: `git add vba_source/ && git commit -m "description"`
   - If FAIL: `xlvba snapshot restore latest`

## Notes

- Sheet code-behinds (Document modules) cannot be removed/re-imported.
  They are updated by clearing all lines and inserting new content.
- The `--no-backup` flag skips creating a backup before injection.
  Only use this for rapid iteration when you have a snapshot.
