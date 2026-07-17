---
description: Install, configure, discover, and invoke xlvbatools safely from PowerShell or Python.
---

# xlvbatools Get Started Workflow

1. Create or use the consumer repository's `.venv`, then install an exact
   approved release or full reviewed Git revision:

   ```powershell
   py -3.12 -m venv .venv
   & .\.venv\Scripts\python.exe -m pip install "xlvbatools==X.Y.Z"
   $xlvba = ".\.venv\Scripts\xlvba.exe"
   & $xlvba version --text
   ```

2. Initialize configuration and install packaged guidance when needed:

   ```powershell
   & $xlvba init --workbook "workbook/Model.xlsm" --agents --text
   ```

   Installing `.agents/` alone does not install xlvbatools or create
   `xlvbatools.toml`. Existing guidance is preserved unless `--force` is
   explicitly requested.

3. Discover commands before constructing them:

   ```powershell
   & $xlvba help
   & $xlvba help dump
   & $xlvba dump --help
   ```

   The first two commands return JSON for agents. `--help` is presentation
   help. Put flags after the command they configure.

4. Use explicit flags for scope and safety:

   ```powershell
   & $xlvba lint --source "workbook/vba_source"
   & $xlvba inject --source "workbook/vba_source" --dry-run --timeout 120
   & $xlvba dump --sheets "Input" --screenshot --range "A1:K100" --timeout 90
   & $xlvba run "OnCalculate" --workbook "workbook/Model.xlsm" --timeout 120
   ```

   Default stdout is one JSON result envelope. Use `--text` or `--table` only
   for requested presentation output. Never add `--include-hidden-sheets` or
   `--no-backup` without an explicit need and appropriate rollback.

5. In Python, import only from the public package root and verify both outcome
   and Excel cleanup:

   ```python
   from xlvbatools import Project

   project = Project.from_config()
   result = project.run("OnCalculate", timeout=120, save=False)
   macro = result.require_success()
   result.require_clean_shutdown()
   print(macro.run_id)
   ```

   Use `Project.open(workbook, source=...)` for explicit paths. Use
   `Project.run()` rather than private APIs when named-range inputs, save
   behavior, or visibility controls are required.

6. Continue with `vba-edit.md` for changes, `vba-debug.md` for failures, and
   `.agents/skills/xlvba-toolchain/SKILL.md` for the complete operating
   contract.
