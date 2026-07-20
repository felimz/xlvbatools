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
   & $xlvba lint --source "workbook/vba_source" `
     --write-baseline ".xlvba/lint-baseline.json"
   & $xlvba diff --comparison vba --summary --timeout 120
   & $xlvba inject --source "workbook/vba_source" --dry-run --timeout 120
   & $xlvba dump --sheets "Input" --screenshot --range "A1:K100" --timeout 90
   & $xlvba dump --sheets "Input" --data --rich-text --range "A1:K100" --timeout 90
   & $xlvba run "OnCalculate" --workbook "workbook/Model.xlsm" `
     --named-range "InputValue=42" --no-save --timeout 120
   & $xlvba workflow --workbook "workbook/Model.xlsm" `
     --file "workflow.json" --no-save --timeout 240
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

   Use `Project.open(workbook, source=...)` for explicit paths. CLI callers can
   repeat `--named-range NAME=VALUE`, choose `--save` or `--no-save`, and add
   `--visible` when the isolated owned Excel window is intentionally required.
   Python callers use the corresponding `Project.run()` arguments.
   Use repeatable `--severity`/`--rule` lint filters and
   `--baseline <path> --new-only` for a reviewed regression gate. The default
   VBA diff ignores only identifier/keyword case and insignificant token
   spacing; strings and comments remain exact.

   When related retrieval, range-write, calculation, and inspection steps need
   the same live workbook state, use typed `Project.workflow()` steps or a
   versioned `xlvba workflow --file workflow.json` request. Workflows default
   to no-save, have one overall timeout, stop after the first failed step, and
   are never replayed after `session_start`. Follow `excel-workflow.md` for a
   copy-ready request.

   Worker-start retry is executor-owned and automatic only before Excel
   ownership. Keep `attempt_count` and `diagnostics.attempts` in logs when a
   second attempt occurs; do not wrap the call in another startup retry.

6. Continue with `vba-edit.md` for changes, `vba-debug.md` for failures, and
   `.agents/skills/xlvba-toolchain/SKILL.md` for the complete operating
   contract.
