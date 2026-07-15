# xlvbatools v1 - Agent Guide

## Project contract

`xlvbatools` provides isolated, headless Excel/VBA automation. Application
code uses `xlvbatools.Project`; COM sessions, worker transport dictionaries,
and implementation subpackages are private.

## Required practices

- Run project commands with the repository `.venv` when one exists.
- Use `Project` or the `xlvba` CLI for Excel-backed operations.
- Check both `OperationResult.require_success()` and, for Excel operations,
  `OperationResult.require_clean_shutdown()`.
- Never enumerate or terminate Excel globally. Cleanup may target only the
  worker and Excel PID owned by the current operation.
- Keep hidden worksheets excluded from screenshots unless the task explicitly
  requests them.
- Use `tmp_path` or operating-system temporary directories for generated test
  artifacts; do not pollute the workspace root.
- Treat names in `xlvbatools.__all__` as the only supported Python API.
- Check `xlvba version --json` when reproducibility matters. Package, result-
  schema, and worker-protocol versions are independent contracts.

Task-specific Python, VBA, skill, and workflow guidance lives under
`.agents/`. Run `xlvba agents` to print the integration guide.
