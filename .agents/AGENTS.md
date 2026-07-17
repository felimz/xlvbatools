# xlvbatools v1 - Agent Guide

## Using this template

The Python package, project configuration, and agent guidance are separate:

1. Install a pinned `xlvbatools` release in the repository `.venv`.
2. Run `xlvba agents install` to copy this `.agents/` guidance into an existing
   project, or `xlvba init --agents` to create configuration and guidance
   together.
3. Follow `workflows/get-started.md` to verify the local CLI, configuration,
   common flags, and Python import boundary.
4. Read the task-specific skill, rule, and workflow.
5. Customize workbook paths and acceptance commands for the repository and
   commit those customizations.

Installing `.agents/` files does not install the Python package or create an
`xlvbatools.toml` unless `xlvba init --agents` is used.

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
- Check `xlvba version` when reproducibility matters. Package, result-
  schema, and worker-protocol versions are independent contracts.
- Parse the default JSON envelope. Use `--text` or `--table` only when a human
  presentation is explicitly requested.
- Put flags after the command they configure, use an explicit `--timeout` for
  Excel-backed operations, and use `--dry-run` before injection or formatting
  writes when the command supports it.

Task-specific Python, VBA, skill, and workflow guidance lives under
`.agents/` (plural). In a consumer repository, install these packaged files
with `xlvba agents install`, or use `xlvba init --agents` while initializing a
new project. Installation preserves existing files unless `--force` is
explicitly requested and never deletes project-specific extras. Read this file
first, then the task-specific skill, rule, and workflow. Use `xlvba help` for a
machine-readable command catalog and `xlvba COMMAND --help` for conventional
option help and examples.
