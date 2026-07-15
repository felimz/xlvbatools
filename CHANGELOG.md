# Changelog

All notable changes are documented here. This project follows
[Semantic Versioning](https://semver.org/) and the
[Keep a Changelog](https://keepachangelog.com/) format.

## [1.0.0] - 2026-07-14

### Added

- `Project`, the single high-level API for configured workbook automation.
- Typed `Operation`, `OperationRequest`, `OperationResult`, and `Executor`
  contracts suitable for thin downstream wrappers and dependency injection.
- One project-level VBA symbol index shared by extracted-source and live-
  workbook linting.
- Independent version identifiers for the package, serialized result schema,
  and private worker transport.
- Isolated wheel validation in a fresh consumer virtual environment.

### Changed

- Every Excel-backed `Project` and CLI operation now uses the same typed,
  isolated executor boundary.
- CLI JSON output is consistently the complete `OperationResult` envelope.
- Live-workbook lint now resolves public standard-module symbols across the
  whole VBProject before evaluating module rules.
- Worker protocol advanced to `2.0`.
- Package maturity advanced to `Production/Stable` with `1.0.0` as the first
  supported public contract.
- Documentation, active agent guidance, repository templates, and packaged
  templates now describe the same v1 API, result envelope, owned-PID safety
  rules, hidden-sheet defaults, and release workflow.

### Removed

- Legacy top-level exports including `XlvbaProject`, `ExcelSession`,
  `lint_files`, `lint_workbook`, `run_macro`, and raw workbook helpers.
- `OperationResult.from_legacy` and the legacy dictionary-conversion layer.
- Public compatibility promises for implementation subpackages. Low-level
  COM, VBA, workbook, and macro modules are private worker backends.
