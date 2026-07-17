# Changelog

All notable changes are documented here. This project follows
[Semantic Versioning](https://semver.org/) and the
[Keep a Changelog](https://keepachangelog.com/) format.

## [Unreleased]

### Added

- A copy-ready PowerShell and Python getting-started guide, plus a packaged
  agent onboarding workflow covering installation, configuration, common
  flags, public imports, result handling, and cleanup checks.
- Versioned machine-readable command discovery through `xlvba help [command]`,
  including parser-derived flags, defaults, choices, and nested subcommands,
  plus complete conventional help with examples for every public command.
- Incremental packaged agent-guidance installation for existing repositories
  through `xlvba agents install`, with explicit scoped overwrite support.
- Immutable operation-specific output models and a typed public snapshot
  service with structured snapshot metadata and errors.
- Automated Ruff, mypy, offline coverage, wheel, and manual Excel lifecycle
  validation workflows.

### Changed

- Agent templates are self-contained for downstream repositories and now
  distinguish package installation, `.agents/` installation, and project
  configuration. Documentation now indexes every guide and names every public
  export explicitly.
- The current reliability contract is documented at
  `docs/headless-reliability.md` rather than under a migration-era filename.
- CLI commands now emit one versioned JSON result envelope to stdout by
  default. Text and table presentations require `--text`, `--table`, or the
  corresponding explicit `--output-format` value.
- Graph payload representation is selected independently with
  `--graph-format`; dump file creation is explicit through `--write-json` and
  `--write-markdown`.
- Operation arguments are recursively immutable and only the idempotent
  modification operation may request a transient retry.
- Snapshot metadata uses atomic writes and crash-safe operating-system locks.
- The CLI parser, entry point, and command implementations are separate
  modules.
- Worker failure details retain structured error, timeout, traceback, and log
  evidence in the public result envelope.

### Removed

- The image-wide Excel termination helper and direct public exposure of the
  internal snapshot store.
- A redundant third copy of agent templates; packaged templates are now the
  canonical source mirrored by the active `.agents/` tree.

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
