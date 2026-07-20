# Changelog

All notable changes are documented here. This project follows
[Semantic Versioning](https://semver.org/) and the
[Keep a Changelog](https://keepachangelog.com/) format.

## [Unreleased]

## [1.2.2] - 2026-07-20

### Fixed

- XL-16: compiler dialogs are now considered dismissed only after their Win32
  window is confirmed hidden or destroyed. Busy VBE dialogs receive a bounded
  queued-click fallback, reused dialog handles remain eligible for dismissal,
  and the PID-scoped watchdog stays active through the complete graceful-exit
  window. Canonical invalid-workbook lint therefore preserves compiler
  evidence while requiring clean Excel shutdown instead of exact-PID force
  termination.

## [1.2.1] - 2026-07-20

### Fixed

- XL-15: live lint now invokes the VBE's actual whole-project Compile command
  as the typed command-bar button (`Type=1`, `ID=578`) instead of treating a
  temporary no-op macro as compile proof. The operation activates and verifies
  the requested workbook's exact VBProject, keeps the VBE hidden, preserves
  module/line/dialog evidence for real failures, and fails closed when the
  compile state remains ambiguous.
- VBE and dialog-watchdog COM references are released without nested proxy
  leaks, and dialog protection now remains active through workbook close and
  Excel quit so valid and invalid compile operations shut down cleanly.

## [1.2.0] - 2026-07-19

### Added

- `DV001` static analysis for case-insensitive duplicate parameter, local,
  constant, and module declarations that prevent VBA compilation.
- Native screenshot content metrics and structured
  `render_content_mismatch` diagnostics when populated workbook ranges
  repeatedly produce implausibly blank bitmaps.
- VBA-token-aware workbook/source comparison, with case- and spacing-only code
  differences reported as `equivalent` while string and comment text remain
  exact. Raw text comparison remains available explicitly.
- Opt-in, bounded partial rich-text font runs in worksheet cell dumps and
  one-session inspection steps.
- Repeatable lint severity and rule selection, versioned line-stable baselines,
  atomic baseline creation, and multiset-aware new-finding output for both
  source and live-workbook analysis.

### Changed

- The default pytest command is now a fast offline gate. Live Excel,
  sequential stress, distribution, and explicitly supplied downstream
  workbooks have separate enforced tiers and a seekable-output test runner.
- Live test workbooks are synthetic, built in an isolated child process, and
  copied per test; the library suite no longer scans local workbook folders or
  depends on a consumer project.
- Multiple snapshots created in one second receive deterministic numeric
  suffixes instead of blocking until the wall clock advances.
- Repository ignore rules now cover generated Python, packaging, test, Excel,
  snapshot, log, crash, and local-secret state while preserving reviewed test
  fixtures and both maintained agent-guidance trees.
- Extraction, injection, differencing, inspection, modification, component
  listing, and live lint now suppress workbook events before `Workbooks.Open`.
  Non-executing operations also force-disable macros; live compile enables only
  its explicit post-open probe while events remain suppressed.
- Noninteractive sessions continuously hide the owned VBE main window. The VBE
  may remain visible only in the explicit interactive debugger.
- Live compile tests fail closed when VBE compilation cannot be verified.
- Screenshot capture forces a repaint in the owned Excel process, retries the
  complete copy/export transaction, validates native pixels before adding
  headers or gridlines, and restores the caller's visibility and
  `ScreenUpdating` state.
- `Project.diff()` and `xlvba diff` now use VBA-aware comparison by default;
  use `comparison="text"` or `--comparison text` for raw line differencing.
- The additive result schema is `1.3`, the private worker protocol is `2.2`,
  and the lint-baseline schema is independently versioned at `1.0`.
- Version diagnostics distinguish the authoritative imported-code version from
  potentially stale editable-install metadata and flag any mismatch directly.

## [1.1.0] - 2026-07-18

### Added

- Typed, versioned one-session workflows through `Project.workflow()` and
  `xlvba workflow`, with ordered macro, range-write, and inspection steps,
  fail-fast results, explicit save-on-success, durable step progress, and one
  overall timeout.
- Typed `AttemptDiagnostic` and `WorkerExitReport` evidence for every isolated
  executor attempt, including the retry reason and explicit worker reaping.
- Repeatable typed `xlvba run --named-range NAME=VALUE` inputs, explicit
  `--save`/`--no-save` behavior, and opt-in isolated Excel visibility through
  `--visible`.
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

- `IsolatedExecutor` now owns a single shared two-attempt ceiling and one
  overall timeout budget. It automatically retries only a proven failure to
  create a worker or a cleanly reaped worker exit before `session_start`;
  post-ownership, timeout, dialog, protocol, VBA, and ambiguous-cleanup
  failures are never replayed.
- Worker progress now publishes `session_start` durably before any session
  construction, and worker-process exit evidence is separate from Excel
  cleanup evidence.
- The additive result schema is `1.2`, the private worker protocol is `2.1`,
  and the new workflow request schema is independently versioned at `1.0`.
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
- Distribution metadata uses an SPDX license expression and ships the MIT
  license text in source and wheel artifacts.

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
