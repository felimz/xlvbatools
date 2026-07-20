---
trigger: glob
globs: ["**/*.py", "pyproject.toml"]
description: "Python API, packaging, versioning, isolation, and test rules for xlvbatools consumers."
---

# Python Development Rules

## Public API

- Build application wrappers around `Project` and `OperationResult`.
- Use typed `MacroStep`, `ModifyStep`, and `InspectStep` values with
  `Project.workflow()` when related operations need one live workbook state.
- Do not add automatic worker-start retry around `Project`; the executor owns
  the two-attempt ceiling and exposes its evidence in `diagnostics.attempts`.
- Import supported names from `xlvbatools`; do not import application-facing
  behavior from `core`, `vba`, `macro`, `workbook`, `analysis`, or `snapshot`.
- Never pass COM proxies or worker transport dictionaries into application
  code.
- Use operation-specific outputs (`MacroOutput`, `ExtractionOutput`,
  `InjectionOutput`, `ComponentDiff`, `ModificationOutput`, and
  `WorkflowOutput`) instead of assuming `data` is a transport dictionary.
- Use `SnapshotService` and immutable `SnapshotRecord` values; never import the
  internal snapshot store.
- Inject an `Executor` test double when a wrapper needs offline unit tests.
- Use `require_success()` for operation success and
  `require_clean_shutdown()` when an Excel lifecycle must end gracefully.
- Treat CLI stdout as one JSON result envelope by default. Request `--text` or
  `--table` only for explicit presentation needs.
- Put command flags after the command they configure. Set `--workbook` or
  `--source` explicitly when overriding configuration, and set a bounded
  `--timeout` on Excel-backed operations.
- Preview supported writes with `--dry-run`. Do not use `--no-backup` unless a
  verified snapshot or equivalent rollback already exists.
- In Python, import public types from `xlvbatools` and translate CLI flags to
  the corresponding `Project` keyword arguments. For macros, repeat
  `--named-range NAME=VALUE`, choose `--save` or `--no-save`, and add
  `--visible` only when the isolated owned Excel window is required. The same
  controls are available through `Project.run()`.
- A workflow has one overall timeout, stops after its first failed step,
  defaults to no-save, and is never replayed after `session_start`. Do not
  describe it as a database transaction; use disposable workbook and external
  input copies when rollback-like test behavior is required.

## Packaging and versions

- Use the repository `.venv` for development commands.
- Consumer projects should pin a released package version or an exact full Git
  revision. Do not rely on whichever editable checkout happens to be active.
- When developing xlvbatools itself, build wheels with normal PEP 517 isolation
  and test them in a fresh consumer virtual environment outside the source
  tree.
- Treat package, serialized-result, worker-protocol, and workflow-schema
  versions, plus the lint-baseline schema, as separate contracts; do not infer
  one version from another.
- Declare package data under `[tool.setuptools.package-data]`.

## Process safety and tests

- Never terminate Excel by image name or select an unrelated desktop Excel
  instance. Only an operation-owned PID may be cleaned up.
- Write generated files under `tmp_path` or an OS temporary directory.
- Mark live Excel coverage with `@pytest.mark.com` or `@pytest.mark.e2e`.
- Run offline tests with `.venv\Scripts\python.exe -m pytest -m "not com and not e2e"`.
- Run the repository's configured Ruff and mypy gates before release. In this
  xlvbatools repository, the exact release commands live in
  `docs/release-validation.md`; consumer projects should use their own checked-
  in quality configuration.
- Run the complete suite before release and confirm no owned Excel or worker
  process remains.
