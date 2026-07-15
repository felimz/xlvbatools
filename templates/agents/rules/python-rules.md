---
trigger: glob
globs: ["**/*.py", "pyproject.toml"]
description: "Python API, packaging, versioning, isolation, and test rules for xlvbatools consumers."
---

# Python Development Rules

## Public API

- Build application wrappers around `Project` and `OperationResult`.
- Import supported names from `xlvbatools`; do not import application-facing
  behavior from `core`, `vba`, `macro`, `workbook`, `analysis`, or `snapshot`.
- Never pass COM proxies or worker transport dictionaries into application
  code.
- Inject an `Executor` test double when a wrapper needs offline unit tests.
- Use `require_success()` for operation success and
  `require_clean_shutdown()` when an Excel lifecycle must end gracefully.

## Packaging and versions

- Use the repository `.venv` for development commands.
- Build wheels with normal PEP 517 isolation and test them in a fresh consumer
  virtual environment outside the source tree.
- Treat package, serialized-result, and worker-protocol versions as separate
  contracts. Change them according to `docs/versioning.md`.
- Declare package data under `[tool.setuptools.package-data]`.

## Process safety and tests

- Never terminate Excel by image name or select an unrelated desktop Excel
  instance. Only an operation-owned PID may be cleaned up.
- Write generated files under `tmp_path` or an OS temporary directory.
- Mark live Excel coverage with `@pytest.mark.com` or `@pytest.mark.e2e`.
- Run offline tests with `.venv\Scripts\python.exe -m pytest -m "not com and not e2e"`.
- Run the complete suite before release and confirm no owned Excel or worker
  process remains.
