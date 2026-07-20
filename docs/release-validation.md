# Release validation

A release is accepted only after source, wheel, live Excel, process lifecycle,
and real-workbook checks pass from the repository `.venv`.

## Environment record

Before testing, record:

```powershell
.venv\Scripts\python.exe --version
.venv\Scripts\python.exe -m pip show xlvbatools pywin32 pytest
.venv\Scripts\xlvba.exe version
```

Excel-backed tests require Windows, desktop Excel, and Trust Center access to
the VBA project object model. Do not substitute an interpreter inherited from
an agent host.

## Validation tiers

```powershell
# Broad offline and packaging coverage.
.venv\Scripts\python.exe -m pytest -m "not com and not e2e"

# Static quality and minimum offline coverage gates.
.venv\Scripts\ruff.exe check src tests
.venv\Scripts\mypy.exe --follow-imports=skip `
  src/xlvbatools/project.py `
  src/xlvbatools/execution.py `
  src/xlvbatools/results.py `
  src/xlvbatools/outputs.py `
  src/xlvbatools/workflow.py `
  src/xlvbatools/core/workflow.py `
  src/xlvbatools/analysis/filtering.py `
  src/xlvbatools/vba/differ.py `
  src/xlvbatools/workbook/dumper.py `
  src/xlvbatools/workbook/modifier.py `
  src/xlvbatools/snapshots.py `
  src/xlvbatools/cli
.venv\Scripts\python.exe -m pytest -m "not com and not e2e" --cov=xlvbatools --cov-fail-under=60

# Build-isolated wheel installed into a fresh consumer environment.
.venv\Scripts\python.exe -m pytest tests/test_distribution.py -v

# Exact release artifacts. The output directory should be empty before use.
.venv\Scripts\python.exe -m build --sdist --wheel --outdir dist

# Public Project API through live Excel.
.venv\Scripts\python.exe -m pytest tests/test_project.py -m "com or e2e" -v
.venv\Scripts\python.exe -m pytest tests/test_workflow_live.py -v

# Sequential COM lifecycle in one interpreter and a parent subprocess.
.venv\Scripts\python.exe -m pytest tests/test_session.py -m com -q
.venv\Scripts\python.exe -m pytest tests/test_sequential_com.py -q

# Complete release gate.
.venv\Scripts\python.exe -m pytest
```

The wheel test uses normal PEP 517 isolation, installs the wheel without
editable/source-path leakage into a fresh virtual environment outside the
repository, and verifies the public API plus package, result-schema, and
worker-protocol and workflow-schema versions. It also invokes the installed
`xlvba help` catalog and installs packaged guidance into a consumer `.agents/`
directory. Before publication, install the exact wheel from `dist` into a
second fresh environment and verify it reports the release version. Publish
that wheel and its matching source distribution without rebuilding them.

## Real-workbook acceptance

Run at least one representative downstream workbook through the installed v1
surface. For WA-OCEAN:

```powershell
.venv\Scripts\xlvba.exe lint --workbook C:\Users\felim\AntigravityProjects\wa_ocean\workbook\WA-OCEAN-AFR.xlsm --timeout 240
```

Acceptance requires:

- process exit code 0;
- no ERROR-severity lint finding;
- no false UV001 finding for cross-module public symbols such as `FileCount`;
- no CT001 compile finding;
- cleanup reports graceful exit without forced worker termination;
- no session-owned Excel or worker remains.

Use disposable copies for injection, deliberate compile/runtime failures,
infinite-loop timeouts, modal dialog fixtures, and destructive macro tests.

## UI and dialog checks

Compile probing must leave the VBE hidden and must not invoke command-bar
control 578, which can flash the VBE File menu. Modal tests must confirm
multiline error capture, bounded dismissal, and PID scope. A separate unrelated
Excel instance must remain open during a worker timeout test.

## Native teardown checks

A zero pytest exit code is insufficient when stdout or stderr contains native
finalizer evidence such as `Windows fatal exception`, `0x800706ba`, or
`0x80010108`. Sequential tests must inspect both streams and confirm child
COM proxies are released while Excel remains alive.

## v1.0.0 validation record

The integrated v1 refactor completed with:

- 257 passing tests in the complete suite;
- four passing live `Project` API tests;
- a build-isolated wheel imported from a clean consumer environment;
- WA-OCEAN live lint with zero errors, zero `FileCount` false positives, and
  zero compile findings across 1,069 reported issues;
- zero residual Excel and worker processes.

Re-run these gates for the release commit; do not treat this historical record
as evidence for a later checkout.

## v1.1.0 validation record

The `v1.1.0` release completed with:

- 293 passing offline tests with 68.15% coverage, plus clean Ruff and mypy
  gates;
- 323 passing tests in the complete suite, including 50 sequential
  `Project.run()` operations and 25 sequential one-session workflows;
- WA-OCEAN live lint and compile with zero errors and zero dialogs across
  1,078 reported issues;
- a build-isolated wheel installed and exercised from a clean consumer
  environment; and
- zero residual Excel or xlvbatools worker processes after live validation.

## v1.2.0 candidate validation record

The unreleased `v1.2.0` candidate completed with:

- clean Ruff, CI-scoped mypy, documentation/template parity, and diff checks;
- 337 passing offline/non-COM tests, including the build-isolated wheel
  contract;
- 372 passing tests in the complete suite in 1,105.38 seconds;
- clean same-interpreter COM tests, five repeated direct macro/formatting
  sessions, 50 sequential `Project.run()` operations, and 25 sequential
  one-session workflows with no native finalizer diagnostics;
- a sentinel `Workbook_Open` that remained unexecuted across extract, diff,
  inject, and live compile lint;
- a duplicate parameter/local declaration reported as both `DV001` static
  evidence and `CT001` Excel compile evidence;
- screenshot capture succeeding after VBA left `ScreenUpdating=False`, then
  restoring that state for the following workflow step; and
- zero residual Excel or xlvbatools worker processes after the complete run.

The 1.2.0 feature pass also live-verified that VBA case/spacing changes are
`equivalent` while raw text reports them as modified, that an analyzer failure
can be baselined and then cleared by `new_only` without hiding lifecycle
failures, and that a disposable workbook exposes two partial font spans through
`Project.inspect(include_rich_text=True)`. The final complete-run audit found
zero Excel processes and zero xlvbatools worker processes before and after the
suite.

This is upstream candidate evidence only. Re-run the WA-OCEAN Section 5
rejected-travel screenshot and broken-startup acceptance cases after installing
the released wheel before closing XL-13 and XL-14 downstream.
