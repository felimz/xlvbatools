# Release validation

A release is accepted only after source, wheel, live Excel, process lifecycle,
and real-workbook checks pass from the repository `.venv`.

## Environment record

Before testing, record:

```powershell
.venv\Scripts\python.exe --version
.venv\Scripts\python.exe -m pip show xlvbatools pywin32 pytest
.venv\Scripts\xlvba.exe version --json
```

Excel-backed tests require Windows, desktop Excel, and Trust Center access to
the VBA project object model. Do not substitute an interpreter inherited from
an agent host.

## Validation tiers

```powershell
# Broad offline and packaging coverage.
.venv\Scripts\python.exe -m pytest -m "not com and not e2e"

# Build-isolated wheel installed into a fresh consumer environment.
.venv\Scripts\python.exe -m pytest tests/test_distribution.py -v

# Public Project API through live Excel.
.venv\Scripts\python.exe -m pytest tests/test_project.py -m "com or e2e" -v

# Sequential COM lifecycle in one interpreter and a parent subprocess.
.venv\Scripts\python.exe -m pytest tests/test_session.py -m com -q
.venv\Scripts\python.exe -m pytest tests/test_sequential_com.py -q

# Complete release gate.
.venv\Scripts\python.exe -m pytest
```

The wheel test uses normal PEP 517 isolation, installs the wheel without
editable/source-path leakage into a fresh virtual environment outside the
repository, and verifies the public API plus package, result-schema, and
worker-protocol versions.

## Real-workbook acceptance

Run at least one representative downstream workbook through the installed v1
surface. For WA-OCEAN:

```powershell
.venv\Scripts\xlvba.exe lint --workbook C:\Users\felim\AntigravityProjects\wa_ocean\workbook\WA-OCEAN-AFR.xlsm --json --timeout 240
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

- 237 passing tests in the complete suite;
- four passing live `Project` API tests;
- a build-isolated wheel imported from a clean consumer environment;
- WA-OCEAN live lint with zero errors, zero `FileCount` false positives, and
  zero compile findings;
- zero residual Excel and worker processes.

Re-run these gates for the release commit; do not treat this historical record
as evidence for a later checkout.
