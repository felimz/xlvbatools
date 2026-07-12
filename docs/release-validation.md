# Headless Reliability Release Validation

## Environment

- Interpreter: repository `.venv`
- Python: 3.14.4, 64-bit
- pywin32: 312
- pytest: 9.1.1
- Excel: desktop COM automation on Windows

All authoritative commands use `.venv\Scripts\python.exe -m pytest` rather than an interpreter inherited from the calling agent or shell.

## Repository validation

The release validation tiers are:

```powershell
.venv\Scripts\python.exe -m pytest -m unit
.venv\Scripts\python.exe -m pytest -m integration
.venv\Scripts\python.exe -m pytest tests/test_watchdog.py -v
.venv\Scripts\python.exe -m pytest tests/test_session.py -v
.venv\Scripts\python.exe -m pytest tests/test_integration_samples.py -v
```

The `integration` marker includes live Excel sessions, sample-workbook extraction, modal UI handling, worker timeouts, and PID-isolation cases.

## WA-OCEAN-AFR disposable-copy matrix

Validation used copies of `workbook/WA-OCEAN-AFR.xlsm` and the configured metric RISA model. The dirty project workbook and tracked VBA source were not modified.

| Case | Result |
|---|---|
| Baseline compile | No compile error found; control 578 remained enabled, so `compile_verified` is false with a warning |
| `OnRetrieve` | Success; no dialogs; owned PID exited gracefully |
| `OnCalculate` | Success; no dialogs; owned PID exited gracefully |
| Injected multiline runtime error | Correct `runtime_error`; both lines captured; `&End` clicked; owned PID exited |
| Injected `Option Explicit` compile error | Failure located at injected module, line, and column without a visible VBE window |
| Missing named range | Stopped in `named_range_setup`; macro was not invoked |
| Infinite VBA loop | Parent returned at deadline and terminated only the worker-owned Excel PID |
| Modal `MsgBox` and file picker | Captured and dismissed headlessly |
| Unrelated live Excel instance | Remained open during worker timeout cleanup |
| Extreme CG-position inputs | Calculation completed successfully; this did not produce a deterministic unreachable-target fixture |

The project-specific unreachable-CG error case remains dependent on a known model/input combination that violates the solver's rank or residual acceptance criteria. It should be added when such a fixture is identified; extreme coordinate magnitude alone is not a valid substitute.

## Headless UI verification

Compile tests inspect control 578 but do not execute the VBE command-bar control, because that API can synchronously flash its File menu. Compilation is forced through a temporary unsaved no-op VBA probe, while `VBE.MainWindow.Visible` remains false. Integration coverage asserts that the editor remains hidden for an injected compile failure.

## Process safety

No release-validation path uses image-wide Excel termination. Timeout and exit cleanup use the PID reported by the isolated session. Final validation must end with no session-owned `EXCEL.EXE` process remaining.
