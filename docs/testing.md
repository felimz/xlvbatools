# Testing

The default xlvbatools test command is intentionally fast and offline:

```powershell
.venv\Scripts\python.exe -m pytest
```

It does not start Excel, build a wheel, run lifecycle stress loops, scan a
workbook directory, or invoke a downstream project's tests. Use the smallest
tier that can prove the change, then expand deliberately before release.

For named tiers, the repository runner is the preferred interface:

```powershell
.venv\Scripts\python.exe scripts/test.py fast
.venv\Scripts\python.exe scripts/test.py excel-smoke
```

The runner invokes the same pytest markers shown below, but captures child
output in a seekable temporary file. This prevents an Excel descendant from
retaining a console pipe and hiding pytest's final exit status. In detached or
externally supervised environments, add `--log-file PATH` to persist the full
transcript and `TEST_RUNNER_EXIT_CODE` even if the terminal disconnects.

## Test tiers

| Tier | Marker | What it proves | Command |
| --- | --- | --- | --- |
| Fast default | `unit` and `integration` | Pure behavior plus short offline subprocess boundaries | `.venv\Scripts\python.exe -m pytest` |
| Unit | `unit` | Isolated in-process behavior | `.venv\Scripts\python.exe -m pytest -m unit` |
| Offline integration | `integration` | Multi-component and worker-protocol behavior without Excel | `.venv\Scripts\python.exe -m pytest -m integration` |
| Excel smoke | `excel and smoke` | Short public-API and worker lifecycle checks | `.venv\Scripts\python.exe scripts/test.py excel-smoke` |
| Excel acceptance | `excel and not stress` | Disposable live-workbook behavior, dialogs, screenshots, and cleanup | `.venv\Scripts\python.exe scripts/test.py excel` |
| Excel stress | `excel and stress` | Sequential native-session replay plus repeated operations in one long-lived parent interpreter | `.venv\Scripts\python.exe scripts/test.py stress` |
| Distribution | `distribution` | PEP 517 wheel install and use in a fresh consumer environment | `.venv\Scripts\python.exe scripts/test.py distribution` |
| External acceptance | `external` | An explicitly supplied downstream workbook | See below |

The `stress` and `smoke` markers are secondary scheduling labels. Every test
must have exactly one primary marker: `unit`, `integration`, `excel`,
`distribution`, or `external`. Collection fails if the taxonomy is missing or
ambiguous. A test using a live Excel fixture must belong to `excel`.

## Downstream project separation

The library suite uses only synthetic workbooks created under pytest's
temporary directory. It never scans `sample_workbooks/`, another repository,
or a consumer project's test tree.

To validate a real consumer workbook, pass each path explicitly:

```powershell
.venv\Scripts\python.exe scripts/test.py external `
  --external-workbook C:\path\to\disposable\Model.xlsm
```

`--external-workbook` may be repeated. This opt-in test extracts the requested
workbook and runs xlvbatools' static linter and dependency graph against the
extracted source. It does not run the consumer project's pytest suite. The
consumer repository remains responsible for its own domain and workbook
acceptance tests against a pinned xlvbatools release.

## Coverage and release progression

Use the default suite for iteration and coverage:

```powershell
.venv\Scripts\python.exe -m pytest `
  --cov=xlvbatools --cov-report=term-missing --cov-fail-under=60
```

Before release, run Ruff and mypy, then the default, distribution, Excel
acceptance, and stress tiers separately. Keeping separate commands makes a
failure's ownership and runtime obvious. Run external acceptance in the
downstream repository or with an explicitly supplied disposable workbook;
never make it an implicit upstream release dependency.

Live tests own only the worker and Excel PIDs they create. A passing live tier
must also contain no native finalizer diagnostics and leave no owned process
behind. Never use global Excel termination as test cleanup.
