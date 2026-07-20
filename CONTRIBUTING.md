# Contributing to xlvbatools

This guide covers development and release validation for the v1 architecture.

---

## Development Setup

### 1. Prerequisites
* **OS:** Microsoft Windows (Excel COM automation APIs require Windows).
* **Excel:** Microsoft Excel installed locally with VBA macros enabled.
* **Python:** Python 3.10 or higher.

### 2. Environment Setup
Clone the repository and create a virtual environment:

```bash
git clone https://github.com/felimz/xlvbatools.git
cd xlvbatools
python -m venv .venv
.venv\Scripts\activate
```

Install development dependencies:

```bash
.venv\Scripts\python.exe -m pip install -e ".[dev]"
```

---

## Running Tests

Tests are managed using `pytest`. Plain `pytest` is the fast offline suite.
Excel, stress, wheel, and external-workbook tests are explicit opt-ins; see
[docs/testing.md](docs/testing.md) for the enforced taxonomy.

Run the normal iteration suite:
```bash
.venv\Scripts\python.exe -m pytest
```

Run only offline/unit tests:
```bash
.venv\Scripts\python.exe -m pytest -m unit
```

Run tests with coverage:
```bash
.venv\Scripts\python.exe -m pytest --cov=xlvbatools --cov-fail-under=60
```

Run live Excel smoke coverage without the long lifecycle loops:
```bash
.venv\Scripts\python.exe scripts/test.py excel-smoke
```

Use `scripts/test.py` for named live tiers so pytest output is captured through
a seekable file and Excel descendants cannot retain the terminal pipe.

Every test has exactly one primary marker: `unit`, `integration`, `excel`,
`distribution`, or `external`. Use `stress` and `smoke` only as secondary
markers on `excel` tests. The collection hook fails on missing, conflicting,
or live-fixture markers. Library tests must use synthetic disposable workbooks;
downstream workbooks are accepted only through an explicit
`--external-workbook` path and never through directory auto-discovery.

---

## Coding Guidelines

* **Public API:** Add public behavior through `Project`, typed requests/results,
  and `xlvbatools.__all__`. Do not expose worker backends or COM objects.
* **Type Safety:** The project uses PEP 561 types. New public behavior requires
  explicit annotations and typed result data.
* **Workflow Safety:** Validate all workflow steps before Excel starts, keep
  `session_start` as the no-replay boundary, and use one parent-enforced
  timeout for the complete workflow.
* **Platform Safety:** Check platform status before using `win32com.client` or
  native Win32 APIs.
* **COM Safety:** Never scan or terminate Excel globally. Cleanup may target
  only the PIDs owned and reported by the current operation.
* **UI Safety:** Watchdog calls use bounded `SendMessageTimeoutW`, never an
  unbounded cross-process `SendMessageW`.

## Releases

Follow [docs/versioning.md](docs/versioning.md) and
[docs/release-validation.md](docs/release-validation.md). Release wheels must
be built with PEP 517 isolation and installed into a fresh consumer virtual
environment before tagging.

## Documentation and agent templates

When public behavior changes, update `README.md`, the relevant file under
`docs/`, `CHANGELOG.md`, and the active `.agents/` skill in the same change.
The packaged templates under `src/xlvbatools/templates/agents/` are canonical.
The active `.agents/` tree must remain byte-for-byte identical so repository
work and newly initialized consumer projects use the same guidance.
Both trees are intentionally tracked: root `.agents/` is active immediately in
this repository, while `src/xlvbatools/templates/agents/` is package data used
by `xlvba agents install`. The root tree is not a generated local artifact and
must not be added to `.gitignore`.

Validate documented CLI examples against `xlvba <command> --help`. Safety
guidance must never recommend global Excel termination or direct use of
implementation subpackages.

Keep the `xlvba help` discovery catalog, argparse summaries/examples, and
command handlers synchronized. Distribution tests must prove that the wheel
contains the `.agents/` resources and that its installed `xlvba` entry point
can discover commands and install those resources outside the source checkout.

CLI stdout is a public machine interface: non-interactive commands emit one
versioned JSON result envelope by default. New presentation behavior belongs
behind `--output-format text|table` (or the `--text`/`--table` shortcuts) and
must not add incidental text to default stdout.
