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

Tests are managed using `pytest`. The test suite is divided into `unit` (offline) and `com` (requires Excel) tests.

Run all tests:
```bash
.venv\Scripts\python.exe -m pytest
```

Run only offline/unit tests:
```bash
.venv\Scripts\python.exe -m pytest -m unit
```

Run tests with coverage:
```bash
.venv\Scripts\python.exe -m pytest --cov=xlvbatools
```

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
