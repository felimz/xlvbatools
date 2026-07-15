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
The source templates under `templates/agents/` and the packaged copies under
`src/xlvbatools/templates/agents/` must remain byte-for-byte identical.

Validate documented CLI examples against `xlvba <command> --help`. Safety
guidance must never recommend global Excel termination or direct use of
implementation subpackages.
