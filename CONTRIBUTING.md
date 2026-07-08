# Contributing to xlvbatools

Thank you for your interest in contributing! This guide helps you set up a local development environment and run tests.

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
pip install -e .[dev]
```

---

## Running Tests

Tests are managed using `pytest`. The test suite is divided into `unit` (offline) and `com` (requires Excel) tests.

Run all tests:
```bash
pytest
```

Run only offline/unit tests:
```bash
pytest -m unit
```

Run tests with coverage:
```bash
pytest --cov=xlvbatools
```

---

## Coding Guidelines

* **Type Safety:** The project uses PEP 561 types. All new modules must have explicit annotations and be compliant with static type checkers.
* **Compatibility:** Always check platform status before using `win32com.client` or native win32 calls by referencing `xlvbatools._compat.IS_WINDOWS`.
* **COM Safety:** Never perform a raw `SendMessageW` from the watchdog thread; use `SendMessageTimeoutW` to prevent deadlocks when VBE or Excel hangs.
