# xlvbatools

> General-purpose Python toolkit for headless Excel VBA automation, debugging, and version control.

## Features

- **Excel COM Session Management** -- Safe context manager with automatic process cleanup and dialog protection
- **Dialog Watchdog** -- Background thread that captures and auto-dismisses modal dialogs during headless automation
- **VBA Source Control** -- Extract, inject, and diff VBA code between workbooks and git-tracked source files
- **Static Analysis** -- 12-category VBA linter that catches reserved keywords, unbalanced blocks, implicit typing, and more
- **Workbook Inspection** -- Screenshot worksheets, dump cell values/formulas, inspect named ranges and shapes
- **Macro Execution** -- Run VBA macros with full dialog protection and structured result reporting
- **Snapshot System** -- Timestamped checkpoint/rollback for workbook and VBA source state
- **Configuration** -- Per-project `xlvbatools.toml` for workbook paths, snapshot limits, and lint rules

## Quick Start

```bash
pip install xlvbatools
```

### As a Python Library

```python
from xlvbatools import ExcelSession

with ExcelSession("path/to/workbook.xlsm") as session:
    result = session.run_macro("MyMacro")
    if session.had_errors:
        print(session.error_summary)
```

### As a CLI

```bash
xlvba init                          # Initialize project
xlvba extract                       # Extract VBA to disk
xlvba inject                        # Inject VBA from disk
xlvba diff                          # Compare workbook vs. disk
xlvba lint                          # Static analysis
xlvba run MyMacro                   # Execute a macro
xlvba snapshot create               # Create checkpoint
xlvba snapshot restore latest       # Rollback
xlvba dump --screenshot --data      # Inspect worksheets
```

## Requirements

- **Python 3.10+**
- **Windows** (COM automation requires Excel installed)
- VBA source file operations (lint, diff, search) work on any platform

## License

MIT
