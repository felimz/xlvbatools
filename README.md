# xlvbatools

> General-purpose Python toolkit for headless Excel VBA automation, debugging, and version control.

## Features

- **Excel COM Session Management** -- Safe context manager with automatic process cleanup and dialog protection
- **Dialog Watchdog** -- Background thread that captures and auto-dismisses modal dialogs during headless automation
- **VBA Source Control** -- Extract, inject, and diff VBA code between workbooks and git-tracked source files
- **Static Analysis** -- 30-rule VBA linter catching Dim hoisting, block declarations, unused/dead code, reserved keyword variable names, invalid class/typed object members, executable code outside procedures, hardcoded secrets, absolute paths, Hungarian suffixes, fragile Select/Selection, and silent error suppression
- **VBA Code Formatter** -- Non-destructive indentation normalizer with dry-run and directory-wide support
- **Call Dependency Graph** -- Parses Sub/Function definitions and call sites, outputs Mermaid/DOT/JSON
- **Workbook Inspection** -- Screenshot worksheets, dump cell values/formulas, inspect named ranges and shapes
- **Workbook Modification** -- Programmatically set cell values, formulas, and named ranges
- **Macro Execution** -- Run VBA macros with full dialog protection and structured result reporting
- **Snapshot System** -- Timestamped checkpoint/rollback for workbook and VBA source state
- **VBA Source Search** -- Full-text search across extracted .bas/.cls files (no COM needed)
- **Configuration** -- Per-project `xlvbatools.toml` for workbook paths, snapshot limits, and lint rules
- **Agent Integration** -- Ready-made `.agents/` templates for AI coding assistants

## Quick Start

```bash
pip install xlvbatools
```

### As a Python Library

```python
from xlvbatools.core.session import ExcelSession

with ExcelSession("path/to/workbook.xlsm") as session:
    result = session.run_macro("MyMacro")
    if not result["success"]:
        print(result["error"])
        for event in result["dialog_events"]:
            print(f"  [{event['type']}] {event['text']}")
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
xlvba dump --sheets Sheet1 --json   # Inspect worksheets
xlvba modify --cell A1 --value 42   # Modify cells
xlvba search "MsgBox"               # Search VBA source
xlvba fmt                           # Format VBA code
xlvba graph                         # Call dependency graph
xlvba debug                         # Open Excel + VBE visibly
```

## Project Structure

```text
xlvbatools/
├── src/xlvbatools/              # Python library
│   ├── cli/                     # CLI entry point (xlvba command)
│   │   ├── main.py              # 13 subcommands
│   │   └── init_cmd.py          # Project initialization
│   ├── core/                    # Excel COM automation
│   │   ├── session.py           # ExcelSession context manager
│   │   ├── watchdog.py          # DialogWatchdog background thread
│   │   └── process.py           # Excel process management
│   ├── vba/                     # VBA source operations
│   │   ├── extractor.py         # Extract VBA from workbook
│   │   ├── injector.py          # Inject VBA into workbook
│   │   ├── differ.py            # Diff workbook vs source
│   │   ├── search.py            # Full-text search
│   │   ├── formatter.py         # Code formatting
│   │   ├── dependency.py        # Call graph analysis
│   │   └── manifest.py          # Component manifest tracking
│   ├── analysis/                # Static analysis
│   │   ├── rules.py             # 30 configurable lint rules
│   │   ├── preflight.py         # Top-level lint API
│   │   └── issue.py             # VBAIssue dataclass
│   ├── macro/                   # Macro execution
│   │   └── runner.py            # run_macro() with dialog protection
│   ├── workbook/                # Workbook inspection
│   │   ├── dumper.py            # Screenshots, JSON/MD dumps
│   │   ├── modifier.py          # Cell/formula/named range CRUD
│   │   └── debugger.py          # Interactive debug launcher
│   ├── snapshot/                # Checkpoint/rollback
│   │   └── manager.py           # SnapshotManager
│   ├── config/                  # Configuration
│   │   ├── loader.py            # TOML config loader
│   │   └── schema.py            # XlvbaConfig dataclass
│   ├── logging.py               # Centralized rotating logs
│   └── _compat.py               # Platform compatibility
├── templates/                   # Agent integration templates
│   └── agents/                  # .agents/ directory template
│       ├── AGENTS.md            # VBA development rules
│       ├── skills/              # xlvba-toolchain skill
│       └── workflows/           # vba-edit, vba-debug workflows
├── tests/                       # 118 unit & integration tests
│   ├── test_preflight.py        # Lint rules (37 tests)
│   ├── test_search.py           # VBA search (9 tests)
│   ├── test_snapshot.py         # Snapshot + manifest (12 tests)
│   ├── test_config.py           # Config loading (8 tests)
│   ├── test_formatter.py        # Code formatting (9 tests)
│   ├── test_dependency.py       # Call graph (5 tests)
│   └── ...                      # Core tests (38 tests)
├── docs/                        # Documentation
│   └── agent-integration.md     # Agent usage guide
├── pyproject.toml               # Package config
└── xlvbatools.toml              # Per-project config (user creates)
```

## Configuration

Create `xlvbatools.toml` in your project root:

```toml
[xlvbatools]
workbook = "workbook/MyProject.xlsm"
vba_source = "workbook/vba_source"
snapshots_dir = "snapshots"
log_dir = "logs"

[xlvbatools.snapshots]
rolling_limit = 10

[xlvbatools.backups]
limit = 5

[xlvbatools.lint]
disabled_rules = ["PF001", "PF003"]
```

## Lint Rules

| Rule | Severity | Description |
|:---|:---|:---|
| DS001 | ERROR | Dim after executable code (VBA requires all Dim at top) |
| LC001 | WARNING | Orphaned line continuation (no space before _) |
| SB001 | ERROR | Unbalanced Sub/Function...End blocks |
| PF001 | WARNING | MsgBox call (hangs in headless mode) |
| PF002 | WARNING | Implicit Variant (Dim without As clause) |
| PF003 | WARNING | ActiveSheet/ActiveCell usage (fragile) |
| OE001 | WARNING | Option Explicit missing |
| BK001 | WARNING | Block-level variable declaration (hoisting risk) |
| SD002 | STYLE | Multiple variable declarations on one line |
| PF004 | STYLE | Avoid Integer data type (prefer Long) |
| SD005 | STYLE | Missing parameter modifier (explicit ByVal/ByRef required) |
| SD006 | STYLE | Missing access modifier (Sub/Function should declare Public/Private) |
| SD010 | STYLE | Line exceeds maximum length limit (120 characters) |
| SD014 | STYLE | Obsolete 'Call' keyword used for procedure call |
| SC001 | WARNING | Hardcoded password or secret string |
| SC002 | WARNING | Absolute file path literal detected |
| CL001 | STYLE | Hungarian type suffix used (%, &, $, etc.) |
| CL002 | STYLE | Selection or ActiveWindow dependency pattern |
| SF001 | WARNING | Silent error suppression (On Error Resume Next with no check) |
| DC001 | WARNING | Unused local variable |
| DC002 | WARNING | Empty procedure declaration |
| DC003 | WARNING | Dead procedure (0 incoming calls, graph-level analysis) |
| SD015 | STYLE | Multiple consecutive blank lines |
| SD016 | STYLE | Double-spaced code blocks (alternating blank lines) |
| RK001 | WARNING | Variable/parameter name is a VBA reserved keyword |
| IP001 | ERROR | Executable code found outside procedure block |
| DP001 | ERROR | Duplicate public procedure declared across modules |
| SM001 | WARNING | Invalid property/method referenced on class or typed object |
| CT001 | ERROR | VBE compile test failed (synthetic, COM-only) |

## Agent Integration

For AI coding assistants, install the `.agents/` template:

```bash
xlvba init --agents
```

This installs:
- `.agents/AGENTS.md` -- VBA development rules
- `.agents/skills/xlvba-toolchain/SKILL.md` -- CLI quick reference
- `.agents/workflows/vba-edit.md` -- Edit-verify cycle
- `.agents/workflows/vba-debug.md` -- Error diagnosis

See [docs/agent-integration.md](docs/agent-integration.md) for the complete guide.

## Requirements

- **Python 3.10+**
- **Windows** (COM automation requires Excel installed)
- VBA source operations (lint, diff, search, format, graph) work on **any platform**

## License

MIT
