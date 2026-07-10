# xlvbatools - Mission Control

## Project Overview
`xlvbatools` is a professional, developer-friendly toolkit for headless Excel VBA automation, debugging, static analysis, and version control.

## Technology Stack
* **Python**: Core scripting language (supporting Python 3.10+).
* **win32com**: COM automation interface for Excel.
* **pytest**: Test framework for unit, COM, and end-to-end testing.

## Key Developer Rules
* **Strict Namespace Boundaries**: Keep the top-level namespace clean. All code, CLI commands, and templates reside under `src/xlvbatools/`.
* **Zero Pollution**: Test runs must never write temporary artifacts or debug folders to the workspace root. Always use `tmp_path` or standard test directories.
* **Progressive Disclosure**: Specific coding standards for VBA and Python are defined in the task-specific glob rules under `.agents/rules/` to prevent context bloat.
* **Agent Integration Guide**: Run `xlvba agents` or `xlvba --agents` to print the full integration instructions and best practices for the toolkit.
