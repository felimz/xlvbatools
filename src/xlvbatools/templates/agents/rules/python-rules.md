---
trigger: glob
globs: ["**/*.py", "pyproject.toml"]
description: "Python packaging, absolute imports pathways, entry point setups, and test suite execution guidelines."
---

# Python Development Rules

## Packaging & Namespaces
* Always place submodules inside the package namespace (e.g., `xlvbatools/cli/` instead of root-level `cli/`).
* Declare CLI entry point scripts in `pyproject.toml` using absolute imports: `xlvba = "xlvbatools.cli.main:main"`.
* Specify package data inside `[tool.setuptools.package-data]` for non-python files (like markdown templates).

## Testing Practices
* Run unit tests under pytest with `pytest tests/ -v`.
* Use pytest `tmp_path` fixture for any test that writes output files (e.g., logging or temp workbooks) to prevent workspace pollution.
* Mark slow COM-dependent integration tests with `@pytest.mark.com` or `@pytest.mark.e2e`.
