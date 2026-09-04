"""Repository-layout and ignore contracts for maintainers."""

from __future__ import annotations

import ast
import importlib
import importlib.util
from pathlib import Path
import shutil
import subprocess

import pytest


pytestmark = pytest.mark.unit
ROOT = Path(__file__).resolve().parents[1]


def _require_git_checkout() -> None:
    if shutil.which("git") is None or not (ROOT / ".git").exists():
        pytest.skip("repository hygiene checks require a Git checkout")


def _ignored_paths(paths: tuple[str, ...]) -> set[str]:
    completed = subprocess.run(
        ["git", "check-ignore", "--no-index", "--stdin", "-z"],
        cwd=ROOT,
        input="\0".join(paths) + "\0",
        capture_output=True,
        text=True,
        check=False,
    )
    assert completed.returncode in {0, 1}, completed.stderr
    return {path for path in completed.stdout.split("\0") if path}


def test_gitignore_covers_generated_local_and_binary_state():
    _require_git_checkout()

    ignored = (
        ".venv/Lib/site.py",
        "src/xlvbatools/__pycache__/project.pyc",
        "src/xlvbatools.egg-info/PKG-INFO",
        "build/lib/xlvbatools.py",
        "dist/xlvbatools.whl",
        "xlvbatools.whl",
        ".coverage",
        "coverage.xml",
        "htmlcov/index.html",
        ".tox/py/python.exe",
        ".hypothesis/examples/state",
        "test-results/junit.xml",
        ".env",
        ".env.local",
        "logs/test.log",
        "crash.dmp",
        "sample_workbooks/Consumer.xlsm",
        "~$Consumer.xlsm",
        "Consumer.xlsm",
    )
    assert set(ignored) - _ignored_paths(ignored) == set()


def test_gitignore_preserves_reviewable_source_config_and_fixtures():
    _require_git_checkout()

    reviewable = (
        ".agents/AGENTS.md",
        "src/xlvbatools/templates/agents/AGENTS.md",
        ".env.example",
        ".env.template",
        "xlvbatools.toml",
        "src/model.frx",
        "docs/architecture.png",
        "tests/fixtures/tiny.xlsm",
        "tests/fixtures/tiny.xlsx",
        "tests/fixtures/tiny.xls",
        "tests/fixtures/tiny.xlsb",
        "tests/fixtures/nested/tiny.xlsm",
    )
    assert _ignored_paths(reviewable) == set()


def test_no_tracked_file_is_now_classified_as_generated():
    _require_git_checkout()

    completed = subprocess.run(
        ["git", "ls-files", "-ci", "--exclude-standard"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=False,
    )
    assert completed.returncode == 0, completed.stderr
    assert completed.stdout.strip() == ""

    tracked = subprocess.run(
        ["git", "ls-files"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=True,
    ).stdout.splitlines()
    forbidden_parts = {"__pycache__", ".pytest_cache", ".mypy_cache", ".ruff_cache"}
    assert [path for path in tracked if forbidden_parts.intersection(Path(path).parts)] == []
    assert [
        path for path in tracked
        if Path(path).suffix.lower() in {".xls", ".xlsb", ".xlsm", ".xlsx"}
        and not path.startswith("tests/fixtures/")
    ] == []


def test_runtime_package_has_no_registry_dependency_or_mutation():
    """Default and optional package paths must not depend on registry access."""
    forbidden = ("winreg", "LoadBehavior", "SetValueEx", "Set-ItemProperty")
    matches = []
    for path in sorted((ROOT / "src" / "xlvbatools").rglob("*.py")):
        source = path.read_text(encoding="utf-8")
        for token in forbidden:
            if token in source:
                matches.append(f"{path.relative_to(ROOT)}: {token}")

    assert matches == []


def test_all_xlvbatools_symbols_imported_by_tests_still_resolve():
    """Catch stale test imports during the fast suite instead of at runtime."""
    missing = []
    for path in sorted((ROOT / "tests").glob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
        for node in ast.walk(tree):
            if not isinstance(node, ast.ImportFrom):
                continue
            module_name = node.module or ""
            if not module_name.startswith("xlvbatools"):
                continue
            module = importlib.import_module(module_name)
            for alias in node.names:
                if alias.name == "*" or hasattr(module, alias.name):
                    continue
                submodule = f"{module_name}.{alias.name}"
                try:
                    found = importlib.util.find_spec(submodule) is not None
                except (AttributeError, ModuleNotFoundError, ValueError):
                    found = False
                if not found:
                    missing.append(
                        f"{path.relative_to(ROOT)}:{node.lineno}: "
                        f"from {module_name} import {alias.name}"
                    )

    assert missing == []
