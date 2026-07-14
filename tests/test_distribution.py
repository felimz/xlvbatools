"""Installed-wheel contract test, isolated from the repository source tree."""

import json
import os
from pathlib import Path
import subprocess
import sys

import pytest


@pytest.mark.integration
@pytest.mark.distribution
def test_built_wheel_exposes_public_wrapper_api(tmp_path):
    project_root = Path(__file__).resolve().parents[1]
    wheel_dir = tmp_path / "wheel"
    site_dir = tmp_path / "site"
    outside = tmp_path / "outside"
    wheel_dir.mkdir()
    outside.mkdir()

    built = subprocess.run(
        [
            sys.executable, "-m", "pip", "wheel", str(project_root),
            "--no-deps", "--no-build-isolation", "--wheel-dir", str(wheel_dir),
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )
    assert built.returncode == 0, built.stdout + built.stderr
    wheel = next(wheel_dir.glob("xlvbatools-*.whl"))

    installed = subprocess.run(
        [
            sys.executable, "-m", "pip", "install", str(wheel),
            "--no-deps", "--target", str(site_dir),
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )
    assert installed.returncode == 0, installed.stdout + installed.stderr

    code = (
        "import json, pathlib, sys; "
        f"sys.path.insert(0, {str(site_dir)!r}); "
        "import xlvbatools; "
        "from xlvbatools import OperationResult, XlvbaProject, lint_files; "
        "from xlvbatools.analysis import VBAIssue, lint_workbook; "
        "from xlvbatools.core.worker import WORKER_PROTOCOL_VERSION; "
        "print(json.dumps({'module': xlvbatools.__file__, "
        "'exports': [item.__name__ for item in "
        "(OperationResult, XlvbaProject, VBAIssue, lint_files, lint_workbook)], "
        "'worker_protocol': WORKER_PROTOCOL_VERSION}))"
    )
    environment = os.environ.copy()
    environment.pop("PYTHONPATH", None)
    imported = subprocess.run(
        [sys.executable, "-c", code],
        cwd=outside,
        env=environment,
        capture_output=True,
        text=True,
        timeout=30,
    )
    assert imported.returncode == 0, imported.stdout + imported.stderr
    payload = json.loads(imported.stdout)
    assert Path(payload["module"]).resolve().is_relative_to(site_dir.resolve())
    assert payload["exports"] == [
        "OperationResult", "XlvbaProject", "VBAIssue", "lint_files",
        "lint_workbook",
    ]
    assert payload["worker_protocol"] == "1.0"
