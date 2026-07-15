"""Installed-wheel contract test in a clean consumer environment."""

import json
from pathlib import Path
import subprocess
import sys

import pytest


@pytest.mark.integration
@pytest.mark.distribution
def test_built_wheel_exposes_public_wrapper_api(tmp_path):
    project_root = Path(__file__).resolve().parents[1]
    wheel_dir = tmp_path / "wheel"
    venv_dir = tmp_path / "consumer-venv"
    outside = tmp_path / "outside"
    wheel_dir.mkdir()
    outside.mkdir()

    built = subprocess.run(
        [
            sys.executable, "-m", "pip", "wheel", str(project_root),
            "--no-deps", "--wheel-dir", str(wheel_dir),
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )
    assert built.returncode == 0, built.stdout + built.stderr
    wheel = next(wheel_dir.glob("xlvbatools-*.whl"))
    assert wheel.name.startswith("xlvbatools-1.0.0-")

    created = subprocess.run(
        [sys.executable, "-m", "venv", str(venv_dir)],
        capture_output=True,
        text=True,
        timeout=120,
    )
    assert created.returncode == 0, created.stdout + created.stderr
    consumer_python = venv_dir / ("Scripts/python.exe" if sys.platform == "win32" else "bin/python")

    installed = subprocess.run(
        [
            str(consumer_python), "-m", "pip", "install", str(wheel),
            "--no-deps",
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )
    assert installed.returncode == 0, installed.stdout + installed.stderr

    code = (
        "import json, pathlib, sys; "
        "import xlvbatools; "
        "from xlvbatools import Operation, OperationRequest, OperationResult, Project, VBAIssue; "
        "from xlvbatools.core.protocol import WORKER_PROTOCOL_VERSION; "
        "print(json.dumps({'module': xlvbatools.__file__, 'version': xlvbatools.__version__, "
        "'exports': [item.__name__ for item in "
        "(Project, Operation, OperationRequest, OperationResult, VBAIssue)], "
        "'worker_protocol': WORKER_PROTOCOL_VERSION}))"
    )
    imported = subprocess.run(
        [str(consumer_python), "-I", "-c", code],
        cwd=outside,
        capture_output=True,
        text=True,
        timeout=30,
    )
    assert imported.returncode == 0, imported.stdout + imported.stderr
    payload = json.loads(imported.stdout)
    assert Path(payload["module"]).resolve().is_relative_to(venv_dir.resolve())
    assert payload["version"] == "1.0.0"
    assert payload["exports"] == [
        "Project", "Operation", "OperationRequest", "OperationResult", "VBAIssue",
    ]
    assert payload["worker_protocol"] == "2.0"
