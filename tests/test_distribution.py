"""Installed-wheel contract test in a clean consumer environment."""

import json
from pathlib import Path
import subprocess
import sys

import pytest


@pytest.mark.integration
@pytest.mark.distribution
def test_built_wheel_exposes_public_wrapper_api(tmp_path):
    from xlvbatools import __version__

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
    assert wheel.name.startswith(f"xlvbatools-{__version__}-")

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
            "--no-index",
            "--disable-pip-version-check",
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )
    assert installed.returncode == 0, installed.stdout + installed.stderr

    code = (
        "import importlib.resources, json, pathlib, sys; "
        "import xlvbatools; "
        "from xlvbatools import (MacroStep, Operation, OperationRequest, OperationResult, "
        "Project, VBAIssue, WORKFLOW_SCHEMA_VERSION, WorkflowOutput); "
        "from xlvbatools.core.protocol import WORKER_PROTOCOL_VERSION; "
        "print(json.dumps({'module': xlvbatools.__file__, 'version': xlvbatools.__version__, "
        "'exports': [item.__name__ for item in "
        "(Project, Operation, OperationRequest, OperationResult, VBAIssue, MacroStep, "
        "WorkflowOutput)], "
        "'worker_protocol': WORKER_PROTOCOL_VERSION, "
        "'workflow_schema': WORKFLOW_SCHEMA_VERSION, "
        "'agent_template': importlib.resources.files('xlvbatools')"
        ".joinpath('templates/agents/AGENTS.md').is_file()}))"
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
    assert payload["version"] == __version__
    assert payload["exports"] == [
        "Project", "Operation", "OperationRequest", "OperationResult", "VBAIssue",
        "MacroStep", "WorkflowOutput",
    ]
    assert payload["worker_protocol"] == "2.1"
    assert payload["workflow_schema"] == "1.0"
    assert payload["agent_template"] is True

    consumer_xlvba = venv_dir / ("Scripts/xlvba.exe" if sys.platform == "win32" else "bin/xlvba")
    discovered = subprocess.run(
        [str(consumer_xlvba), "help"],
        cwd=outside,
        capture_output=True,
        text=True,
        timeout=30,
    )
    assert discovered.returncode == 0, discovered.stdout + discovered.stderr
    discovery = json.loads(discovered.stdout)
    assert discovery["operation"] == "help"
    assert discovery["data"]["agent_templates"]["destination"] == ".agents/"
    workflow_command = next(
        item for item in discovery["data"]["commands"] if item["name"] == "workflow"
    )
    assert workflow_command["input_schema"]["workflow_schema_version"] == "1.0"

    agent_install = subprocess.run(
        [str(consumer_xlvba), "agents", "install"],
        cwd=outside,
        capture_output=True,
        text=True,
        timeout=30,
    )
    assert agent_install.returncode == 0, agent_install.stdout + agent_install.stderr
    install_payload = json.loads(agent_install.stdout)
    assert install_payload["operation"] == "agents_install"
    assert (outside / ".agents/AGENTS.md").is_file()
    assert (outside / ".agents/workflows/get-started.md").is_file()
    assert (outside / ".agents/workflows/excel-workflow.md").is_file()
