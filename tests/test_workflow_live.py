"""Live Excel acceptance for isolated one-session workflows."""

from __future__ import annotations

import json
from pathlib import Path
import subprocess
import sys
import tempfile
import textwrap
import xml.etree.ElementTree as ET
import zipfile

import pytest


pytestmark = [
    pytest.mark.com,
    pytest.mark.integration,
    pytest.mark.e2e,
    pytest.mark.skipif(sys.platform != "win32", reason="Windows only"),
]


def _saved_cell_value(workbook: str, address: str) -> float | None:
    namespace = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(workbook) as archive:
        root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    cell = root.find(f".//x:c[@r='{address}']/x:v", namespace)
    return float(cell.text) if cell is not None and cell.text is not None else None


def _calculation_steps(*, screenshots: bool = False, output_dir: str = "screenshots"):
    from xlvbatools import InspectStep, MacroStep, ModifyStep

    return [
        MacroStep("retrieve", "WorkflowRetrieve"),
        ModifyStep("inputs", "Sheet1", {"A2": 5}, calculate=False),
        MacroStep("calculate", "WorkflowCalculate"),
        InspectStep(
            "results",
            ("Sheet1",),
            output_dir=output_dir,
            cell_range="A1:A5",
            include_data=True,
            include_screenshots=screenshots,
        ),
    ]


def test_live_workflow_shares_one_excel_session_and_discards_without_save(
    runtime_error_workbook, tmp_path,
):
    from xlvbatools import MacroOutput, Project

    result = Project.open(runtime_error_workbook).workflow(
        _calculation_steps(
            screenshots=True,
            output_dir=str(tmp_path / "screenshots"),
        ),
        timeout=90,
        save=False,
    )

    workflow = result.require_success()
    cleanup = result.require_clean_shutdown()
    assert cleanup.still_running is False
    assert [step.status for step in workflow.steps] == ["succeeded"] * 4
    retrieve = workflow.step("retrieve").data
    calculate = workflow.step("calculate").data
    assert isinstance(retrieve, MacroOutput)
    assert isinstance(calculate, MacroOutput)
    assert retrieve.excel_pid == calculate.excel_pid == result.diagnostics.excel_pid
    inspection = workflow.step("results").data
    assert inspection.workbook_data["sheets"]["Sheet1"]["cells"]["A3"]["value"] == 15
    screenshot = Path(inspection.screenshots["Sheet1"])
    assert screenshot.is_file()
    assert result.artifacts[0].path == str(screenshot)
    assert _saved_cell_value(runtime_error_workbook, "A3") is None


def test_screenshot_repaints_after_macro_and_restores_screen_updating(
    runtime_error_workbook, tmp_path,
):
    from xlvbatools import InspectStep, MacroStep, Project
    from xlvbatools.workbook.dumper import _native_image_metrics

    result = Project.open(runtime_error_workbook).workflow(
        [
            MacroStep("prepare", "LeaveScreenUpdatingOff"),
            InspectStep(
                "inspect",
                ("Sheet1",),
                output_dir=str(tmp_path / "screenshots"),
                cell_range="A1:B2",
                include_data=True,
                include_screenshots=True,
            ),
            MacroStep("verify-state", "VerifyScreenUpdatingStillOff"),
        ],
        timeout=90,
        save=False,
    )

    workflow = result.require_success()
    screenshot = Path(workflow.step("inspect").data.screenshots["Sheet1"])
    assert screenshot.is_file()
    assert _native_image_metrics(str(screenshot))["meaningful_pixel_count"] > 64
    assert workflow.step("verify-state").status == "succeeded"
    assert result.require_clean_shutdown().still_running is False


def test_live_workflow_saves_once_only_after_complete_success(runtime_error_workbook):
    from xlvbatools import MacroStep, ModifyStep, Project

    result = Project.open(runtime_error_workbook).workflow(
        [
            MacroStep("retrieve", "WorkflowRetrieve"),
            ModifyStep("inputs", "Sheet1", {"A2": 5}),
            MacroStep("calculate", "WorkflowCalculate"),
        ],
        timeout=90,
        save=True,
    )

    workflow = result.require_success()
    result.require_clean_shutdown()
    assert workflow.save_requested is True
    assert workflow.saved is True
    assert _saved_cell_value(runtime_error_workbook, "A3") == 15


def test_live_workflow_is_fail_fast_and_does_not_save_failed_state(
    runtime_error_workbook,
):
    from xlvbatools import InspectStep, MacroStep, Project

    result = Project.open(runtime_error_workbook).workflow(
        [
            MacroStep("fail", "WorkflowFail"),
            MacroStep("never", "WorkflowMustNotRun"),
            InspectStep("results", ("Sheet1",), include_screenshots=False),
        ],
        timeout=90,
        save=True,
    )

    assert result.success is False
    assert result.error.code == "workflow_step_failed"
    assert result.data.failed_step_id == "fail"
    assert [step.status for step in result.data.steps] == [
        "failed", "not_run", "not_run",
    ]
    assert result.data.saved is False
    assert result.attempt_count == 1
    result.require_clean_shutdown()
    assert _saved_cell_value(runtime_error_workbook, "A4") is None
    assert _saved_cell_value(runtime_error_workbook, "A5") is None


def test_live_workflow_timeout_retains_step_progress_and_does_not_replay(
    runtime_error_workbook,
):
    from xlvbatools import MacroStep, Project

    result = Project.open(runtime_error_workbook).workflow(
        [
            MacroStep("loop", "LoopForever"),
            MacroStep("never", "WorkflowMustNotRun"),
        ],
        timeout=8,
        save=False,
    )

    assert result.success is False
    assert result.error.code == "timeout"
    assert result.attempt_count == 1
    assert result.data.failed_step_id == "loop"
    assert result.data.save_requested is False
    assert result.diagnostics.progress["step_id"] == "loop"
    assert result.diagnostics.progress["step_phase"] == "macro_execution"
    cleanup = result.diagnostics.cleanup
    assert cleanup is not None
    assert cleanup.still_running is False


def test_live_workflow_cli_emits_one_machine_result_envelope(
    runtime_error_workbook, tmp_path,
):
    workflow_file = tmp_path / "workflow.json"
    workflow_file.write_text(
        json.dumps({
            "workflow_schema_version": "1.0",
            "steps": [
                {"id": "retrieve", "kind": "macro", "macro": "WorkflowRetrieve"},
                {
                    "id": "inputs", "kind": "modify", "sheet": "Sheet1",
                    "values": {"A2": 5},
                },
                {
                    "id": "calculate", "kind": "macro",
                    "macro": "WorkflowCalculate",
                },
                {
                    "id": "results", "kind": "inspect", "sheets": ["Sheet1"],
                    "cell_range": "A1:A3", "include_screenshots": False,
                },
            ],
        }),
        encoding="utf-8",
    )
    completed = subprocess.run(
        [
            sys.executable, "-m", "xlvbatools.cli.main", "workflow",
            "--workbook", runtime_error_workbook,
            "--file", str(workflow_file),
            "--no-save", "--timeout", "90",
        ],
        capture_output=True,
        text=True,
        timeout=120,
    )

    assert completed.returncode == 0, completed.stdout + completed.stderr
    payload = json.loads(completed.stdout)
    assert payload["operation"] == "workflow"
    assert payload["success"] is True
    assert len(payload["data"]["steps"]) == 4
    cells = payload["data"]["steps"][3]["data"]["workbook_data"]["sheets"][
        "Sheet1"
    ]["cells"]
    assert cells["A3"]["value"] == 15
    assert payload["diagnostics"]["cleanup"]["still_running"] is False


def test_twenty_five_live_workflows_leave_no_process_or_finalizer_diagnostics(
    runtime_error_workbook,
):
    code = textwrap.dedent(
        """
        import json
        import sys

        from xlvbatools import InspectStep, MacroStep, ModifyStep, Project
        from xlvbatools.core.process import is_process_running

        project = Project.open(sys.argv[1])
        for iteration in range(25):
            result = project.workflow(
                [
                    MacroStep("retrieve", "WorkflowRetrieve"),
                    ModifyStep("inputs", "Sheet1", {"A2": iteration}),
                    MacroStep("calculate", "WorkflowCalculate"),
                    InspectStep(
                        "results", ("Sheet1",), cell_range="A1:A3",
                        include_screenshots=False,
                    ),
                ],
                timeout=90,
                save=False,
            )
            workflow = result.require_success()
            cleanup = result.require_clean_shutdown()
            assert not is_process_running(cleanup.pid), result.to_dict()
            value = workflow.step("results").data.workbook_data[
                "sheets"
            ]["Sheet1"]["cells"]["A3"]["value"]
            assert value == 10 + iteration, result.to_dict()
            print("WORKFLOW_PROGRESS=" + json.dumps({
                "iteration": iteration + 1,
                "request_id": result.request_id,
                "attempt_count": result.attempt_count,
                "excel_pid": result.diagnostics.excel_pid,
                "cleanup": result.to_dict()["diagnostics"]["cleanup"],
            }), flush=True)
        print("WORKFLOW_STRESS_COMPLETE=25", flush=True)
        """
    )
    with tempfile.TemporaryFile(mode="w+", encoding="utf-8", errors="replace") as output:
        completed = subprocess.run(
            [sys.executable, "-X", "faulthandler", "-c", code, runtime_error_workbook],
            stdout=output,
            stderr=subprocess.STDOUT,
            timeout=600,
        )
        output.seek(0)
        combined = output.read()

    assert completed.returncode == 0, combined
    assert "WORKFLOW_STRESS_COMPLETE=25" in combined, combined
    assert combined.count("WORKFLOW_PROGRESS=") == 25, combined
    for signature in (
        "Windows fatal exception", "0x800706ba", "0x80010108",
        "RPC server is unavailable", "CoInitialize has not been called",
    ):
        assert signature not in combined, combined
