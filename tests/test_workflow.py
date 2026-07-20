"""Typed public workflow contract and Project boundary tests."""

from __future__ import annotations

import json
import math
from types import SimpleNamespace

import pytest


class RecordingExecutor:
    def __init__(self, response=None):
        self.response = response
        self.requests = []

    def execute(self, request):
        from xlvbatools import OperationResult

        self.requests.append(request)
        return self.response or OperationResult(
            operation=request.operation.value,
            success=True,
            phase="complete",
            data={
                "workflow_schema_version": "1.0",
                "steps": [],
                "save_requested": False,
                "saved": False,
            },
        )


@pytest.mark.unit
def test_workflow_step_contracts_validate_and_freeze_inputs():
    from xlvbatools import InspectStep, MacroStep, ModifyStep

    named_ranges = {"FilePath": "model.r3d"}
    values = {"C102:C104": [[0.1], [0.0], [-0.1]]}
    macro = MacroStep("retrieve", "OnRetrieve", named_ranges)
    modify = ModifyStep("inputs", "Input", values)
    inspect = InspectStep("results", ("Input",), include_screenshots=False)
    named_ranges["FilePath"] = "changed.r3d"
    values["C102:C104"].append([99])

    assert macro.named_ranges["FilePath"] == "model.r3d"
    assert modify.values["C102:C104"] == ((0.1,), (0.0,), (-0.1,))
    assert inspect.include_data is True
    with pytest.raises(TypeError):
        macro.named_ranges["new"] = 1
    with pytest.raises(TypeError):
        modify.values["A1"] = 1


@pytest.mark.unit
@pytest.mark.parametrize(
    "case, match",
    [
        ("bad_id", "step id"),
        ("empty_macro", "non-empty"),
        ("empty_sheet", "sheet"),
        ("no_values", "at least one"),
        ("bad_range", "invalid A1"),
        ("range_column_limit", "outside Excel limits"),
        ("range_row_limit", "outside Excel limits"),
        ("ragged", "rectangular"),
        ("nonfinite", "must not contain"),
        ("empty_inspect", "must include"),
    ],
)
def test_workflow_step_contracts_reject_invalid_requests(case, match):
    from xlvbatools import InspectStep, MacroStep, ModifyStep

    factories = {
        "bad_id": lambda: MacroStep("bad id", "OnRetrieve"),
        "empty_macro": lambda: MacroStep("retrieve", ""),
        "empty_sheet": lambda: ModifyStep("inputs", "", {"A1": 1}),
        "no_values": lambda: ModifyStep("inputs", "Input", {}),
        "bad_range": lambda: ModifyStep("inputs", "Input", {"not a range": 1}),
        "range_column_limit": lambda: ModifyStep("inputs", "Input", {"XFE1": 1}),
        "range_row_limit": lambda: ModifyStep(
            "inputs", "Input", {"A1048577": 1},
        ),
        "ragged": lambda: ModifyStep(
            "inputs", "Input", {"A1:B2": [[1], [2, 3]]},
        ),
        "nonfinite": lambda: ModifyStep(
            "inputs", "Input", {"A1": math.inf},
        ),
        "empty_inspect": lambda: InspectStep(
            "results", ("Input",), include_data=False, include_screenshots=False,
        ),
    }
    with pytest.raises((TypeError, ValueError), match=match):
        factories[case]()


@pytest.mark.unit
def test_project_workflow_validates_before_executor(tmp_path):
    from xlvbatools import MacroStep, Project

    executor = RecordingExecutor()
    project = Project.open(tmp_path / "book.xlsm", executor=executor)

    with pytest.raises(ValueError, match="at least one"):
        project.workflow([])
    with pytest.raises(ValueError, match="unique"):
        project.workflow([MacroStep("same", "One"), MacroStep("SAME", "Two")])
    with pytest.raises(TypeError, match="visible must be boolean"):
        project.workflow([MacroStep("one", "One")], visible="false")
    with pytest.raises(TypeError, match="save must be boolean"):
        project.workflow([MacroStep("one", "One")], save=1)

    assert executor.requests == []


@pytest.mark.unit
def test_project_workflow_sends_one_request_and_normalizes_typed_steps(tmp_path):
    from xlvbatools import (
        InspectStep,
        MacroOutput,
        MacroStep,
        ModifyStep,
        ModifyStepOutput,
        Operation,
        OperationResult,
        Project,
    )

    response = OperationResult(
        operation="workflow",
        success=True,
        phase="complete",
        data={
            "workflow_schema_version": "1.0",
            "save_requested": False,
            "saved": False,
            "failed_step_id": None,
            "steps": [
                {
                    "id": "retrieve", "kind": "macro", "status": "succeeded",
                    "phase": "complete", "elapsed_seconds": 1.0,
                    "data": {"macro": "OnRetrieve", "run_id": "run-1"},
                },
                {
                    "id": "inputs", "kind": "modify", "status": "succeeded",
                    "phase": "complete", "elapsed_seconds": 0.2,
                    "data": {
                        "applied": True, "calculated": False,
                        "writes": [{
                            "sheet": "Input", "range": "C102:C104",
                            "rows": 3, "columns": 1,
                        }],
                    },
                },
                {
                    "id": "results", "kind": "inspect", "status": "succeeded",
                    "phase": "complete", "elapsed_seconds": 0.3,
                    "data": {
                        "workbook_data": {"sheets": {}},
                        "screenshots": {"Input": "artifacts/Input.png"},
                    },
                },
            ],
        },
    )
    executor = RecordingExecutor(response)
    project = Project.open(tmp_path / "book.xlsm", executor=executor)
    result = project.workflow([
        MacroStep("retrieve", "OnRetrieve"),
        ModifyStep("inputs", "Input", {"C102:C104": [[0.1], [0.0], [-0.1]]}),
        InspectStep(
            "results", ("Input",), output_dir="artifacts",
            include_screenshots=True,
        ),
    ])

    assert len(executor.requests) == 1
    request = executor.requests[0]
    assert request.operation is Operation.WORKFLOW
    assert request.timeout == 240
    assert request.retry_transient is False
    assert request.arguments["workflow_schema_version"] == "1.0"
    assert request.arguments["steps"][1]["values"]["C102:C104"] == (
        (0.1,), (0.0,), (-0.1,)
    )
    assert result.data.failed_step_id is None
    assert isinstance(result.data.step("retrieve").data, MacroOutput)
    assert isinstance(result.data.step("inputs").data, ModifyStepOutput)
    assert result.data.by_id["results"].data.workbook_data == {"sheets": {}}
    assert result.artifacts[0].metadata["step_id"] == "results"
    json.dumps(result.to_dict())


@pytest.mark.unit
def test_project_workflow_preserves_failed_step_evidence(tmp_path):
    from xlvbatools import ErrorInfo, InspectStep, MacroStep, OperationResult, Project

    response = OperationResult(
        operation="workflow",
        success=False,
        phase="macro_execution",
        error=ErrorInfo(
            message="calculation failed",
            code="workflow_step_failed",
            details={"step_id": "calculate"},
        ),
        data={
            "workflow_schema_version": "1.0",
            "failed_step_id": "calculate",
            "save_requested": True,
            "saved": False,
            "steps": [
                {
                    "id": "calculate", "kind": "macro", "status": "failed",
                    "phase": "macro_execution", "data": {
                        "macro": "OnCalculate", "run_id": "run-failed",
                    },
                    "error": {
                        "message": "calculation failed", "code": "macro_failed",
                        "details": {"dialog": "runtime error"},
                    },
                },
                {
                    "id": "results", "kind": "inspect", "status": "not_run",
                    "phase": "not_run", "data": None,
                },
            ],
        },
    )
    project = Project.open(
        tmp_path / "book.xlsm", executor=RecordingExecutor(response),
    )
    result = project.workflow([
        MacroStep("calculate", "OnCalculate"),
        InspectStep("results", ("Input",)),
    ], save=True)

    assert result.success is False
    assert result.data.failed_step_id == "calculate"
    assert result.data.step("calculate").error.code == "macro_failed"
    assert result.data.step("results").status == "not_run"
    assert result.data.saved is False


@pytest.mark.unit
def test_project_workflow_retains_request_intent_and_progress_without_worker_data(
    tmp_path,
):
    from xlvbatools import Diagnostics, ErrorInfo, MacroStep, OperationResult, Project

    response = OperationResult(
        operation="workflow",
        success=False,
        phase="workflow_step",
        error=ErrorInfo(message="timed out", code="timeout"),
        diagnostics=Diagnostics(progress={
            "step_id": "calculate",
            "step_phase": "macro_execution",
        }),
    )
    project = Project.open(
        tmp_path / "book.xlsm", executor=RecordingExecutor(response),
    )

    result = project.workflow(
        [MacroStep("calculate", "OnCalculate")], save=True,
    )

    assert result.data.steps == ()
    assert result.data.failed_step_id == "calculate"
    assert result.data.save_requested is True
    assert result.data.saved is False


@pytest.mark.unit
def test_workflow_payload_parser_is_strict_and_typed():
    from xlvbatools.workflow import _steps_from_payload

    steps = _steps_from_payload([
        {"id": "retrieve", "kind": "macro", "macro": "OnRetrieve"},
        {
            "id": "inputs", "kind": "modify", "sheet": "Input",
            "values": {"A1:A2": [[1], [2]]},
        },
        {"id": "results", "kind": "inspect", "sheets": ["Input"]},
    ])

    assert [step.id for step in steps] == ["retrieve", "inputs", "results"]
    with pytest.raises(ValueError, match="unsupported kind"):
        _steps_from_payload([{"id": "bad", "kind": "shell"}])
    with pytest.raises(ValueError, match="unknown fields: typo"):
        _steps_from_payload([
            {"id": "one", "kind": "macro", "macro": "One", "typo": True},
        ])
    with pytest.raises(TypeError, match="calculate must be boolean"):
        _steps_from_payload([{
            "id": "one", "kind": "modify", "sheet": "Input",
            "values": {"A1": 1}, "calculate": "false",
        }])
    with pytest.raises(TypeError, match="sheets must be an array"):
        _steps_from_payload([{
            "id": "one", "kind": "inspect", "sheets": "Input",
        }])
    with pytest.raises(TypeError, match="macro must be a string"):
        _steps_from_payload([{"id": "one", "kind": "macro", "macro": 1}])
    with pytest.raises(TypeError, match="named_ranges must be an object"):
        _steps_from_payload([{
            "id": "one", "kind": "macro", "macro": "One",
            "named_ranges": None,
        }])


@pytest.mark.unit
def test_range_writer_uses_existing_session_and_calculates_only_when_requested():
    from xlvbatools.workbook.modifier import _write_ranges_in_session

    class FakeRange:
        def __init__(self, rows, columns):
            self.Rows = SimpleNamespace(Count=rows)
            self.Columns = SimpleNamespace(Count=columns)
            self.Value = "unchanged"

    ranges = {"A1": FakeRange(1, 1), "B2:C3": FakeRange(2, 2)}
    worksheet = SimpleNamespace(Range=lambda address: ranges[address])
    calculations = []
    session = SimpleNamespace(
        wb=SimpleNamespace(Worksheets=lambda sheet: worksheet),
        excel=SimpleNamespace(Calculate=lambda: calculations.append(True)),
    )

    result = _write_ranges_in_session(
        session,
        sheet="Input",
        values={"A1": None, "B2:C3": [[1, 2], [3, 4]]},
        calculate=False,
    )

    assert ranges["A1"].Value is None
    assert ranges["B2:C3"].Value == ((1, 2), (3, 4))
    assert result["applied"] is True
    assert result["calculated"] is False
    assert calculations == []


@pytest.mark.unit
def test_range_writer_rejects_shape_mismatch_before_assignment():
    from xlvbatools.workbook.modifier import _write_ranges_in_session

    target = SimpleNamespace(
        Rows=SimpleNamespace(Count=2),
        Columns=SimpleNamespace(Count=2),
        Value="unchanged",
    )
    session = SimpleNamespace(
        wb=SimpleNamespace(
            Worksheets=lambda sheet: SimpleNamespace(Range=lambda address: target),
        ),
        excel=SimpleNamespace(Calculate=lambda: None),
    )

    with pytest.raises(ValueError, match="shape 2x1; range is 2x2"):
        _write_ranges_in_session(
            session, sheet="Input", values={"A1:B2": [[1], [2]]},
        )
    assert target.Value == "unchanged"


@pytest.mark.unit
def test_inspection_primitive_reuses_existing_session(monkeypatch):
    from xlvbatools.workbook import dumper

    session = object()
    seen = []
    monkeypatch.setattr(
        dumper,
        "export_screenshots",
        lambda *args, **kwargs: seen.append(
            ("render", kwargs["_session"], kwargs["expected_visible_content"])
        ) or {"Input": "Input.png"},
    )
    monkeypatch.setattr(
        dumper,
        "dump_sheet_data",
        lambda *args, **kwargs: seen.append(
            ("data", kwargs["_session"], None)
        ) or {
            "sheets": {
                "Input": {
                    "cells": {"A1": {"text": "ready", "value": "ready"}},
                    "shapes": [],
                },
            },
        },
    )

    result = dumper._inspect_existing_session(
        session,
        workbook_path="book.xlsm",
        sheets=["Input"],
        include_data=True,
        include_screenshots=True,
    )

    assert seen == [
        ("data", session, None),
        ("render", session, {"Input": 1}),
    ]
    assert result == {
        "workbook_data": {
            "sheets": {
                "Input": {
                    "cells": {"A1": {"text": "ready", "value": "ready"}},
                    "shapes": [],
                },
            },
        },
        "screenshots": {"Input": "Input.png"},
    }


@pytest.mark.unit
def test_inspection_step_preserves_render_mismatch_code(monkeypatch):
    from xlvbatools.core import workflow
    from xlvbatools.workbook import dumper

    error = dumper.ScreenshotRenderError(
        "blank native bitmap",
        details={"sheet": "Input", "range": "$A$1:$B$2"},
    )
    monkeypatch.setattr(
        dumper,
        "_inspect_existing_session",
        lambda *args, **kwargs: (_ for _ in ()).throw(error),
    )
    reporter = SimpleNamespace(workflow_step=lambda *args, **kwargs: None)

    with pytest.raises(workflow._StepFailure) as captured:
        workflow._inspect_step(
            object(),
            "book.xlsm",
            {
                "id": "inspect",
                "kind": "inspect",
                "sheets": ["Input"],
                "include_data": True,
                "include_screenshots": True,
            },
            reporter,
            index=1,
            count=1,
        )

    assert captured.value.code == "render_content_mismatch"
    assert captured.value.phase == "render_validation"
    assert captured.value.details == {"sheet": "Input", "range": "$A$1:$B$2"}
