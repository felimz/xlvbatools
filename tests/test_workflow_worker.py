"""Offline execution tests for the one-session workflow worker."""

from __future__ import annotations

import pytest


class FakeReporter:
    def __init__(self):
        self.value = {"phase": "session_start"}
        self.events = []

    def phase(self, phase):
        self.value["phase"] = phase
        self.events.append((phase, None))

    def excel_started(self, pid):
        self.value.update(phase="workbook_open", excel_pid=pid)

    def workflow_step(self, step, *, index, count, step_phase):
        self.value.update({
            "phase": "workflow_step",
            "step_id": step["id"],
            "step_kind": step["kind"],
            "step_index": index,
            "step_count": count,
            "step_phase": step_phase,
        })
        self.events.append((step["id"], step_phase))


class FakeSession:
    instances = []
    macro_results = {}
    save_error = None

    def __init__(self, workbook_path, **kwargs):
        self.workbook_path = workbook_path
        self.kwargs = kwargs
        self.excel_pid = 321
        self.dialog_events = []
        self.cleanup_result = {
            "pid": 321,
            "quit_requested": True,
            "exited_gracefully": True,
            "force_terminated": False,
            "still_running": False,
        }
        self.phase = "session_start"
        self.named_ranges = []
        self.macros = []
        self.saved = False
        self.exit_exc_type = None
        type(self).instances.append(self)

    def __enter__(self):
        self.phase = "ready"
        self.kwargs["on_excel_started"](self.excel_pid)
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.exit_exc_type = exc_type
        return False

    def set_named_range(self, name, value, strict=False):
        self.named_ranges.append((name, value, strict))
        return True

    def run_macro(self, macro):
        self.macros.append(macro)
        return dict(type(self).macro_results.get(macro) or {
            "success": True,
            "phase": "macro_execution",
            "macro": macro,
            "run_id": f"run-{macro}",
            "elapsed_seconds": 0.1,
            "dialog_events": [],
            "cleanup": self.cleanup_result,
        })

    def save(self):
        if type(self).save_error is not None:
            raise type(self).save_error
        self.saved = True


@pytest.fixture(autouse=True)
def reset_fake_session():
    FakeSession.instances = []
    FakeSession.macro_results = {}
    FakeSession.save_error = None


def _arguments(steps, *, save=False):
    return {
        "workflow_schema_version": "1.0",
        "workbook_path": "book.xlsm",
        "steps": steps,
        "visible": False,
        "save_on_success": save,
    }


@pytest.mark.unit
def test_workflow_executes_all_steps_in_one_session_and_saves_last(monkeypatch):
    from xlvbatools.core import workflow
    from xlvbatools.workbook import dumper, modifier

    monkeypatch.setattr(workflow, "ExcelSession", FakeSession)
    modify_calls = []
    inspect_calls = []
    monkeypatch.setattr(
        modifier,
        "_write_ranges_in_session",
        lambda session, **kwargs: modify_calls.append((session, kwargs)) or {
            "applied": True,
            "writes": [{"sheet": "Input", "range": "A1", "rows": 1, "columns": 1}],
            "calculated": False,
        },
    )
    monkeypatch.setattr(
        dumper,
        "_inspect_existing_session",
        lambda session, **kwargs: inspect_calls.append((session, kwargs)) or {
            "workbook_data": {"sheets": {}}, "screenshots": {},
        },
    )
    reporter = FakeReporter()

    result = workflow.execute_workflow(_arguments([
        {
            "id": "retrieve", "kind": "macro", "macro": "OnRetrieve",
            "named_ranges": {"FilePath": "model.r3d"},
        },
        {
            "id": "inputs", "kind": "modify", "sheet": "Input",
            "values": {"A1": 42}, "calculate": False,
        },
        {
            "id": "results", "kind": "inspect", "sheets": ["Input"],
            "include_data": True, "include_screenshots": False,
        },
    ], save=True), reporter)

    assert result["success"] is True
    assert result["phase"] == "complete"
    assert len(FakeSession.instances) == 1
    session = FakeSession.instances[0]
    assert modify_calls[0][0] is session
    assert inspect_calls[0][0] is session
    assert session.macros == ["OnRetrieve"]
    assert session.named_ranges == [("FilePath", "model.r3d", True)]
    assert session.saved is True
    assert session.exit_exc_type is None
    assert [item["status"] for item in result["data"]["steps"]] == [
        "succeeded", "succeeded", "succeeded",
    ]
    assert result["data"]["saved"] is True


@pytest.mark.unit
def test_workflow_failure_is_fail_fast_not_saved_and_not_replayed(monkeypatch):
    from xlvbatools.core import workflow

    FakeSession.macro_results["Fail"] = {
        "success": False,
        "phase": "macro_execution",
        "macro": "Fail",
        "run_id": "run-fail",
        "primary_error": "calculation failed",
        "dialog_events": [],
        "cleanup": {},
    }
    monkeypatch.setattr(workflow, "ExcelSession", FakeSession)
    reporter = FakeReporter()

    result = workflow.execute_workflow(_arguments([
        {"id": "retrieve", "kind": "macro", "macro": "OnRetrieve"},
        {"id": "calculate", "kind": "macro", "macro": "Fail"},
        {"id": "never", "kind": "macro", "macro": "MustNotRun"},
    ], save=True), reporter)

    session = FakeSession.instances[0]
    assert result["success"] is False
    assert result["error"]["code"] == "workflow_step_failed"
    assert result["data"]["failed_step_id"] == "calculate"
    assert [item["status"] for item in result["data"]["steps"]] == [
        "succeeded", "failed", "not_run",
    ]
    assert session.macros == ["OnRetrieve", "Fail"]
    assert session.saved is False
    assert session.exit_exc_type is not None
    assert result["data"]["saved"] is False


@pytest.mark.unit
def test_workflow_save_failure_is_outer_failure_after_successful_steps(monkeypatch):
    from xlvbatools.core import workflow

    FakeSession.save_error = OSError("disk full")
    monkeypatch.setattr(workflow, "ExcelSession", FakeSession)

    result = workflow.execute_workflow(_arguments([
        {"id": "calculate", "kind": "macro", "macro": "OnCalculate"},
    ], save=True), FakeReporter())

    assert result["success"] is False
    assert result["phase"] == "workbook_save"
    assert result["error"]["code"] == "workbook_save_failed"
    assert result["data"]["failed_step_id"] is None
    assert result["data"]["steps"][0]["status"] == "succeeded"
    assert result["cleanup"]["workbook_save_error"] == "disk full"


@pytest.mark.unit
def test_workflow_rejects_schema_before_constructing_session(monkeypatch):
    from xlvbatools.core import workflow

    monkeypatch.setattr(workflow, "ExcelSession", FakeSession)

    with pytest.raises(ValueError, match="Unsupported workflow schema"):
        workflow.execute_workflow(
            {
                "workflow_schema_version": "99.0",
                "workbook_path": "book.xlsm",
                "steps": [{"id": "one", "kind": "macro", "macro": "One"}],
            },
            FakeReporter(),
        )
    assert FakeSession.instances == []


@pytest.mark.unit
@pytest.mark.parametrize(
    "override, match",
    [
        ({"visible": "false"}, "visible must be boolean"),
        ({"save_on_success": "false"}, "save_on_success must be boolean"),
        ({"unexpected": True}, "unknown workflow request fields"),
    ],
)
def test_workflow_rejects_invalid_top_level_controls_before_session(
    monkeypatch, override, match,
):
    from xlvbatools.core import workflow

    monkeypatch.setattr(workflow, "ExcelSession", FakeSession)
    arguments = _arguments([{"id": "one", "kind": "macro", "macro": "One"}])
    arguments.update(override)

    with pytest.raises((TypeError, ValueError), match=match):
        workflow.execute_workflow(arguments, FakeReporter())
    assert FakeSession.instances == []


@pytest.mark.unit
def test_workflow_session_failure_marks_every_step_not_run(monkeypatch):
    from xlvbatools.core import workflow

    class FailingSession(FakeSession):
        def __enter__(self):
            self.phase = "workbook_open"
            raise RuntimeError("cannot open workbook")

    monkeypatch.setattr(workflow, "ExcelSession", FailingSession)
    result = workflow.execute_workflow(_arguments([
        {"id": "one", "kind": "macro", "macro": "One"},
        {"id": "two", "kind": "macro", "macro": "Two"},
    ]), FakeReporter())

    assert result["success"] is False
    assert result["phase"] == "workbook_open"
    assert result["error"]["code"] == "workflow_session_failed"
    assert [step["status"] for step in result["data"]["steps"]] == [
        "not_run", "not_run",
    ]
