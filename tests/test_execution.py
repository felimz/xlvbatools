"""Typed executor boundary tests."""

import pytest


def _worker_response(operation="extract", **overrides):
    response = {
        "protocol_version": "2.1",
        "request_id": "request-1",
        "operation": operation,
        "success": True,
        "phase": "complete",
    }
    response.update(overrides)
    return response


def _safe_startup_failure(**overrides):
    response = _worker_response(
        success=False,
        phase="worker_start",
        primary_error="Worker exited before session start",
        worker_output="startup stderr",
        error={
            "code": "worker_start_failed",
            "message": "Worker exited before session start",
        },
        worker_pid=100,
        excel_pid=None,
        dialog_events=[],
        cleanup={"pid": None, "still_running": False},
        worker_exit={
            "pid": 100,
            "exit_code": 1,
            "exited": True,
            "reaped": True,
            "force_terminated": False,
            "still_running": False,
        },
    )
    response.update(overrides)
    return response


@pytest.mark.unit
def test_operation_request_validates_and_freezes_arguments():
    from xlvbatools import Operation, OperationRequest

    source = {"workbook_path": "book.xlsm", "options": {"sheets": ["Input"]}}
    request = OperationRequest(Operation.EXTRACT, source, timeout=5)
    source["workbook_path"] = "changed.xlsm"
    source["options"]["sheets"].append("Results")

    assert request.arguments["workbook_path"] == "book.xlsm"
    assert request.arguments["options"]["sheets"] == ("Input",)
    with pytest.raises(TypeError):
        request.arguments["new"] = True
    with pytest.raises(TypeError):
        request.arguments["options"]["new"] = True
    with pytest.raises(ValueError, match="greater than zero"):
        OperationRequest(Operation.EXTRACT, {}, timeout=0)
    with pytest.raises(ValueError, match="only for modification"):
        OperationRequest(Operation.RUN, {}, retry_transient=True)
    with pytest.raises(TypeError, match="JSON-compatible"):
        OperationRequest(Operation.EXTRACT, {"path": object()})


@pytest.mark.unit
def test_executor_thaws_request_for_worker_transport(monkeypatch):
    from xlvbatools import IsolatedExecutor, Operation, OperationRequest
    from xlvbatools.core import worker

    captured = {}

    def execute(operation, arguments, **kwargs):
        captured.update(arguments)
        return {
            "protocol_version": "2.1",
            "request_id": "request-1",
            "operation": operation,
            "success": True,
            "phase": "complete",
        }

    monkeypatch.setattr(worker, "execute_worker_request", execute)
    request = OperationRequest(
        Operation.EXTRACT,
        {"options": {"sheets": ["Input"]}},
    )

    IsolatedExecutor().execute(request)

    assert captured == {"options": {"sheets": ["Input"]}}
    assert isinstance(captured["options"], dict)
    assert isinstance(captured["options"]["sheets"], list)


@pytest.mark.unit
def test_executor_converts_private_transport_to_public_result(monkeypatch):
    from xlvbatools import IsolatedExecutor, Operation, OperationRequest
    from xlvbatools.core import worker

    monkeypatch.setattr(
        worker,
        "execute_worker_request",
        lambda *args, **kwargs: {
            "protocol_version": "2.1",
            "request_id": "request-1",
            "operation": "extract",
            "success": True,
            "phase": "complete",
            "data": {"components": []},
            "worker_pid": 20,
            "excel_pid": 21,
            "elapsed_seconds": 1.5,
            "cleanup": {
                "pid": 21,
                "quit_requested": True,
                "exited_gracefully": True,
                "still_running": False,
            },
        },
    )

    result = IsolatedExecutor().execute(
        OperationRequest(Operation.EXTRACT, {"workbook_path": "book.xlsm"})
    )

    assert result.success is True
    assert result.data == {"components": []}
    assert result.request_id == "request-1"
    assert result.elapsed_seconds >= 0
    assert result.diagnostics.attempts[0].elapsed_seconds == 1.5
    assert result.require_clean_shutdown().pid == 21


@pytest.mark.unit
def test_executor_rejects_protocol_mismatch(monkeypatch):
    from xlvbatools import IsolatedExecutor, Operation, OperationRequest
    from xlvbatools.core import worker

    monkeypatch.setattr(
        worker,
        "execute_worker_request",
        lambda *args, **kwargs: {
            "protocol_version": "1.0",
            "request_id": "request-1",
            "operation": "extract",
            "success": True,
        },
    )

    result = IsolatedExecutor().execute(
        OperationRequest(Operation.EXTRACT, {"workbook_path": "book.xlsm"})
    )

    assert result.success is False
    assert result.phase == "transport"
    assert result.error.code == "protocol_mismatch"


@pytest.mark.unit
def test_executor_preserves_structured_worker_failure_details(monkeypatch):
    from xlvbatools import IsolatedExecutor, Operation, OperationRequest
    from xlvbatools.core import worker

    monkeypatch.setattr(
        worker,
        "execute_worker_request",
        lambda *args, **kwargs: {
            "protocol_version": "2.1",
            "request_id": "request-1",
            "operation": "extract",
            "success": False,
            "phase": "worker_error",
            "primary_error": "boom",
            "error_type": "RuntimeError",
            "traceback": "trace text",
            "worker_output": "worker log",
        },
    )

    result = IsolatedExecutor().execute(OperationRequest(Operation.EXTRACT, {}))

    assert result.error.message == "boom"
    assert result.error.error_type == "RuntimeError"
    assert result.error.details == {
        "traceback": "trace text",
        "worker_output": "worker log",
    }


@pytest.mark.unit
def test_executor_preserves_structured_workflow_timeout_progress(monkeypatch):
    from xlvbatools import IsolatedExecutor, Operation, OperationRequest
    from xlvbatools.core import worker

    calls = []
    progress = {
        "phase": "workflow_step",
        "step_id": "calculate",
        "step_kind": "macro",
        "step_index": 2,
        "step_count": 4,
        "step_phase": "macro_execution",
    }

    def execute(*args, **kwargs):
        calls.append(1)
        return _worker_response(
            operation="workflow",
            success=False,
            phase="workflow_step",
            primary_error="workflow exceeded 240 seconds",
            timed_out=True,
            progress=progress,
        )

    monkeypatch.setattr(worker, "execute_worker_request", execute)
    result = IsolatedExecutor().execute(OperationRequest(Operation.WORKFLOW, {}))

    assert result.success is False
    assert result.error.code == "timeout"
    assert result.error.details["progress"]["step_id"] == "calculate"
    assert result.diagnostics.progress["step_phase"] == "macro_execution"
    assert result.attempt_count == 1
    assert len(calls) == 1


@pytest.mark.unit
def test_executor_allows_only_pre_session_start_retry_for_workflow(monkeypatch):
    from xlvbatools import IsolatedExecutor, Operation, OperationRequest
    from xlvbatools.core import worker

    responses = iter([
        _safe_startup_failure(operation="workflow"),
        _worker_response(operation="workflow", request_id="request-2", data={"steps": []}),
    ])
    monkeypatch.setattr(
        worker, "execute_worker_request", lambda *args, **kwargs: next(responses),
    )

    result = IsolatedExecutor().execute(OperationRequest(Operation.WORKFLOW, {}))

    assert result.success is True
    assert result.attempt_count == 2
    assert result.diagnostics.attempts[0].retry_reason == (
        "worker_exited_before_session_start"
    )


@pytest.mark.unit
def test_executor_retries_worker_creation_failure_once(monkeypatch):
    from xlvbatools import IsolatedExecutor, Operation, OperationRequest
    from xlvbatools.core import worker

    responses = iter([
        worker.WorkerCreationError("unable to create worker"),
        _worker_response(request_id="request-2"),
    ])
    calls = []

    def execute(*args, **kwargs):
        calls.append(kwargs["timeout"])
        response = next(responses)
        if isinstance(response, Exception):
            raise response
        return response

    monkeypatch.setattr(worker, "execute_worker_request", execute)
    result = IsolatedExecutor().execute(OperationRequest(Operation.EXTRACT, {}))

    assert result.success is True
    assert len(calls) == 2
    assert result.attempt_count == 2
    assert result.diagnostics.attempts[0].retry_reason == "worker_creation_failed"
    assert result.diagnostics.attempts[1].retryable is False


@pytest.mark.unit
def test_executor_retries_proven_reaped_pre_session_worker_exit(monkeypatch):
    from xlvbatools import IsolatedExecutor, Operation, OperationRequest
    from xlvbatools.core import worker

    responses = iter([_safe_startup_failure(), _worker_response(request_id="request-2")])
    monkeypatch.setattr(
        worker, "execute_worker_request", lambda *args, **kwargs: next(responses),
    )

    result = IsolatedExecutor().execute(OperationRequest(Operation.EXTRACT, {}))

    assert result.success is True
    assert result.attempt_count == 2
    first = result.diagnostics.attempts[0]
    assert first.retry_reason == "worker_exited_before_session_start"
    assert first.worker is not None and first.worker.is_clean
    assert first.error_details["worker_output"] == "startup stderr"
    assert first.cleanup is not None and first.cleanup.still_running is False


@pytest.mark.unit
def test_executor_never_exceeds_two_total_attempts(monkeypatch):
    from xlvbatools import IsolatedExecutor, Operation, OperationRequest
    from xlvbatools.core import worker

    responses = iter([
        _safe_startup_failure(operation="modify"),
        _worker_response(
            operation="modify",
            success=False,
            phase="modify",
            primary_error="0x800706ba RPC server is unavailable",
        ),
    ])
    calls = []

    def execute(*args, **kwargs):
        calls.append(args[0])
        return next(responses)

    monkeypatch.setattr(worker, "execute_worker_request", execute)
    result = IsolatedExecutor().execute(OperationRequest(
        Operation.MODIFY, {}, retry_transient=True,
    ))

    assert result.success is False
    assert len(calls) == 2
    assert result.attempt_count == 2


@pytest.mark.unit
@pytest.mark.parametrize(
    "overrides",
    [
        {"phase": "session_start"},
        {"excel_pid": 200},
        {"dialog_events": [{"title": "Excel"}]},
        {"timed_out": True},
        {"worker_exit": None},
        {"worker_exit": {
            "pid": 100, "exit_code": None, "exited": False, "reaped": False,
            "force_terminated": False, "still_running": True,
        }},
        {"cleanup": {"pid": None, "worker_terminated": True}},
    ],
)
def test_executor_does_not_retry_ambiguous_or_post_ownership_failures(
    monkeypatch, overrides,
):
    from xlvbatools import IsolatedExecutor, Operation, OperationRequest
    from xlvbatools.core import worker

    calls = []

    def execute(*args, **kwargs):
        calls.append(1)
        return _safe_startup_failure(**overrides)

    monkeypatch.setattr(worker, "execute_worker_request", execute)
    result = IsolatedExecutor().execute(OperationRequest(Operation.EXTRACT, {}))

    assert result.success is False
    assert len(calls) == 1
    assert result.attempt_count == 1
    assert result.diagnostics.attempts[0].retryable is False


@pytest.mark.unit
def test_executor_does_not_retry_protocol_failure(monkeypatch):
    from xlvbatools import IsolatedExecutor, Operation, OperationRequest
    from xlvbatools.core import worker

    calls = []

    def execute(*args, **kwargs):
        calls.append(1)
        return _worker_response(protocol_version="1.0")

    monkeypatch.setattr(worker, "execute_worker_request", execute)
    result = IsolatedExecutor().execute(OperationRequest(Operation.EXTRACT, {}))

    assert result.error.code == "protocol_mismatch"
    assert len(calls) == 1


@pytest.mark.unit
def test_executor_transient_retry_is_limited_to_opted_in_modification(monkeypatch):
    from xlvbatools import IsolatedExecutor, Operation, OperationRequest
    from xlvbatools.core import worker

    failure = _worker_response(
        operation="modify",
        success=False,
        phase="modify",
        primary_error="0x800706ba RPC server is unavailable",
        cleanup={"pid": 201, "still_running": False},
        worker_exit={
            "pid": 101,
            "exit_code": 1,
            "exited": True,
            "reaped": True,
            "force_terminated": False,
            "still_running": False,
        },
    )
    responses = iter([failure, _worker_response(operation="modify")])
    monkeypatch.setattr(
        worker, "execute_worker_request", lambda *args, **kwargs: next(responses),
    )

    result = IsolatedExecutor().execute(OperationRequest(
        Operation.MODIFY, {}, retry_transient=True,
    ))

    assert result.success is True
    assert result.attempt_count == 2
    assert result.diagnostics.attempts[0].retry_reason == (
        "transient_modify_com_failure"
    )


@pytest.mark.unit
@pytest.mark.parametrize(
    "evidence",
    [
        {},
        {"cleanup": {"pid": 201, "still_running": False}},
        {
            "cleanup": {"pid": 201, "still_running": True},
            "worker_exit": {
                "pid": 101, "exit_code": 1, "exited": True, "reaped": True,
                "force_terminated": False, "still_running": False,
            },
        },
    ],
)
def test_executor_does_not_retry_transient_modify_without_exit_proof(
    monkeypatch, evidence,
):
    from xlvbatools import IsolatedExecutor, Operation, OperationRequest
    from xlvbatools.core import worker

    calls = []

    def execute(*args, **kwargs):
        calls.append(1)
        return _worker_response(
            operation="modify",
            success=False,
            phase="modify",
            primary_error="0x800706ba RPC server is unavailable",
            **evidence,
        )

    monkeypatch.setattr(worker, "execute_worker_request", execute)
    result = IsolatedExecutor().execute(OperationRequest(
        Operation.MODIFY, {}, retry_transient=True,
    ))

    assert result.success is False
    assert len(calls) == 1
    assert result.attempt_count == 1


@pytest.mark.unit
def test_executor_uses_one_overall_timeout_budget(monkeypatch):
    from xlvbatools import IsolatedExecutor, Operation, OperationRequest
    from xlvbatools.core import worker
    from xlvbatools import execution

    ticks = iter([0.0, 0.0, 0.0, 3.0, 4.0])
    monkeypatch.setattr(execution.time, "monotonic", lambda: next(ticks))
    timeouts = []
    responses = iter([_safe_startup_failure(), _worker_response(request_id="request-2")])

    def execute(*args, **kwargs):
        timeouts.append(kwargs["timeout"])
        return next(responses)

    monkeypatch.setattr(worker, "execute_worker_request", execute)
    result = IsolatedExecutor().execute(OperationRequest(
        Operation.EXTRACT, {}, timeout=10,
    ))

    assert result.success is True
    assert timeouts == [10.0, 7.0]
    assert result.elapsed_seconds == 4.0


@pytest.mark.unit
def test_executor_does_not_misclassify_unproven_transport_errors(monkeypatch):
    from xlvbatools import IsolatedExecutor, Operation, OperationRequest
    from xlvbatools.core import worker

    calls = []

    def execute(*args, **kwargs):
        calls.append(1)
        raise OSError("parent transport failed after unknown state")

    monkeypatch.setattr(worker, "execute_worker_request", execute)
    result = IsolatedExecutor().execute(OperationRequest(Operation.EXTRACT, {}))

    assert result.success is False
    assert result.error.code == "worker_transport_failed"
    assert result.attempt_count == 1
    assert len(calls) == 1
