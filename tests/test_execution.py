"""Typed executor boundary tests."""

import pytest


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
            "protocol_version": "2.0",
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
            "protocol_version": "2.0",
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
    assert result.elapsed_seconds == 1.5
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
            "protocol_version": "2.0",
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
