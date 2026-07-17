"""Contract tests for the shared isolated Excel worker protocol."""

import json
import os
from types import SimpleNamespace

import pytest


@pytest.mark.unit
def test_macro_worker_payload_keeps_operation_data_inside_envelope(monkeypatch):
    from xlvbatools.core.worker_entry import _dispatch
    from xlvbatools.macro import runner

    monkeypatch.setattr(
        runner,
        "_run_macro_once",
        lambda **kwargs: {
            "success": True,
            "phase": "complete",
            "run_id": "run-123",
            "return_value": 42,
            "dialog_events": [],
            "cleanup": {"still_running": False},
        },
    )
    reporter = SimpleNamespace(
        phase=lambda value: None,
        excel_started=lambda pid: None,
    )

    result = _dispatch("run_macro", {}, reporter)

    assert result["data"] == {"run_id": "run-123", "return_value": 42}
    assert "run_id" not in result
    assert "return_value" not in result


@pytest.mark.integration
def test_file_protocol_runs_offline_dry_run_in_separate_interpreter(tmp_path):
    from xlvbatools.core.worker import (
        WORKER_PROTOCOL_VERSION,
        execute_worker_request,
    )

    source = tmp_path / "vba_source"
    source.mkdir()
    (source / "manifest.json").write_text(
        json.dumps({
            "components": [{
                "name": "modExample",
                "type_code": 1,
                "type_name": "standard_module",
                "file": "modules/modExample.bas",
                "line_count": 3,
                "sha256": "abc",
            }],
        }),
        encoding="utf-8",
    )

    result = execute_worker_request(
        "inject",
        {
            "workbook_path": str(tmp_path / "not-opened.xlsm"),
            "source_dir": str(source),
            "component": None,
            "backup": False,
            "dry_run": True,
            "backup_limit": 5,
        },
        timeout=15,
    )

    assert result["success"] is True, result
    assert result["protocol_version"] == WORKER_PROTOCOL_VERSION
    assert result["request_id"]
    assert result["worker_pid"] != os.getpid()
    assert result["worker_pid"] == result["executor_pid"]
    assert result["worker_exit"] == {
        "pid": result["worker_pid"],
        "exit_code": 0,
        "exited": True,
        "reaped": True,
        "force_terminated": False,
        "still_running": False,
    }
    assert result["excel_pid"] is None
    assert result["cleanup"] == {}
    assert result["data"][0]["status"] == "dry-run"


@pytest.mark.integration
def test_worker_setup_failure_is_structured(tmp_path):
    from xlvbatools.core.worker import execute_worker_request

    result = execute_worker_request(
        "extract", {"workbook_path": str(tmp_path / "missing.xlsm")},
        timeout=15,
    )

    assert result["success"] is False
    assert result["operation"] == "extract"
    assert result["worker_pid"] != os.getpid()
    assert result["primary_error"]
    assert result["cleanup"]["pid"] is None
    assert result["cleanup"]["still_running"] is False


@pytest.mark.unit
def test_worker_transport_executes_exactly_one_attempt(monkeypatch):
    from xlvbatools.core import worker

    calls = []

    def fake_once(operation, arguments, *, timeout):
        calls.append((operation, arguments, timeout))
        return {
            "success": False,
            "primary_error": "(-2147023174, 'The RPC server is unavailable.')",
            "cleanup": {"still_running": False},
        }

    monkeypatch.setattr(worker, "_execute_worker_request_once", fake_once)
    result = worker.execute_worker_request(
        "modify", {"workbook_path": "book.xlsm"}, timeout=9,
    )

    assert result["success"] is False
    assert len(calls) == 1


@pytest.mark.unit
def test_worker_wraps_only_process_creation_failure(monkeypatch):
    from xlvbatools.core import worker

    def fail_creation(*args, **kwargs):
        raise OSError("CreateProcess failed")

    monkeypatch.setattr(worker.subprocess, "Popen", fail_creation)

    with pytest.raises(worker.WorkerCreationError, match="CreateProcess failed"):
        worker.execute_worker_request("extract", {}, timeout=1)


@pytest.mark.unit
def test_session_start_is_published_before_excel_session_construction(monkeypatch):
    from xlvbatools.core import session, worker_entry

    phases = []

    class SentinelError(RuntimeError):
        pass

    class FakeExcelSession:
        def __init__(self, *args, **kwargs):
            assert phases == ["session_start"]
            raise SentinelError("stop before COM")

    monkeypatch.setattr(session, "ExcelSession", FakeExcelSession)
    reporter = SimpleNamespace(phase=phases.append, excel_started=lambda pid: None)

    with pytest.raises(SentinelError, match="stop before COM"):
        worker_entry._session_result(
            "list_components",
            {"workbook_path": "book.xlsm"},
            reporter,
        )

    assert phases == ["session_start"]


@pytest.mark.unit
def test_worker_mode_injection_raises_before_partial_save(tmp_path, monkeypatch):
    from xlvbatools.vba import injector

    source = tmp_path / "vba_source"
    source.mkdir()
    (source / "manifest.json").write_text(
        json.dumps({
            "components": [{
                "name": "modBroken",
                "type_code": 1,
                "type_name": "standard_module",
                "file": "modules/modBroken.bas",
            }],
        }),
        encoding="utf-8",
    )
    (source / "modules").mkdir()
    (source / "modules" / "modBroken.bas").write_text(
        "Option Explicit\n", encoding="utf-8",
    )
    monkeypatch.setattr(
        injector, "_inject_single",
        lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError("broken import")),
    )

    with pytest.raises(RuntimeError, match="modBroken: broken import"):
        injector.inject_all(
            str(tmp_path / "book.xlsm"),
            str(source),
            backup=False,
            _session=SimpleNamespace(vb_project=object()),
            _raise_on_error=True,
        )
