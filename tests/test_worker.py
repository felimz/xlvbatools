"""Contract tests for the shared isolated Excel worker protocol."""

import json
import os

import pytest


@pytest.mark.integration
def test_file_protocol_runs_offline_dry_run_in_separate_interpreter(tmp_path):
    from xlvbatools.core.worker import (
        WORKER_PROTOCOL_VERSION,
        run_isolated_operation,
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

    result = run_isolated_operation(
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
    assert result["excel_pid"] is None
    assert result["cleanup"] == {}
    assert result["data"][0]["status"] == "dry-run"


@pytest.mark.integration
def test_worker_setup_failure_is_structured(tmp_path):
    from xlvbatools.core.worker import run_isolated_operation

    result = run_isolated_operation(
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
def test_opt_in_transient_retry_starts_one_fresh_worker(monkeypatch):
    from xlvbatools.core import worker

    results = iter([
        {
            "success": False,
            "primary_error": "(-2147023174, 'The RPC server is unavailable.')",
            "cleanup": {"still_running": False},
        },
        {"success": True, "data": True, "cleanup": {"still_running": False}},
    ])
    calls = []

    def fake_once(operation, arguments, *, timeout):
        calls.append((operation, arguments, timeout))
        return next(results)

    monkeypatch.setattr(worker, "_run_isolated_operation_once", fake_once)
    result = worker.run_isolated_operation(
        "modify", {"workbook_path": "book.xlsm"}, timeout=9,
        retry_transient=True,
    )

    assert result["success"] is True
    assert result["attempt_count"] == 2
    assert len(calls) == 2


@pytest.mark.unit
def test_worker_mode_injection_raises_before_partial_save(tmp_path, monkeypatch):
    from types import SimpleNamespace
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
