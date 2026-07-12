"""Unit tests for structured macro orchestration."""

import pytest


@pytest.mark.unit
def test_strict_setup_failure_does_not_run_macro(monkeypatch):
    from xlvbatools.macro import runner

    calls = []

    class FakeSession:
        def __init__(self, *args, **kwargs):
            self.phase = "ready"
            self.cleanup_result = {"still_running": False}
            self.dialog_events = []

        def __enter__(self):
            return self

        def __exit__(self, *args):
            return False

        def set_named_range(self, name, value, strict=False):
            assert strict is True
            raise KeyError(name)

        def run_macro(self, macro_name, timeout=120):
            calls.append(macro_name)

    monkeypatch.setattr(runner, "ExcelSession", FakeSession)
    result = runner._run_macro_once(
        "book.xlsm", "NeverRuns", {"Missing": 1}, 10, False, True, True, "run-1"
    )

    assert result["success"] is False
    assert result["phase"] == "named_range_setup"
    assert result["macro"] == "NeverRuns"
    assert calls == []


@pytest.mark.unit
def test_timeout_requires_positive_value():
    from xlvbatools.macro.runner import run_macro
    with pytest.raises(ValueError, match="greater than zero"):
        run_macro("book.xlsm", "Macro", timeout=0)


@pytest.mark.unit
def test_targeted_timeout_cleanup(monkeypatch):
    from xlvbatools.macro import runner

    killed = []

    class Process:
        alive = True

        def join(self, timeout):
            pass

        def is_alive(self):
            return self.alive

        def terminate(self):
            self.alive = False

    states = iter([True, False])
    monkeypatch.setattr(runner, "is_process_running", lambda pid: next(states, False))
    monkeypatch.setattr(runner, "kill_process_by_pid", lambda pid: killed.append(pid) or True)
    cleanup = runner._terminate_timed_out_run(Process(), 4321, grace_period=0)

    assert killed == [4321]
    assert cleanup["pid"] == 4321
    assert cleanup["force_terminated"] is True
    assert cleanup["worker_terminated"] is True
    assert cleanup["still_running"] is False


@pytest.mark.com
@pytest.mark.integration
def test_worker_macro_completes(runtime_error_workbook):
    from xlvbatools.macro.runner import run_macro

    result = run_macro(runtime_error_workbook, "CompleteNormally", timeout=20, save_on_exit=False)
    assert result["success"] is True
    assert result["cleanup"]["still_running"] is False


@pytest.mark.com
@pytest.mark.integration
def test_worker_returns_multiline_runtime_error(runtime_error_workbook):
    from xlvbatools.macro.runner import run_macro

    result = run_macro(runtime_error_workbook, "RaiseMultilineError", timeout=20, save_on_exit=False)
    assert result["success"] is False
    assert result["phase"] == "macro_execution"
    assert "Diagnostic line one." in result["primary_error"]
    assert "Diagnostic line two." in result["primary_error"]
    assert result["cleanup"]["still_running"] is False


@pytest.mark.com
@pytest.mark.integration
def test_worker_enforces_infinite_loop_timeout(runtime_error_workbook):
    from xlvbatools.macro.runner import run_macro

    result = run_macro(runtime_error_workbook, "LoopForever", timeout=6, save_on_exit=False)
    assert result["success"] is False
    assert result["timed_out"] is True
    assert result["phase"] == "macro_execution"
    assert result["excel_pid"] is not None
    assert result["cleanup"]["still_running"] is False


@pytest.mark.com
@pytest.mark.integration
@pytest.mark.parametrize("macro_name", ["ShowMessage", "ShowFilePicker"])
def test_worker_dismisses_modal_ui(runtime_error_workbook, macro_name):
    from xlvbatools.macro.runner import run_macro

    result = run_macro(runtime_error_workbook, macro_name, timeout=20, save_on_exit=False)
    assert result.get("timed_out", False) is False
    assert result["success"] is True
    assert result["dialog_events"]
    assert result["dialog_events"][0]["dismissed"] is True
    assert result["cleanup"]["still_running"] is False


@pytest.mark.com
@pytest.mark.integration
def test_timeout_preserves_unrelated_excel(runtime_error_workbook):
    import win32com.client
    import win32process
    from xlvbatools.core.process import is_process_running
    from xlvbatools.macro.runner import run_macro

    unrelated_excel = win32com.client.DispatchEx("Excel.Application")
    unrelated_excel.Visible = False
    unrelated_wb = unrelated_excel.Workbooks.Add()
    _, unrelated_pid = win32process.GetWindowThreadProcessId(unrelated_excel.Hwnd)
    try:
        result = run_macro(runtime_error_workbook, "LoopForever", timeout=6, save_on_exit=False)
        assert result["timed_out"] is True
        assert result["excel_pid"] != unrelated_pid
        assert is_process_running(unrelated_pid)
        assert unrelated_wb.Name
    finally:
        unrelated_wb.Close(False)
        unrelated_excel.Quit()
