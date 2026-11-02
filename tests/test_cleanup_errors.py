"""Cleanup diagnostics must not retain COM proxies through exception frames."""

import gc
import sys
from types import SimpleNamespace
import weakref

import pytest

from xlvbatools import CleanupReport
from xlvbatools.core.session import ExcelSession


pytestmark = pytest.mark.skipif(sys.platform != "win32", reason="Windows COM required")


@pytest.mark.unit
@pytest.mark.parametrize("stage,message", [
    ("sentinel", "sentinel preparation failed"),
    ("save", "workbook save failed"),
    ("close", "workbook close failed"),
    ("save", ""),
    ("close", ""),
])
def test_cleanup_errors_release_proxy_frames_before_quit(monkeypatch, stage, message):
    from xlvbatools.core import session as session_module

    proxies = []
    state = {"quit": False, "killed": False}
    retained_at_quit = []

    def quit_application(*_):
        retained_at_quit.extend(reference() is not None for reference in proxies)
        state["quit"] = True

    dispatch = SimpleNamespace(GetIDsOfNames=lambda name: 302, Invoke=quit_application)

    class Application:
        _oleobj_ = dispatch

        @property
        def Workbooks(self):
            if stage == "sentinel":
                # The raised exception's frame retains this Application proxy.
                raise RuntimeError(message)
            return SimpleNamespace(Add=lambda: SimpleNamespace(Saved=False))

    class Workbook:
        def Save(self):
            if stage == "save":
                raise RuntimeError(message)

        def Close(self, save):
            if stage == "close":
                raise RuntimeError(message)

    session = session_module.ExcelSession("dummy.xlsm", exit_grace_period=0)
    session.excel_pid = 4242
    session.excel = Application()
    session.wb = Workbook()
    proxies.extend((weakref.ref(session.excel), weakref.ref(session.wb)))

    def running(pid):
        return not state["killed"] and (
            not state["quit"] or any(reference() is not None for reference in proxies)
        )

    def terminate(pid):
        state["killed"] = True
        return True

    monkeypatch.setattr(session_module, "is_process_running", running)
    monkeypatch.setattr(session_module, "kill_process_by_pid", terminate)

    session.__exit__(None, None, None)

    assert retained_at_quit == [False, False], session.cleanup_result
    cleanup = CleanupReport.from_mapping(session.cleanup_result)
    assert cleanup.quit_requested and cleanup.exited_gracefully, session.cleanup_result
    assert not cleanup.force_terminated and not cleanup.still_running, session.cleanup_result
    assert state["killed"] is False
    if stage == "sentinel":
        assert cleanup.is_clean, session.cleanup_result
        assert session.cleanup_result["details"]["shutdown_sentinel_error"] == (
            "RuntimeError: " + message
        )
    else:
        assert session.cleanup_result[f"workbook_{stage}_error"] == message
        if message:
            assert not cleanup.is_clean, "A failed save/close must remain an unclean operation"


class _FaultingApplication:
    """Raise from a real COM access after materializing its native proxy."""

    def __init__(self, target, fault):
        object.__setattr__(self, "_target", target)
        object.__setattr__(self, "_fault", fault)

    def __getattr__(self, name):
        if name == "Workbooks" and self._fault == "workbooks":
            _ = self._target.Workbooks  # Retained by the injected traceback.
            raise RuntimeError("controlled shutdown preparation failure")
        return getattr(self._target, name)

    def __setattr__(self, name, value):
        setattr(self._target, name, value)
        if name == "DisplayAlerts" and self._fault == "display_alerts":
            raise RuntimeError("controlled shutdown preparation failure")


@pytest.mark.excel
@pytest.mark.parametrize("fault", ["workbooks", "display_alerts"])
def test_native_preparation_error_releases_application_before_quit(minimal_workbook, fault):
    from xlvbatools.core.process import is_process_running, kill_process_by_pid

    # Preserve the full default grace period and strict verdict. Disabling the
    # fallback lets a broken implementation release its traceback on return,
    # rather than killing Excel while that traceback still owns native proxies.
    session = ExcelSession(minimal_workbook, save_on_exit=False, terminate_owned_process=False)
    try:
        with session:
            session.excel = _FaultingApplication(session.excel, fault)
    finally:
        gc.collect()
        gc.collect()
        if session.excel_pid and is_process_running(session.excel_pid):
            session._wait_for_excel_exit(5)
            if is_process_running(session.excel_pid):
                assert kill_process_by_pid(session.excel_pid)

    # The finalizer above cannot turn an unclean context exit into a pass:
    # only the verdict recorded within __exit__ satisfies this assertion.
    assert CleanupReport.from_mapping(session.cleanup_result).is_clean, session.cleanup_result
    assert session.cleanup_result["details"]["shutdown_sentinel_error"] == (
        "RuntimeError: controlled shutdown preparation failure"
    )
    assert not is_process_running(session.excel_pid)
