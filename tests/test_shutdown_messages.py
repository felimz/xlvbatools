"""Native process exit that depends on a callback to the waiting STA thread."""

import subprocess
import sys
import uuid
from types import SimpleNamespace

import pytest


pytestmark = [
    pytest.mark.integration,
    pytest.mark.skipif(sys.platform != "win32", reason="Windows message queue required"),
]


def test_process_can_finish_exit_via_waiting_thread_callback(tmp_path):
    import pythoncom
    import win32api
    import win32con
    import win32gui

    from xlvbatools import CleanupReport
    from xlvbatools.core.session import ExcelSession

    # pythoncom initializes the main STA. Session cleanup must keep that
    # caller-owned apartment alive and service its messages while waiting.
    gate = tmp_path / "callback-completed"
    callback_message = win32con.WM_APP + 1
    callbacks = []

    def window_proc(hwnd, message, wparam, lparam):
        if message == callback_message:
            gate.write_text("exit", encoding="ascii")
            callbacks.append(message)
            return 0
        return win32gui.DefWindowProc(hwnd, message, wparam, lparam)

    class_name = "xlvba-shutdown-test-" + uuid.uuid4().hex
    instance = win32api.GetModuleHandle(None)
    window_class = win32gui.WNDCLASS()
    window_class.hInstance = instance
    window_class.lpszClassName = class_name
    window_class.lpfnWndProc = window_proc
    win32gui.RegisterClass(window_class)
    hwnd = win32gui.CreateWindow(
        class_name, "", 0, 0, 0, 0, 0, -3, 0, instance, None,
    )  # HWND_MESSAGE: no visible desktop window.
    child = None
    pythoncom.CoInitialize()
    try:
        child = subprocess.Popen(
            [
                getattr(sys, "_base_executable", sys.executable), "-c",
                "import pathlib,sys,time; gate=pathlib.Path(sys.argv[1]); "
                "\nwhile not gate.exists(): time.sleep(0.01)",
                str(gate),
            ],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        session = ExcelSession("dummy.xlsm", save_on_exit=False, exit_grace_period=5)
        session.excel_pid = child.pid
        session.excel = SimpleNamespace(
            Workbooks=SimpleNamespace(Add=lambda: SimpleNamespace(Saved=True)),
            _oleobj_=SimpleNamespace(
                GetIDsOfNames=lambda name: 302,
                Invoke=lambda *args: win32gui.PostMessage(hwnd, callback_message, 0, 0),
            ),
        )
        session.wb = SimpleNamespace(Close=lambda value: None)

        session.__exit__(None, None, None)

        assert CleanupReport.from_mapping(session.cleanup_result).is_clean, session.cleanup_result
        assert callbacks == [callback_message]
        assert child.wait(timeout=5) == 0
    finally:
        # Only the child owned by this test may be stopped.
        gate.write_text("exit", encoding="ascii")
        if child is not None:
            try:
                child.wait(timeout=5)
            except subprocess.TimeoutExpired:
                child.kill()
                child.wait(timeout=5)
        win32gui.DestroyWindow(hwnd)
        win32gui.UnregisterClass(class_name, instance)
        pythoncom.CoUninitialize()
