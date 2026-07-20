"""
Tests for xlvbatools.core.session -- ExcelSession context manager.
"""

import gc
import subprocess
import sys
import textwrap

import pytest
from types import SimpleNamespace


def _vbe_is_visible(excel) -> bool:
    """Read VBE visibility without retaining a nested COM proxy."""
    vbe = None
    main_window = None
    try:
        vbe = excel.VBE
        main_window = vbe.MainWindow
        return bool(main_window.Visible)
    finally:
        main_window = None
        vbe = None
        gc.collect()
        gc.collect()


@pytest.mark.unit
class TestSessionImport:
    """ExcelSession is an internal implementation detail, not public API."""

    def test_session_is_not_exported_from_package(self):
        import xlvbatools

        assert "ExcelSession" not in xlvbatools.__all__
        assert not hasattr(xlvbatools, "ExcelSession")

    def test_import_from_module(self):
        from xlvbatools.core.session import ExcelSession
        assert ExcelSession is not None

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
    def test_missing_workbook_raises(self):
        from xlvbatools.core.session import ExcelSession
        with pytest.raises(FileNotFoundError, match="Workbook not found"):
            with ExcelSession("nonexistent_workbook.xlsm", kill_on_enter=False):
                pass


@pytest.mark.unit
class TestSessionProperties:
    """Test session property defaults before context entry."""

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
    def test_initial_state(self):
        from xlvbatools.core.session import ExcelSession
        session = ExcelSession("dummy.xlsm")
        assert session.excel is None
        assert session.wb is None
        assert session.had_dialogs is False
        assert session.had_errors is False
        assert session.error_summary == ""
        assert session.dialog_events == []

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
    def test_strict_named_range_raises(self):
        from xlvbatools.core.session import ExcelSession
        session = ExcelSession("dummy.xlsm")
        session.wb = SimpleNamespace(Names=lambda name: (_ for _ in ()).throw(RuntimeError("missing")))
        with pytest.raises(KeyError, match="Could not set named range"):
            session.set_named_range("MissingName", 1, strict=True)

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
    def test_exit_force_terminates_only_owned_pid(self, monkeypatch):
        from xlvbatools.core import session as session_module

        owned_pid = 4242
        killed = []
        states = iter([True, True, False])
        monkeypatch.setattr(session_module, "is_process_running", lambda pid: next(states, False))
        monkeypatch.setattr(session_module, "kill_process_by_pid", lambda pid: killed.append(pid) or True)

        session = session_module.ExcelSession(
            "dummy.xlsm", exit_grace_period=0, terminate_owned_process=True
        )
        session.excel_pid = owned_pid
        session.excel = SimpleNamespace(Quit=lambda: None)
        session.wb = SimpleNamespace(Close=lambda value: None)
        session.__exit__(None, None, None)

        assert killed == [owned_pid]
        assert session.cleanup_result["force_terminated"] is True
        assert session.cleanup_result["still_running"] is False

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
    def test_exit_releases_owned_com_apartment_before_waiting_for_pid(
        self, monkeypatch,
    ):
        from xlvbatools.core import session as session_module

        calls = []
        session = session_module.ExcelSession("dummy.xlsm")
        session.excel_pid = 4242
        session.excel = SimpleNamespace(Quit=lambda: None)
        session.wb = SimpleNamespace(Close=lambda value: None)
        session._com_initialized = True
        session._com_thread_id = 1

        def release_com():
            calls.append("release_com")
            session._com_initialized = False

        def process_running(pid):
            calls.append("check_pid")
            return False

        monkeypatch.setattr(session, "_uninitialize_com", release_com)
        monkeypatch.setattr(session_module, "is_process_running", process_running)

        session.__exit__(None, None, None)

        assert calls[:2] == ["release_com", "check_pid"]

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
    def test_exit_keeps_dialog_watchdog_active_until_owned_process_exits(
        self, monkeypatch,
    ):
        from xlvbatools.core import session as session_module

        calls = []

        class Watchdog:
            def stop(self):
                calls.append("watchdog_stop")
                return []

        def process_running(pid):
            calls.append("check_pid")
            return False

        monkeypatch.setattr(session_module, "is_process_running", process_running)

        session = session_module.ExcelSession("dummy.xlsm", save_on_exit=False)
        session.excel_pid = 4242
        session.watchdog = Watchdog()
        session.wb = SimpleNamespace(Close=lambda value: calls.append("workbook_close"))
        session.excel = SimpleNamespace(Quit=lambda: calls.append("excel_quit"))

        session.__exit__(None, None, None)

        assert calls == [
            "workbook_close",
            "excel_quit",
            "check_pid",
            "check_pid",
            "check_pid",
            "watchdog_stop",
        ]

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
    def test_sequential_macro_runs_use_event_sequences(self):
        from xlvbatools.core.session import ExcelSession
        from xlvbatools.core.watchdog import DialogEvent

        old = DialogEvent(1, 1, "old", "runtime_error", ["old error"], sequence=1)
        new = DialogEvent(2, 2, "new", "runtime_error", ["new error"], sequence=2)

        class Watchdog:
            events = [old]

        watchdog = Watchdog()

        def run(_macro):
            watchdog.events.append(new)

        session = ExcelSession("dummy.xlsm")
        session.watchdog = watchdog
        session.excel = SimpleNamespace(Run=run)
        result = session.run_macro("TestMacro")
        assert [event["sequence"] for event in result["dialog_events"]] == [2]
        assert result["primary_error"] == "new error"

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
    @pytest.mark.parametrize(
        ("session_kwargs", "expected_events", "expected_security"),
        [
            ({}, False, 3),
            ({"allow_macro_execution": True}, False, 1),
            ({
                "allow_workbook_events": True,
                "allow_macro_execution": True,
            }, True, 1),
        ],
    )
    def test_watchdog_starts_with_spawned_pid(
        self,
        tmp_path,
        monkeypatch,
        session_kwargs,
        expected_events,
        expected_security,
    ):
        import win32com.client
        import win32process
        from xlvbatools.core import session as session_module
        from xlvbatools.core import watchdog as watchdog_module

        workbook = tmp_path / "book.xlsm"
        workbook.touch()
        created = []

        class FakeWatchdog:
            had_errors = False
            had_dialogs = False
            error_summary = ""
            events = []

            def __init__(self, **kwargs):
                created.append(kwargs)

            def start(self):
                pass

            def stop(self):
                return []

        fake_wb = SimpleNamespace(Close=lambda value: None)
        fake_excel = SimpleNamespace(
            Hwnd=99,
            Visible=False,
            DisplayAlerts=True,
            AutomationSecurity=0,
            Workbooks=SimpleNamespace(Open=lambda *args, **kwargs: fake_wb),
            Quit=lambda: None,
        )
        monkeypatch.setattr(win32com.client, "DispatchEx", lambda name: fake_excel)
        monkeypatch.setattr(win32process, "GetWindowThreadProcessId", lambda hwnd: (1, 4321))
        monkeypatch.setattr(watchdog_module, "DialogWatchdog", FakeWatchdog)
        monkeypatch.setattr(session_module, "is_process_running", lambda pid: False)

        with session_module.ExcelSession(
            str(workbook), kill_on_enter=False, init_delay=0, **session_kwargs,
        ) as session:
            assert session.excel_pid == 4321
            assert fake_excel.EnableEvents is expected_events
            assert fake_excel.AutomationSecurity == expected_security

        assert created[0]["target_pid"] == 4321
        assert created[0]["auto_dismiss"] is True
        assert created[0]["hide_vbe_main_window"] is True

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
    def test_com_apartment_is_balanced(self, monkeypatch):
        import pythoncom
        from xlvbatools.core.session import ExcelSession

        calls = []
        monkeypatch.setattr(pythoncom, "CoInitialize", lambda: calls.append("initialize"))
        monkeypatch.setattr(pythoncom, "CoFreeUnusedLibraries", lambda: calls.append("free"))
        monkeypatch.setattr(pythoncom, "CoUninitialize", lambda: calls.append("uninitialize"))
        monkeypatch.setattr(ExcelSession, "_thread_has_com_apartment", staticmethod(lambda: False))

        session = ExcelSession("dummy.xlsm")
        session._initialize_com()
        session._uninitialize_com()

        assert calls == ["initialize", "free", "uninitialize"]
        assert session._com_initialized is False
        assert session._com_thread_id is None

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
    def test_caller_owned_com_apartment_is_not_uninitialized(self, monkeypatch):
        import pythoncom
        from xlvbatools.core.session import ExcelSession

        calls = []
        monkeypatch.setattr(ExcelSession, "_thread_has_com_apartment", staticmethod(lambda: True))
        monkeypatch.setattr(pythoncom, "CoInitialize", lambda: calls.append("initialize"))
        monkeypatch.setattr(pythoncom, "CoUninitialize", lambda: calls.append("uninitialize"))

        session = ExcelSession("dummy.xlsm")
        session._initialize_com()
        session._uninitialize_com()

        assert calls == []


@pytest.mark.excel
class TestSessionCOM:
    """Integration tests requiring Excel COM. Skipped unless Excel is installed."""

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
    def test_open_close_minimal_workbook(self, minimal_workbook):
        from xlvbatools.core.session import ExcelSession
        with ExcelSession(minimal_workbook, save_on_exit=False) as session:
            assert session.excel is not None
            assert session.wb is not None
            assert session.excel_pid is not None
            assert not session.had_errors
        assert session.cleanup_result["exited_gracefully"] is True
        assert session.cleanup_result["force_terminated"] is False

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
    def test_multiline_runtime_error_is_captured(self, runtime_error_workbook):
        from xlvbatools.core.session import ExcelSession

        with ExcelSession(
            runtime_error_workbook,
            save_on_exit=False,
            allow_workbook_events=True,
            allow_macro_execution=True,
        ) as session:
            result = session.run_macro("RaiseMultilineError")

        assert result["success"] is False
        assert result["dialog_events"][0]["type"] == "runtime_error"
        assert "Diagnostic line one." in result["dialog_events"][0]["text"]
        assert "Diagnostic line two." in result["dialog_events"][0]["text"]
        assert session.cleanup_result["still_running"] is False
        assert session.cleanup_result["force_terminated"] is False

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
    def test_valid_compile_targets_requested_project_and_stays_headless(
        self, minimal_workbook, tmp_path,
    ):
        import shutil

        from xlvbatools.core.session import ExcelSession

        unrelated_path = tmp_path / "unrelated.xlsm"
        shutil.copy2(minimal_workbook, unrelated_path)

        other_workbook = None
        other_component = None
        target_project = None
        target_component = None
        target_code = None
        with ExcelSession(
            minimal_workbook,
            save_on_exit=False,
            allow_workbook_events=False,
            allow_macro_execution=True,
        ) as session:
            target_project = session.vb_project
            target_component = target_project.VBComponents.Add(1)
            target_component.Name = "modValidCompileTarget"
            target_code = target_component.CodeModule
            target_code.AddFromString(
                "Option Explicit\r\nPublic Sub ValidCompileTarget()\r\nEnd Sub\r\n"
            )

            other_workbook = session.excel.Workbooks.Open(
                str(unrelated_path),
                UpdateLinks=0,
                ReadOnly=True,
                AddToMru=False,
            )
            other_component = other_workbook.VBProject.VBComponents.Item(1)
            other_component.Activate()

            result = session.compile_test()
            assert _vbe_is_visible(session.excel) is False

            assert result["success"] is True, result
            assert result["compile_verified"] is True
            assert result["compile_command_executed"] is True
            assert result["target_verified"] is True
            assert result["errors"] == []
            assert result["target_project_file"].casefold() == minimal_workbook.casefold()
            assert result["active_project_file"].casefold() == minimal_workbook.casefold()

            other_component = None
            other_workbook.Close(False)
            other_workbook = None
            target_project.VBComponents.Remove(target_component)
            target_code = None
            target_component = None
            target_project = None
            gc.collect()
            gc.collect()

        assert session.dialog_events == []
        assert session.cleanup_result["still_running"] is False
        assert session.cleanup_result["force_terminated"] is False

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
    def test_compile_error_location_is_headless(self, compile_error_workbook):
        from xlvbatools.core.session import ExcelSession

        with ExcelSession(
            compile_error_workbook,
            save_on_exit=False,
            allow_workbook_events=False,
            allow_macro_execution=True,
        ) as session:
            result = session.compile_test()
            assert _vbe_is_visible(session.excel) is False

        assert result["success"] is False
        assert result["compile_verified"] is True
        assert result["compile_command_executed"] is True
        assert result["target_verified"] is True
        assert result["project_mode_after_compile"] == 2, result
        assert result["warnings"] == [], result["warnings"]
        assert result["error_module"] == "modCompileFailure"
        assert result["error_line"] == 3
        assert result["error_column"] == 5
        assert "undeclaredCompileValue" in "\n".join(result["error_context"])
        assert session.cleanup_result["force_terminated"] is False

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
    def test_graceful_close_target_only(self, minimal_workbook, tmp_path):
        import shutil
        from xlvbatools.core.session import ExcelSession
        from xlvbatools.core.process import is_process_running

        # Both workbooks are disposable synthetic fixtures owned by this test.
        target_workbook = tmp_path / "target.xlsm"
        shutil.copy(minimal_workbook, target_workbook)

        unrelated_workbook = tmp_path / "unrelated.xlsm"
        shutil.copy(minimal_workbook, unrelated_workbook)

        # Keep a second independently owned Excel session open while the target
        # session exits. Managing both through ExcelSession prevents this test
        # itself from retaining raw COM proxies after Application.Quit.
        with ExcelSession(
            str(unrelated_workbook), kill_on_enter=False, save_on_exit=False,
        ) as unrelated_session:
            unrelated_pid = unrelated_session.excel_pid
            # Open and close an isolated target session while an unrelated
            # workbook remains live in a separate Excel process.
            with ExcelSession(str(target_workbook), kill_on_enter=False, save_on_exit=False) as session:
                assert session.excel is not None
                assert session.wb is not None

            # Verify that unrelated_workbook is still open and running.
            assert unrelated_session.wb.Name == "unrelated.xlsm"
            assert is_process_running(unrelated_pid)

        assert unrelated_session.cleanup_result["still_running"] is False

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
    def test_child_com_proxies_finalize_before_excel(self, minimal_workbook):
        """Successful pytest status must not hide native pywin32 diagnostics."""
        code = textwrap.dedent(
            """
            import gc
            import json
            import sys
            from xlvbatools.core.session import ExcelSession

            with ExcelSession(sys.argv[1], save_on_exit=False, kill_on_enter=False) as session:
                sheet = session.wb.Worksheets(1)
                cell = sheet.Range("A1")
                _ = cell.Value
                cell = None
                sheet = None
                gc.collect()

            print(json.dumps(session.cleanup_result))
            """
        )
        completed = subprocess.run(
            [sys.executable, "-X", "faulthandler", "-c", code, minimal_workbook],
            capture_output=True,
            text=True,
            timeout=60,
        )
        combined = completed.stdout + completed.stderr
        assert completed.returncode == 0, combined
        assert '"still_running": false' in completed.stdout.lower()
        for signature in ("Windows fatal exception", "0x800706ba", "0x80010108"):
            assert signature not in combined, combined
