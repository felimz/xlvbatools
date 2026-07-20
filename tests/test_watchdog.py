"""
Tests for xlvbatools.core.watchdog -- Dialog event dataclass and classification.
"""

import pytest
import sys


@pytest.mark.unit
class TestDialogEvent:
    """Unit tests for DialogEvent (no COM or Win32 needed)."""

    @pytest.mark.skipif(sys.platform != "win32", reason="watchdog requires Windows")
    def test_dialog_event_creation(self):
        from xlvbatools.core.watchdog import DialogEvent
        event = DialogEvent(
            timestamp=1234567890.0,
            hwnd=0,
            title="Microsoft Visual Basic",
            dialog_type="compile_error",
            texts=["Compile error:", "Expected: end of statement"],
        )
        assert event.dialog_type == "compile_error"
        assert "Expected: end of statement" in event.text
        assert not event.dismissed

    @pytest.mark.skipif(sys.platform != "win32", reason="watchdog requires Windows")
    def test_dialog_event_to_dict(self):
        from xlvbatools.core.watchdog import DialogEvent
        event = DialogEvent(
            timestamp=100.0,
            hwnd=12345,
            title="Test",
            dialog_type="msgbox",
            texts=["Hello"],
            button_clicked="OK",
            dismissed=True,
        )
        d = event.to_dict()
        assert d["type"] == "msgbox"
        assert d["dismissed"] is True
        assert d["button_clicked"] == "OK"
        assert d["hwnd"] == 12345

    @pytest.mark.skipif(sys.platform != "win32", reason="watchdog requires Windows")
    def test_dialog_event_str(self):
        from xlvbatools.core.watchdog import DialogEvent
        event = DialogEvent(
            timestamp=100.0,
            hwnd=0,
            title="Error",
            dialog_type="runtime_error",
            texts=["Run-time error '1004'"],
            dismissed=True,
        )
        s = str(event)
        assert "DISMISSED" in s
        assert "runtime_error" in s


@pytest.mark.unit
class TestDialogClassifier:
    """Unit tests for dialog type classification."""

    @pytest.mark.skipif(sys.platform != "win32", reason="watchdog requires Windows")
    def test_classify_compile_error(self):
        from xlvbatools.core.watchdog import _classify_dialog
        assert _classify_dialog("Microsoft Visual Basic", ["Compile error: Expected"]) == "compile_error"

    @pytest.mark.skipif(sys.platform != "win32", reason="watchdog requires Windows")
    def test_classify_runtime_error(self):
        from xlvbatools.core.watchdog import _classify_dialog
        assert _classify_dialog("Microsoft Visual Basic", ["Run-time error '1004'"]) == "runtime_error"

    @pytest.mark.skipif(sys.platform != "win32", reason="watchdog requires Windows")
    def test_classify_excel_alert(self):
        from xlvbatools.core.watchdog import _classify_dialog
        assert _classify_dialog("Microsoft Excel", ["Something happened"]) == "excel_alert"

    @pytest.mark.skipif(sys.platform != "win32", reason="watchdog requires Windows")
    def test_classify_save_dialog(self):
        from xlvbatools.core.watchdog import _classify_dialog
        assert _classify_dialog("Microsoft Excel", ["Do you want to save changes?"]) == "save_dialog"

    @pytest.mark.skipif(sys.platform != "win32", reason="watchdog requires Windows")
    def test_classify_unknown(self):
        from xlvbatools.core.watchdog import _classify_dialog
        assert _classify_dialog("Some Other App", ["random text"]) == "unknown"


@pytest.mark.unit
class TestWatchdogLifecycle:
    """Test watchdog start/stop lifecycle."""

    @pytest.mark.skipif(sys.platform != "win32", reason="watchdog requires Windows")
    def test_start_stop_no_events(self, monkeypatch):
        from xlvbatools.core.watchdog import DialogWatchdog
        monkeypatch.setattr(DialogWatchdog, "_scan_for_dialogs", lambda self: None)
        wd = DialogWatchdog(poll_interval=0.1, timeout=5.0, auto_dismiss=False)
        wd.start()
        assert not wd.had_dialogs
        assert not wd.had_errors
        events = wd.stop()
        assert events == []

    @pytest.mark.skipif(sys.platform != "win32", reason="watchdog requires Windows")
    def test_error_summary_empty(self, monkeypatch):
        from xlvbatools.core.watchdog import DialogWatchdog
        monkeypatch.setattr(DialogWatchdog, "_scan_for_dialogs", lambda self: None)
        wd = DialogWatchdog(poll_interval=0.1, timeout=5.0, auto_dismiss=False)
        wd.start()
        wd.stop()
        assert wd.error_summary == ""

    def test_auto_dismiss_requires_pid(self):
        from xlvbatools.core.watchdog import DialogWatchdog
        with pytest.raises(ValueError, match="requires target_pid"):
            DialogWatchdog()

    def test_capture_only_allows_missing_pid(self):
        from xlvbatools.core.watchdog import DialogWatchdog
        assert DialogWatchdog(auto_dismiss=False).target_pid is None

    def test_vbe_hiding_requires_pid(self):
        from xlvbatools.core.watchdog import DialogWatchdog
        with pytest.raises(ValueError, match="VBE-hiding watchdog requires target_pid"):
            DialogWatchdog(auto_dismiss=False, hide_vbe_main_window=True)


@pytest.mark.unit
class TestDialogCapture:
    def test_button_mnemonic_is_normalized(self):
        from xlvbatools.core.watchdog import _normalize_button_text
        assert _normalize_button_text("&End") == "end"

    def test_get_control_text_falls_back_to_window_text(self, monkeypatch):
        from xlvbatools.core import watchdog

        class User32:
            def SendMessageTimeoutW(self, *args):
                return 0

        monkeypatch.setattr(watchdog, "user32", User32())
        monkeypatch.setattr(watchdog, "_get_window_text", lambda hwnd: "line one\nline two")
        assert watchdog._get_control_text(123) == "line one\r\nline two"

    def test_delayed_capture_deduplicates_and_classifies(self, monkeypatch):
        from xlvbatools.core import watchdog

        attempts = {"count": 0}

        class User32:
            def EnumChildWindows(self, hwnd, callback, lparam):
                attempts["count"] += 1
                callback(10, 0)
                callback(11, 0)

        monkeypatch.setattr(watchdog, "user32", User32())
        monkeypatch.setattr(watchdog, "EnumChildProc", lambda callback: callback)
        monkeypatch.setattr(watchdog, "_get_window_text", lambda hwnd: "Microsoft Visual Basic")
        monkeypatch.setattr(watchdog, "_get_window_class", lambda hwnd: "Static")
        monkeypatch.setattr(
            watchdog, "_get_control_text",
            lambda hwnd: "" if attempts["count"] == 1 else "Run-time error '5':\r\n\r\nline one\r\nline two",
        )
        wd = watchdog.DialogWatchdog(auto_dismiss=False, capture_retry_delay=0)
        wd._handle_dialog(1)
        event = wd.events[0]
        assert attempts["count"] == 2
        assert event.dialog_type == "runtime_error"
        assert event.texts == ["Run-time error '5':\r\n\r\nline one\r\nline two"]
        assert "line one\r\nline two" in event.text
        assert event.sequence == 1

    def test_dialog_from_other_pid_is_ignored(self, monkeypatch):
        from xlvbatools.core import watchdog

        class User32:
            def EnumWindows(self, callback, lparam):
                callback(123, 0)

        monkeypatch.setattr(watchdog, "user32", User32())
        monkeypatch.setattr(watchdog, "EnumWindowsProc", lambda callback: callback)
        monkeypatch.setattr(watchdog, "_is_window_visible", lambda hwnd: True)
        monkeypatch.setattr(watchdog, "_get_window_class", lambda hwnd: watchdog.DIALOG_CLASS)
        monkeypatch.setattr(watchdog, "_get_window_pid", lambda hwnd: (1, 999))
        monkeypatch.setattr(watchdog, "_get_window_text", lambda hwnd: "Microsoft Excel")
        handled = []
        wd = watchdog.DialogWatchdog(target_pid=42)
        monkeypatch.setattr(wd, "_handle_dialog", handled.append)
        wd._scan_for_dialogs()
        assert handled == []

    def test_hidden_dialog_handle_is_recaptured_if_vbe_reuses_it(
        self, monkeypatch,
    ):
        from xlvbatools.core import watchdog

        visible = iter((True, False, True))

        class User32:
            def EnumWindows(self, callback, lparam):
                callback(123, 0)

            def EnumChildWindows(self, hwnd, callback, lparam):
                return None

        monkeypatch.setattr(watchdog, "user32", User32())
        monkeypatch.setattr(watchdog, "EnumWindowsProc", lambda callback: callback)
        monkeypatch.setattr(watchdog, "EnumChildProc", lambda callback: callback)
        monkeypatch.setattr(watchdog, "_is_window_visible", lambda hwnd: next(visible))
        monkeypatch.setattr(watchdog, "_get_window_class", lambda hwnd: watchdog.DIALOG_CLASS)
        monkeypatch.setattr(
            watchdog,
            "_get_window_text",
            lambda hwnd: "Microsoft Visual Basic for Applications",
        )

        wd = watchdog.DialogWatchdog(auto_dismiss=False, capture_attempts=1)
        wd._scan_for_dialogs()
        wd._scan_for_dialogs()
        wd._scan_for_dialogs()

        assert [event.hwnd for event in wd.events] == [123, 123]
        assert [event.sequence for event in wd.events] == [1, 2]

    def test_dismissal_requires_confirmation_and_queues_a_busy_button_click(
        self, monkeypatch,
    ):
        from xlvbatools.core import watchdog

        actions = []
        hidden = iter((False, True))
        monkeypatch.setattr(
            watchdog,
            "_click_button",
            lambda hwnd: actions.append(("send", hwnd)) or False,
        )
        monkeypatch.setattr(
            watchdog,
            "_post_button_click",
            lambda hwnd: actions.append(("post", hwnd)) or True,
        )
        monkeypatch.setattr(
            watchdog,
            "_wait_for_window_to_hide",
            lambda hwnd: next(hidden),
        )

        event = watchdog.DialogEvent(
            timestamp=1.0,
            hwnd=123,
            title="Microsoft Visual Basic for Applications",
            dialog_type="compile_error",
        )
        wd = watchdog.DialogWatchdog(target_pid=42)

        dismissed = wd._dismiss_dialog(123, [(456, "OK")], event)

        assert dismissed is True
        assert event.button_clicked == "OK"
        assert actions == [("send", 456), ("post", 456)]

    def test_failed_dismissal_is_retried_for_same_visible_handle(
        self, monkeypatch,
    ):
        from xlvbatools.core import watchdog

        class User32:
            def EnumWindows(self, callback, lparam):
                callback(123, 0)

            def EnumChildWindows(self, hwnd, callback, lparam):
                callback(456, 0)

        monkeypatch.setattr(watchdog, "user32", User32())
        monkeypatch.setattr(watchdog, "EnumWindowsProc", lambda callback: callback)
        monkeypatch.setattr(watchdog, "EnumChildProc", lambda callback: callback)
        monkeypatch.setattr(watchdog, "_is_window_visible", lambda hwnd: True)
        monkeypatch.setattr(
            watchdog,
            "_get_window_class",
            lambda hwnd: watchdog.DIALOG_CLASS if hwnd == 123 else "Button",
        )
        monkeypatch.setattr(
            watchdog,
            "_get_window_text",
            lambda hwnd: (
                "Microsoft Visual Basic for Applications" if hwnd == 123 else "OK"
            ),
        )
        monkeypatch.setattr(watchdog, "_get_control_text", lambda hwnd: "OK")
        monkeypatch.setattr(
            watchdog.DialogWatchdog,
            "_dismiss_dialog",
            lambda self, hwnd, buttons, event: False,
        )

        wd = watchdog.DialogWatchdog(target_pid=42, capture_attempts=1)
        monkeypatch.setattr(watchdog, "_get_window_pid", lambda hwnd: (1, 42))
        wd._scan_for_dialogs()
        wd._scan_for_dialogs()

        assert [event.sequence for event in wd.events] == [1, 2]
        assert all(event.dismissed is False for event in wd.events)

    def test_owned_vbe_main_window_is_hidden_without_becoming_a_dialog(
        self, monkeypatch,
    ):
        from xlvbatools.core import watchdog

        hidden = []

        class User32:
            def EnumWindows(self, callback, lparam):
                callback(123, 0)

        monkeypatch.setattr(watchdog, "user32", User32())
        monkeypatch.setattr(watchdog, "EnumWindowsProc", lambda callback: callback)
        monkeypatch.setattr(watchdog, "_is_window_visible", lambda hwnd: True)
        monkeypatch.setattr(watchdog, "_get_window_class", lambda hwnd: "wndclass_desked_gsk")
        monkeypatch.setattr(watchdog, "_get_window_pid", lambda hwnd: (1, 42))
        monkeypatch.setattr(
            watchdog,
            "_get_window_text",
            lambda hwnd: "Microsoft Visual Basic for Applications - Book1",
        )
        monkeypatch.setattr(watchdog, "_hide_window", hidden.append)
        wd = watchdog.DialogWatchdog(
            target_pid=42,
            hide_vbe_main_window=True,
        )
        wd._scan_for_dialogs()

        assert hidden == [123]
        assert wd.events == []


@pytest.mark.unit
def test_compile_test_fails_closed_when_compile_control_is_unavailable(monkeypatch):
    from types import SimpleNamespace
    from xlvbatools.core import watchdog

    project = SimpleNamespace(Name="Model", FileName="C:/model.xlsm")
    find_calls = []

    def find_control(**kwargs):
        find_calls.append(kwargs)
        return None

    vbe = SimpleNamespace(
        ActiveVBProject=project,
        CommandBars=SimpleNamespace(FindControl=find_control),
    )
    excel = SimpleNamespace(VBE=vbe)
    workbook = SimpleNamespace(
        FullName="C:/model.xlsm",
        VBProject=project,
        Activate=lambda: None,
    )
    capture = SimpleNamespace(events=[])
    monkeypatch.setattr(watchdog, "_hide_vbe_ui", lambda excel: None)

    result = watchdog.compile_test_with_watchdog(excel, workbook, capture)

    assert result["success"] is False
    assert result["compile_verified"] is False
    assert result["target_verified"] is True
    assert find_calls == [{"Type": 1, "Id": 578}]
    assert result["errors"] == [{
        "type": "compile_unverified",
        "message": (
            "VBE compile button (Type=1, ID=578) was unavailable; "
            "compilation could not be verified."
        ),
    }]


@pytest.mark.unit
def test_compile_test_executes_button_and_verifies_exact_project(monkeypatch):
    from types import SimpleNamespace
    from xlvbatools.core import watchdog

    class CompileButton:
        Enabled = True

        def __init__(self):
            self.executions = 0

        def Execute(self):
            self.executions += 1
            self.Enabled = False

    project = SimpleNamespace(Name="Model", FileName="C:/model.xlsm")
    button = CompileButton()
    find_calls = []

    def find_control(**kwargs):
        find_calls.append(kwargs)
        return button

    vbe = SimpleNamespace(
        ActiveVBProject=project,
        CommandBars=SimpleNamespace(FindControl=find_control),
    )
    excel = SimpleNamespace(VBE=vbe)
    workbook = SimpleNamespace(VBProject=project, Activate=lambda: None)
    capture = SimpleNamespace(events=[])
    monkeypatch.setattr(watchdog, "_hide_vbe_ui", lambda excel: None)

    result = watchdog.compile_test_with_watchdog(excel, workbook, capture)

    assert result["success"] is True
    assert result["compile_verified"] is True
    assert result["compile_command_executed"] is True
    assert result["target_verified"] is True
    assert result["target_project_file"] == "C:/model.xlsm"
    assert result["active_project_file"] == "C:/model.xlsm"
    assert find_calls == [{"Type": 1, "Id": 578}]
    assert button.executions == 1


@pytest.mark.unit
def test_compile_test_refuses_to_compile_another_active_project(monkeypatch):
    from types import SimpleNamespace
    from xlvbatools.core import watchdog

    class Component:
        def Activate(self):
            return None

    class Components:
        def Item(self, index):
            assert index == 1
            return Component()

    target = SimpleNamespace(
        Name="Target",
        FileName="C:/target.xlsm",
        VBComponents=Components(),
    )
    other = SimpleNamespace(Name="Other", FileName="C:/other.xlam")
    find_calls = []
    class ReadOnlyVBE:
        CommandBars = SimpleNamespace(
            FindControl=lambda **kwargs: find_calls.append(kwargs)
        )

        @property
        def ActiveVBProject(self):
            return other

        @ActiveVBProject.setter
        def ActiveVBProject(self, value):
            raise AttributeError("read-only test project")

    vbe = ReadOnlyVBE()
    excel = SimpleNamespace(VBE=vbe)
    workbook = SimpleNamespace(VBProject=target, Activate=lambda: None)
    capture = SimpleNamespace(events=[])
    monkeypatch.setattr(watchdog, "_hide_vbe_ui", lambda excel: None)
    clock = iter((0.0, 2.0))
    monkeypatch.setattr(watchdog.time, "time", lambda: next(clock))

    result = watchdog.compile_test_with_watchdog(excel, workbook, capture)

    assert result["success"] is False
    assert result["compile_verified"] is False
    assert result["target_verified"] is False
    assert result["active_project_file"] == "C:/other.xlam"
    assert result["errors"][0]["type"] == "compile_unverified"
    assert "refusing to compile another project" in result["errors"][0]["message"]
    assert find_calls == []
