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
