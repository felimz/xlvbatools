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
    def test_start_stop_no_events(self):
        from xlvbatools.core.watchdog import DialogWatchdog
        wd = DialogWatchdog(poll_interval=0.1, timeout=5.0)
        wd.start()
        assert not wd.had_dialogs
        assert not wd.had_errors
        events = wd.stop()
        assert events == []

    @pytest.mark.skipif(sys.platform != "win32", reason="watchdog requires Windows")
    def test_error_summary_empty(self):
        from xlvbatools.core.watchdog import DialogWatchdog
        wd = DialogWatchdog(poll_interval=0.1, timeout=5.0)
        wd.start()
        wd.stop()
        assert wd.error_summary == ""
