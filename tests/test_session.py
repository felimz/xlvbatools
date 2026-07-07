"""
Tests for xlvbatools.core.session -- ExcelSession context manager.
"""

import os
import sys
import pytest


@pytest.mark.unit
class TestSessionImport:
    """Test that ExcelSession can be imported (even on non-Windows, via lazy import)."""

    def test_import_from_package(self):
        from xlvbatools import ExcelSession
        assert ExcelSession is not None

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


@pytest.mark.com
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
