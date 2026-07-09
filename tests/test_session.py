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

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
    def test_graceful_close_target_only(self, tmp_path):
        import os
        import shutil
        import win32com.client
        from xlvbatools.core.session import ExcelSession

        # Load real sample workbook from repository
        workbook_src = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            "sample_workbooks",
            "Project_Code_Filter.xlsm"
        )
        assert os.path.exists(workbook_src), f"Source workbook {workbook_src} not found"

        target_workbook = tmp_path / "Project_Code_Filter.xlsm"
        shutil.copy(workbook_src, target_workbook)

        unrelated_workbook = tmp_path / "unrelated.xlsm"
        shutil.copy(workbook_src, unrelated_workbook)

        # 1. Open unrelated_workbook manually
        excel1 = win32com.client.Dispatch("Excel.Application")
        excel1.Visible = False
        wb1 = excel1.Workbooks.Open(str(unrelated_workbook))

        # 2. Open target_workbook manually
        excel2 = win32com.client.Dispatch("Excel.Application")
        excel2.Visible = False
        wb2 = excel2.Workbooks.Open(str(target_workbook))

        try:
            # 3. Enter ExcelSession on target_workbook.
            # This should close target_workbook (excel2), but keep unrelated_workbook (excel1) open.
            with ExcelSession(str(target_workbook), kill_on_enter=True, save_on_exit=False) as session:
                assert session.excel is not None
                assert session.wb is not None

            # 4. Verify that unrelated_workbook is still open and running
            assert wb1.Name == "unrelated.xlsm"
        finally:
            # Cleanup manually opened instances
            try:
                wb1.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                excel1.Quit()
            except Exception:
                pass
            try:
                wb2.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                excel2.Quit()
            except Exception:
                pass
