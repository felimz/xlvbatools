"""
Tests for xlvbatools.core.process -- Excel process management.
"""

import pytest
import sys


@pytest.mark.unit
class TestProcessUtilities:
    """Unit tests for process management functions (no COM needed)."""

    def test_is_excel_running_returns_bool(self):
        from xlvbatools.core.process import is_excel_running
        result = is_excel_running()
        assert isinstance(result, bool)

    def test_is_process_running_nonexistent_pid(self):
        from xlvbatools.core.process import is_process_running
        # PID 99999999 almost certainly doesn't exist
        assert is_process_running(99999999) is False

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
    def test_kill_excel_returns_bool(self):
        from xlvbatools.core.process import kill_excel
        # Just verify it returns a bool (don't actually kill if Excel is running)
        result = kill_excel(timeout=0.1)
        assert isinstance(result, bool)
