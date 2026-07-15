"""
Tests for xlvbatools.core.process -- Excel process management.
"""

import pytest


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
