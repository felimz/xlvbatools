"""
Tests for xlvbatools.core.process -- Excel process management.
"""

from pathlib import Path
import subprocess
import sys

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

    @pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
    def test_exact_pid_probe_and_termination_use_win32_handles(self):
        from xlvbatools.core.process import is_process_running, kill_process_by_pid

        child = subprocess.Popen(
            [sys.executable, "-c", "import time; time.sleep(60)"],
        )
        try:
            assert is_process_running(child.pid) is True
            assert kill_process_by_pid(child.pid) is True
            assert child.wait(timeout=10) != 0
            assert is_process_running(child.pid) is False
        finally:
            if child.poll() is None:
                child.kill()
                child.wait(timeout=10)

    def test_exact_pid_helpers_do_not_shell_out_to_task_utilities(self):
        from xlvbatools.core import process as process_module

        source = Path(process_module.__file__).read_text(encoding="utf-8")
        assert '["tasklist", "/fi", f"PID eq {pid}"' not in source
        assert '["taskkill", "/f", "/pid"' not in source
