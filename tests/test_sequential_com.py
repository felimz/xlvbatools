"""Process-level regression for sequential Excel/VBE session teardown."""

import os
import subprocess
import sys

import pytest


@pytest.mark.com
@pytest.mark.integration
@pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
def test_session_com_cases_finish_in_one_interpreter():
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    completed = subprocess.run(
        [
            sys.executable,
            "-X",
            "faulthandler",
            "-m",
            "pytest",
            "tests/test_session.py",
            "-m",
            "com",
            "-q",
        ],
        cwd=project_root,
        capture_output=True,
        text=True,
        timeout=180,
    )
    combined = completed.stdout + completed.stderr
    assert completed.returncode == 0, combined
    assert "5 passed" in completed.stdout, combined
    assert "Windows fatal exception" not in combined, combined
    assert "0x800706ba" not in combined, combined
    assert "0x80010108" not in combined, combined
