"""Process-level regression for sequential Excel/VBE session teardown."""

import os
import subprocess
import sys
import tempfile

import pytest


@pytest.mark.com
@pytest.mark.integration
@pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
def test_session_com_cases_finish_in_one_interpreter():
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    # Do not use PIPE here. If the child interpreter crashes while an Excel
    # process is still alive, Excel can retain an inherited pipe handle and
    # make subprocess.run()/communicate wait forever for EOF. A seekable file
    # lets the parent regain control and report the actual child exit status.
    with tempfile.TemporaryFile(mode="w+", encoding="utf-8", errors="replace") as output:
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
            stdout=output,
            stderr=subprocess.STDOUT,
            timeout=180,
        )
        output.seek(0)
        combined = output.read()
    assert completed.returncode == 0, combined
    assert "5 passed" in combined, combined
    assert "Windows fatal exception" not in combined, combined
    assert "0x800706ba" not in combined, combined
    assert "0x80010108" not in combined, combined
