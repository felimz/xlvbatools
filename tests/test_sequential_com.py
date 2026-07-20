"""Process-level regression for sequential Excel/VBE session teardown."""

import os
import subprocess
import sys
import tempfile
import textwrap

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


@pytest.mark.com
@pytest.mark.integration
@pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
def test_repeated_formula_formatting_and_macro_sessions(runtime_error_workbook):
    """Repeated rich COM reads and macros survive GC in one interpreter."""
    code = textwrap.dedent(
        """
        import gc
        import json
        import sys

        from xlvbatools.core.process import is_process_running
        from xlvbatools.core.session import ExcelSession

        results = []
        for iteration in range(5):
            sheet = None
            cell = None
            interior = None
            vb_project = None
            components = None
            component = None
            code_module = None
            with ExcelSession(
                sys.argv[1], save_on_exit=False, kill_on_enter=False,
                allow_workbook_events=True, allow_macro_execution=True,
            ) as session:
                sheet = session.wb.Worksheets(1)
                cell = sheet.Range("B1")
                formula = cell.Formula
                number_format = cell.NumberFormat
                interior = cell.Interior
                color = interior.Color

                vb_project = session.vb_project
                components = vb_project.VBComponents
                component = components.Item("modReliabilityTest")
                code_module = component.CodeModule
                line_count = code_module.CountOfLines

                macro_result = session.run_macro("CompleteNormally")
                assert formula == "=21*2"
                assert number_format == "0.00"
                assert color == 65535
                assert line_count > 0
                assert macro_result["success"] is True, macro_result

                code_module = None
                component = None
                components = None
                vb_project = None
                interior = None
                cell = None
                sheet = None
                gc.collect()
                gc.collect()

            cleanup = dict(session.cleanup_result)
            assert cleanup["still_running"] is False, cleanup
            assert not is_process_running(cleanup["pid"]), cleanup
            results.append({
                "iteration": iteration + 1,
                "pid": cleanup["pid"],
                "cleanup": cleanup,
                "dialog_events": [
                    event.to_dict() for event in session.dialog_events
                ],
            })
            gc.collect()
            gc.collect()

        print("SEQUENTIAL_RESULT=" + json.dumps(results))
        """
    )
    with tempfile.TemporaryFile(mode="w+", encoding="utf-8", errors="replace") as output:
        completed = subprocess.run(
            [
                sys.executable,
                "-X",
                "faulthandler",
                "-c",
                code,
                runtime_error_workbook,
            ],
            stdout=output,
            stderr=subprocess.STDOUT,
            timeout=180,
        )
        output.seek(0)
        combined = output.read()

    assert completed.returncode == 0, combined
    assert "SEQUENTIAL_RESULT=" in combined, combined
    assert combined.count('"iteration":') == 5, combined
    for signature in (
        "Windows fatal exception", "0x800706ba", "0x80010108",
        "RPC server is unavailable", "CoInitialize has not been called",
    ):
        assert signature not in combined, combined


@pytest.mark.com
@pytest.mark.integration
@pytest.mark.e2e
@pytest.mark.skipif(sys.platform != "win32", reason="Windows only")
def test_fifty_project_runs_retain_progress_and_cleanup(runtime_error_workbook):
    """Stress the public executor repeatedly in one long-lived interpreter."""
    code = textwrap.dedent(
        """
        import json
        import sys

        from xlvbatools import Project
        from xlvbatools.core.process import is_process_running

        project = Project.open(sys.argv[1])
        for iteration in range(50):
            result = project.run("CompleteNormally", timeout=60, save=False)
            result.require_success()
            cleanup = result.require_clean_shutdown()
            assert not is_process_running(cleanup.pid), result.to_dict()
            payload = result.to_dict()
            print("PROJECT_RUN_PROGRESS=" + json.dumps({
                "iteration": iteration + 1,
                "success": result.success,
                "phase": result.phase,
                "attempt_count": result.attempt_count,
                "request_id": result.request_id,
                "worker_pid": result.diagnostics.worker_pid,
                "excel_pid": result.diagnostics.excel_pid,
                "elapsed_seconds": result.elapsed_seconds,
                "attempts": payload["diagnostics"]["attempts"],
                "cleanup": payload["diagnostics"]["cleanup"],
            }), flush=True)
        print("PROJECT_RUN_STRESS_COMPLETE=50", flush=True)
        """
    )
    with tempfile.TemporaryFile(mode="w+", encoding="utf-8", errors="replace") as output:
        completed = subprocess.run(
            [
                sys.executable,
                "-X",
                "faulthandler",
                "-c",
                code,
                runtime_error_workbook,
            ],
            stdout=output,
            stderr=subprocess.STDOUT,
            timeout=600,
        )
        output.seek(0)
        combined = output.read()

    assert completed.returncode == 0, combined
    assert "PROJECT_RUN_STRESS_COMPLETE=50" in combined, combined
    assert combined.count("PROJECT_RUN_PROGRESS=") == 50, combined
    for signature in (
        "Windows fatal exception", "0x800706ba", "0x80010108",
        "RPC server is unavailable", "CoInitialize has not been called",
    ):
        assert signature not in combined, combined
