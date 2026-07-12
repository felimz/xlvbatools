"""Execute complete Excel sessions in killable worker processes."""

import logging
import multiprocessing
import os
import time
import uuid
from typing import Callable, Optional

from xlvbatools.core.process import is_process_running, kill_process_by_pid
from xlvbatools.core.session import ExcelSession

logger = logging.getLogger(__name__)


def _run_macro_once(
    workbook_path: str,
    macro_name: str,
    named_ranges: Optional[dict],
    timeout: float,
    visible: bool,
    save_on_exit: bool,
    strict_named_ranges: bool,
    run_id: str,
    on_excel_started: Optional[Callable[[int], None]] = None,
    on_phase: Optional[Callable[[str], None]] = None,
) -> dict:
    """Run one session in the current process; used only by the worker."""
    session = None
    started = time.monotonic()
    phase = "session_start"
    if on_phase:
        on_phase(phase)
    try:
        with ExcelSession(
            workbook_path,
            visible=visible,
            save_on_exit=save_on_exit,
            on_excel_started=on_excel_started,
        ) as session:
            phase = "named_range_setup"
            if on_phase:
                on_phase(phase)
            if named_ranges:
                for name, value in named_ranges.items():
                    session.set_named_range(name, value, strict=strict_named_ranges)

            phase = "macro_execution"
            if on_phase:
                on_phase(phase)
            result = session.run_macro(macro_name, timeout=timeout)
            result["run_id"] = run_id
            result["excel_pid"] = session.excel_pid

        result["cleanup"] = session.cleanup_result
        if session.cleanup_result.get("workbook_save_error"):
            result.update(
                success=False,
                phase="workbook_save",
                primary_error=session.cleanup_result["workbook_save_error"],
            )
        if session.cleanup_result.get("still_running"):
            result.update(
                success=False,
                phase="cleanup",
                primary_error=f"Owned Excel PID {session.excel_pid} remained running after cleanup",
            )
        return result
    except Exception as error:
        cleanup = session.cleanup_result if session is not None else {}
        if phase == "session_start" and session is not None:
            phase = session.phase
        return {
            "success": False,
            "run_id": run_id,
            "macro": macro_name,
            "phase": phase,
            "elapsed_seconds": time.monotonic() - started,
            "primary_error": str(error),
            "error": str(error),
            "dialog_events": [event.to_dict() for event in session.dialog_events] if session else [],
            "cleanup": cleanup,
            "excel_pid": getattr(session, "excel_pid", None),
        }


def _macro_worker(connection, arguments: dict) -> None:
    """Spawn-safe worker entry point. All COM objects remain in this process."""
    pythoncom = None
    try:
        try:
            import pythoncom as _pythoncom
            pythoncom = _pythoncom
            pythoncom.CoInitialize()
        except ImportError:
            pass

        def report_pid(pid: int) -> None:
            connection.send({"kind": "excel_started", "excel_pid": pid, "phase": "workbook_open"})

        def report_phase(phase: str) -> None:
            connection.send({"kind": "phase", "phase": phase})

        result = _run_macro_once(
            **arguments, on_excel_started=report_pid, on_phase=report_phase
        )
        connection.send({"kind": "result", "result": result})
    except BaseException as error:
        try:
            connection.send({
                "kind": "worker_error",
                "error": f"{type(error).__name__}: {error}",
            })
        except Exception:
            pass
    finally:
        connection.close()
        if pythoncom is not None:
            pythoncom.CoUninitialize()


def _terminate_timed_out_run(process, excel_pid: Optional[int], grace_period: float = 2.0) -> dict:
    """Terminate the reported Excel PID first, then the blocked worker."""
    force_terminated = False
    if excel_pid is not None and is_process_running(excel_pid):
        force_terminated = kill_process_by_pid(excel_pid)

    process.join(grace_period)
    worker_terminated = False
    if process.is_alive():
        process.terminate()
        process.join(grace_period)
        worker_terminated = True

    return {
        "pid": excel_pid,
        "quit_requested": False,
        "exited_gracefully": False,
        "force_terminated": force_terminated,
        "worker_terminated": worker_terminated,
        "still_running": bool(excel_pid and is_process_running(excel_pid)),
    }


def run_macro(
    workbook_path: str,
    macro_name: str,
    named_ranges: Optional[dict] = None,
    timeout: float = 120.0,
    visible: bool = False,
    save_on_exit: bool = True,
    strict_named_ranges: bool = True,
) -> dict:
    """Run a complete isolated Excel session with a real parent-enforced timeout."""
    if timeout <= 0:
        raise ValueError("timeout must be greater than zero")

    run_id = str(uuid.uuid4())
    arguments = {
        "workbook_path": os.path.abspath(workbook_path),
        "macro_name": macro_name,
        "named_ranges": named_ranges,
        "timeout": timeout,
        "visible": visible,
        "save_on_exit": save_on_exit,
        "strict_named_ranges": strict_named_ranges,
        "run_id": run_id,
    }

    context = multiprocessing.get_context("spawn")
    parent_connection, child_connection = context.Pipe(duplex=False)
    process = context.Process(target=_macro_worker, args=(child_connection, arguments))
    process.start()
    child_connection.close()

    started = time.monotonic()
    deadline = started + timeout
    excel_pid = None
    phase = "session_start"
    try:
        while True:
            remaining = deadline - time.monotonic()
            if remaining <= 0:
                cleanup = _terminate_timed_out_run(process, excel_pid)
                return {
                    "success": False,
                    "run_id": run_id,
                    "macro": macro_name,
                    "phase": phase,
                    "elapsed_seconds": time.monotonic() - started,
                    "timed_out": True,
                    "timeout_seconds": timeout,
                    "excel_pid": excel_pid,
                    "primary_error": f"Execution timed out after {timeout:.3f} seconds",
                    "dialog_events": [],
                    "cleanup": cleanup,
                }

            if parent_connection.poll(min(remaining, 0.1)):
                try:
                    message = parent_connection.recv()
                except EOFError:
                    message = None
                if message is None:
                    break
                if message["kind"] == "excel_started":
                    excel_pid = message["excel_pid"]
                    phase = message["phase"]
                elif message["kind"] == "phase":
                    phase = message["phase"]
                elif message["kind"] == "result":
                    process.join(5.0)
                    return message["result"]
                elif message["kind"] == "worker_error":
                    process.join(2.0)
                    cleanup = _terminate_timed_out_run(process, excel_pid, grace_period=1.0)
                    return {
                        "success": False,
                        "run_id": run_id,
                        "macro": macro_name,
                        "phase": phase,
                        "elapsed_seconds": time.monotonic() - started,
                        "primary_error": message["error"],
                        "error": message["error"],
                        "dialog_events": [],
                        "cleanup": cleanup,
                    }

            if not process.is_alive() and not parent_connection.poll():
                break

        process.join(1.0)
        cleanup = _terminate_timed_out_run(process, excel_pid, grace_period=1.0)
        return {
            "success": False,
            "run_id": run_id,
            "macro": macro_name,
            "phase": phase,
            "elapsed_seconds": time.monotonic() - started,
            "primary_error": f"Macro worker exited without a result (exit code {process.exitcode})",
            "error": f"Macro worker exited without a result (exit code {process.exitcode})",
            "dialog_events": [],
            "cleanup": cleanup,
        }
    finally:
        parent_connection.close()
        if process.is_alive():
            if excel_pid is not None and is_process_running(excel_pid):
                kill_process_by_pid(excel_pid)
            process.terminate()
            process.join(2.0)
