"""Execute complete Excel sessions through the shared isolated worker."""

import logging
import os
import time
import uuid
from typing import Callable, Optional

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
            allow_workbook_events=True,
            allow_macro_execution=True,
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
    from xlvbatools.core.worker import execute_worker_request

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
    return execute_worker_request("run_macro", arguments, timeout=timeout)
