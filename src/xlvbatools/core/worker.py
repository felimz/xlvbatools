"""Parent-side executor for isolated Excel operations.

The protocol deliberately uses ordinary JSON files rather than inherited
pipes.  Excel and pywin32 can retain duplicated pipe handles during teardown,
which makes an otherwise-finished parent wait indefinitely for end-of-file.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import time
import uuid
from typing import Any, Mapping

from xlvbatools.core.process import is_process_running, kill_process_by_pid


WORKER_PROTOCOL_VERSION = "1.0"


def _worker_python() -> tuple[str, dict[str, str] | None]:
    """Return a directly trackable interpreter preserving the active venv.

    Python's Windows venv executable can be a launcher that creates a second
    process.  Tracking that launcher defeats hard timeouts because it may exit
    while the real interpreter and Excel keep running.  Launching the base
    executable with ``__PYVENV_LAUNCHER__`` retains the venv's prefix and site
    packages without the extra process layer.
    """
    base_executable = getattr(sys, "_base_executable", None)
    if (
        os.name == "nt"
        and isinstance(base_executable, str)
        and os.path.isfile(base_executable)
        and os.path.normcase(base_executable) != os.path.normcase(sys.executable)
    ):
        environment = os.environ.copy()
        environment["__PYVENV_LAUNCHER__"] = sys.executable
        return base_executable, environment
    return sys.executable, None


def _read_json(path: str) -> dict[str, Any] | None:
    try:
        with open(path, encoding="utf-8") as handle:
            value = json.load(handle)
        return value if isinstance(value, dict) else None
    except (OSError, ValueError, TypeError):
        return None


def _terminate_owned_processes(
    process: subprocess.Popen,
    excel_pid: int | None,
    *,
    grace_period: float = 2.0,
) -> dict[str, Any]:
    """Stop the reported Excel process first, then its worker if necessary."""
    force_terminated = False
    if excel_pid is not None and is_process_running(excel_pid):
        force_terminated = kill_process_by_pid(excel_pid)

    try:
        process.wait(timeout=grace_period)
    except subprocess.TimeoutExpired:
        pass

    worker_terminated = False
    if process.poll() is None:
        process.terminate()
        worker_terminated = True
        try:
            process.wait(timeout=grace_period)
        except subprocess.TimeoutExpired:
            process.kill()
            process.wait(timeout=grace_period)

    return {
        "pid": excel_pid,
        "quit_requested": False,
        "exited_gracefully": False,
        "force_terminated": force_terminated,
        "worker_terminated": worker_terminated,
        "still_running": bool(
            excel_pid is not None and is_process_running(excel_pid)
        ),
    }


def _is_transient_com_failure(result: Mapping[str, Any]) -> bool:
    cleanup = result.get("cleanup") or {}
    if isinstance(cleanup, Mapping) and cleanup.get("still_running"):
        return False
    message = " ".join(
        str(result.get(key) or "")
        for key in ("primary_error", "worker_output")
    ).lower()
    return any(marker in message for marker in (
        "0x800706ba", "-2147023174", "rpc server is unavailable",
        "0x80010001", "-2147418111", "call was rejected by callee",
        "0x80010108", "-2147417848", "object invoked has disconnected",
    ))


def run_isolated_operation(
    operation: str,
    arguments: Mapping[str, Any],
    *,
    timeout: float = 60.0,
    retry_transient: bool = False,
) -> dict[str, Any]:
    """Execute one Excel operation in a killable, spawn-clean interpreter."""
    if timeout <= 0:
        raise ValueError("timeout must be greater than zero")
    if not operation or not isinstance(operation, str):
        raise ValueError("operation must be a non-empty string")

    attempts = 2 if retry_transient else 1
    result: dict[str, Any] | None = None
    for attempt in range(1, attempts + 1):
        result = _run_isolated_operation_once(
            operation, arguments, timeout=timeout,
        )
        result["attempt_count"] = attempt
        if result.get("success") or not _is_transient_com_failure(result):
            return result
    assert result is not None
    return result


def _run_isolated_operation_once(
    operation: str,
    arguments: Mapping[str, Any],
    *,
    timeout: float,
) -> dict[str, Any]:
    request_id = str(uuid.uuid4())
    request = {
        "protocol_version": WORKER_PROTOCOL_VERSION,
        "request_id": request_id,
        "operation": operation,
        "arguments": dict(arguments),
    }
    started = time.monotonic()

    with tempfile.TemporaryDirectory(prefix="xlvba-worker-") as temp_dir:
        request_path = os.path.join(temp_dir, "request.json")
        result_path = os.path.join(temp_dir, "result.json")
        progress_path = os.path.join(temp_dir, "progress.json")
        log_path = os.path.join(temp_dir, "worker.log")
        with open(request_path, "w", encoding="utf-8") as handle:
            json.dump(request, handle)

        with open(log_path, "w", encoding="utf-8") as worker_log:
            worker_python, worker_environment = _worker_python()
            process = subprocess.Popen(
                [
                    worker_python,
                    "-m",
                    "xlvbatools.core.worker_entry",
                    request_path,
                    result_path,
                    progress_path,
                ],
                stdin=subprocess.DEVNULL,
                stdout=worker_log,
                stderr=subprocess.STDOUT,
                env=worker_environment,
            )

            progress: dict[str, Any] = {
                "worker_pid": process.pid,
                "phase": "worker_start",
            }
            deadline = started + timeout
            while process.poll() is None and time.monotonic() < deadline:
                latest = _read_json(progress_path)
                if latest:
                    progress.update(latest)
                time.sleep(0.05)

            if process.poll() is None:
                latest = _read_json(progress_path)
                if latest:
                    progress.update(latest)
                excel_pid = progress.get("excel_pid")
                cleanup = _terminate_owned_processes(process, excel_pid)
                return {
                    "protocol_version": WORKER_PROTOCOL_VERSION,
                    "request_id": request_id,
                    "operation": operation,
                    "success": False,
                    "phase": progress.get("phase", "timeout"),
                    "timed_out": True,
                    "timeout_seconds": timeout,
                    "elapsed_seconds": time.monotonic() - started,
                    "worker_pid": process.pid,
                    "executor_pid": process.pid,
                    "excel_pid": excel_pid,
                    "primary_error": (
                        f"{operation} exceeded {timeout:.3f} seconds"
                    ),
                    "dialog_events": [],
                    "cleanup": cleanup,
                }

        result = _read_json(result_path)
        if result is None:
            output = ""
            try:
                with open(log_path, encoding="utf-8", errors="replace") as handle:
                    output = handle.read()[-4000:]
            except OSError:
                pass
            cleanup = _terminate_owned_processes(
                process, progress.get("excel_pid"), grace_period=0.5,
            )
            return {
                "protocol_version": WORKER_PROTOCOL_VERSION,
                "request_id": request_id,
                "operation": operation,
                "success": False,
                "phase": progress.get("phase", "worker_exit"),
                "elapsed_seconds": time.monotonic() - started,
                "worker_pid": process.pid,
                "executor_pid": process.pid,
                "excel_pid": progress.get("excel_pid"),
                "primary_error": (
                    f"Worker exited without a result (exit code {process.returncode})"
                ),
                "worker_output": output,
                "dialog_events": [],
                "cleanup": cleanup,
            }

        result.setdefault("protocol_version", WORKER_PROTOCOL_VERSION)
        result.setdefault("request_id", request_id)
        result.setdefault("operation", operation)
        result.setdefault("worker_pid", process.pid)
        result["executor_pid"] = process.pid
        result.setdefault("elapsed_seconds", time.monotonic() - started)
        if not result.get("success") and "worker_output" not in result:
            try:
                with open(log_path, encoding="utf-8", errors="replace") as handle:
                    output = handle.read()[-4000:]
                if output:
                    result["worker_output"] = output
            except OSError:
                pass
        return result
