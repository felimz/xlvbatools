"""Spawn-clean entry point for the shared isolated Excel worker protocol."""

from __future__ import annotations

import dataclasses
import json
import os
import sys
import time
import traceback
from pathlib import Path
from typing import Any

from xlvbatools.core.protocol import WORKER_PROTOCOL_VERSION


def _json_default(value: Any) -> Any:
    if hasattr(value, "to_dict") and callable(value.to_dict):
        return value.to_dict()
    if dataclasses.is_dataclass(value) and not isinstance(value, type):
        return dataclasses.asdict(value)
    if isinstance(value, Path):
        return str(value)
    raise TypeError(f"Object of type {type(value).__name__} is not JSON serializable")


def _atomic_json(path: str, value: Any) -> None:
    temp_path = path + ".tmp"
    with open(temp_path, "w", encoding="utf-8") as handle:
        json.dump(value, handle, default=_json_default)
    # On Windows a reader can briefly prevent replacement of the progress
    # file.  Retrying preserves atomic reads without turning benign polling
    # overlap into an operation failure.
    for attempt in range(50):
        try:
            os.replace(temp_path, path)
            return
        except PermissionError:
            if attempt == 49:
                raise
            time.sleep(0.01)


class ProgressReporter:
    def __init__(self, path: str, request_id: str, operation: str):
        self.path = path
        self.value: dict[str, Any] = {
            "protocol_version": WORKER_PROTOCOL_VERSION,
            "request_id": request_id,
            "operation": operation,
            "worker_pid": os.getpid(),
            "phase": "worker_start",
            "excel_pid": None,
        }
        self.write()

    def write(self) -> None:
        _atomic_json(self.path, self.value)

    def phase(self, phase: str) -> None:
        self.value["phase"] = phase
        self.write()

    def excel_started(self, pid: int) -> None:
        self.value["excel_pid"] = pid
        self.value["phase"] = "workbook_open"
        self.write()


def _session_result(
    operation: str,
    arguments: dict[str, Any],
    reporter: ProgressReporter,
) -> dict[str, Any]:
    # This durable progress event is the retry safety boundary. Publish it
    # before importing or constructing anything that could activate COM.
    reporter.phase("session_start")
    from xlvbatools.core.session import ExcelSession

    workbook_path = arguments["workbook_path"]
    save_on_exit = operation in {"inject", "modify"}
    read_only = operation in {
        "list_components", "extract", "diff", "lint_workbook",
    }
    session = ExcelSession(
        workbook_path,
        visible=False,
        save_on_exit=save_on_exit,
        kill_on_enter=False,
        read_only=read_only,
        disable_macros=read_only,
        on_excel_started=reporter.excel_started,
    )
    phase = "session_start"
    data: Any = None
    success = True
    primary_error = None
    try:
        with session:
            phase = operation
            reporter.phase(phase)
            if operation == "list_components":
                from xlvbatools.vba.extractor import list_components

                data = list_components(workbook_path, _session=session)
            elif operation == "extract":
                from xlvbatools.vba.extractor import extract_all, extract_component

                component = arguments.get("component")
                if component:
                    data = extract_component(
                        workbook_path, component, arguments["output_dir"],
                        _session=session,
                    )
                    success = data is not None
                    if not success:
                        primary_error = f"Component not found: {component}"
                else:
                    data = extract_all(
                        workbook_path, arguments["output_dir"], _session=session,
                    )
            elif operation == "inject":
                from xlvbatools.vba.injector import inject_all, inject_component

                component = arguments.get("component")
                if component:
                    data = inject_component(
                        workbook_path,
                        arguments["source_dir"],
                        component,
                        backup=arguments.get("backup", True),
                        _session=session,
                        _raise_on_error=True,
                    )
                    success = bool(data)
                else:
                    data = inject_all(
                        workbook_path,
                        arguments["source_dir"],
                        backup=arguments.get("backup", True),
                        dry_run=False,
                        backup_limit=arguments.get("backup_limit", 5),
                        _session=session,
                        _raise_on_error=True,
                    )
                    success = all(
                        item.get("status") != "error" for item in data
                    )
                if not success:
                    primary_error = "One or more VBA components failed to inject"
            elif operation == "diff":
                from xlvbatools.vba.differ import diff_all, diff_component

                component = arguments.get("component")
                if component:
                    data = diff_component(
                        workbook_path,
                        arguments["source_dir"],
                        component,
                        _session=session,
                    )
                    success = data is not None
                    if not success:
                        primary_error = f"Component not found: {component}"
                else:
                    data = diff_all(
                        workbook_path, arguments["source_dir"], _session=session,
                    )
            elif operation == "modify":
                from xlvbatools.workbook.modifier import modify_cell

                kwargs = dict(arguments)
                kwargs.pop("workbook_path", None)
                data = modify_cell(
                    workbook_path,
                    _session=session,
                    _raise_on_error=True,
                    **kwargs,
                )
                success = bool(data)
                if not success:
                    primary_error = "Workbook modification failed"
            elif operation == "lint_workbook":
                from xlvbatools.analysis.preflight import lint_workbook

                data = lint_workbook(
                    workbook_path,
                    disabled_rules=arguments.get("disabled_rules"),
                    compile_test=arguments.get("compile_test", True),
                    _session=session,
                )
                success = not any(issue.severity == "ERROR" for issue in data)
                if not success:
                    primary_error = "Static analysis found workbook errors"
            else:
                raise ValueError(f"Unsupported session operation: {operation}")
        phase = "complete" if success else operation
    except Exception as error:
        success = False
        primary_error = str(error)
        phase = session.phase if phase == "session_start" else phase

    cleanup = dict(session.cleanup_result)
    if cleanup.get("workbook_save_error"):
        success = False
        phase = "workbook_save"
        primary_error = cleanup["workbook_save_error"]
    elif cleanup.get("still_running"):
        success = False
        phase = "cleanup"
        primary_error = (
            f"Owned Excel PID {session.excel_pid} remained running after cleanup"
        )

    return {
        "success": success,
        "phase": phase,
        "data": data,
        "primary_error": primary_error,
        "dialog_events": [event.to_dict() for event in session.dialog_events],
        "cleanup": cleanup,
        "excel_pid": session.excel_pid,
    }


def _dispatch(
    operation: str,
    arguments: dict[str, Any],
    reporter: ProgressReporter,
) -> dict[str, Any]:
    if operation == "inspect":
        reporter.phase("session_start")
        from xlvbatools.workbook.dumper import _inspect_workbook_in_process

        result = _inspect_workbook_in_process(
            **arguments, on_excel_started=reporter.excel_started,
        )
        workbook_data = result.get("data")
        screenshots = dict(result.pop("screenshots", {}) or {})
        result["data"] = {
            "workbook_data": workbook_data,
            "screenshots": screenshots,
        }
        return result

    if operation == "run_macro":
        reporter.phase("session_start")
        from xlvbatools.macro.runner import _run_macro_once

        result = _run_macro_once(
            **arguments,
            on_excel_started=reporter.excel_started,
            on_phase=reporter.phase,
        )
        envelope_keys = {
            "success", "phase", "primary_error", "error_type", "traceback",
            "dialog_events", "cleanup", "excel_pid", "elapsed_seconds",
        }
        operation_data = {
            key: value for key, value in result.items()
            if key not in envelope_keys
        }
        for key in operation_data:
            result.pop(key, None)
        result["data"] = operation_data
        return result

    if operation == "inject" and arguments.get("dry_run"):
        from xlvbatools.vba.injector import inject_all

        reporter.phase("dry_run")
        data = inject_all(
            arguments["workbook_path"],
            arguments["source_dir"],
            backup=arguments.get("backup", True),
            dry_run=True,
            backup_limit=arguments.get("backup_limit", 5),
        )
        return {
            "success": True,
            "phase": "complete",
            "data": data,
            "dialog_events": [],
            "cleanup": {},
        }

    return _session_result(operation, arguments, reporter)


def main() -> int:
    request_path, result_path, progress_path = sys.argv[1:4]
    started = time.monotonic()
    with open(request_path, encoding="utf-8") as handle:
        request = json.load(handle)

    operation = request.get("operation", "unknown")
    request_id = request.get("request_id", "")
    reporter = ProgressReporter(progress_path, request_id, operation)
    result: dict[str, Any]
    try:
        if request.get("protocol_version") != WORKER_PROTOCOL_VERSION:
            raise ValueError(
                "Unsupported worker protocol version: "
                f"{request.get('protocol_version')!r}"
            )
        arguments = request.get("arguments")
        if not isinstance(arguments, dict):
            raise ValueError("Worker arguments must be a JSON object")
        result = _dispatch(operation, arguments, reporter)
    except BaseException as error:
        result = {
            "success": False,
            "phase": reporter.value.get("phase", "worker_error"),
            "primary_error": f"{type(error).__name__}: {error}",
            "error_type": type(error).__name__,
            "traceback": traceback.format_exc(),
            "dialog_events": [],
            "cleanup": {},
        }

    result.update({
        "protocol_version": WORKER_PROTOCOL_VERSION,
        "request_id": request_id,
        "operation": operation,
        "worker_pid": os.getpid(),
        "excel_pid": result.get("excel_pid") or reporter.value.get("excel_pid"),
        "elapsed_seconds": result.get(
            "elapsed_seconds", time.monotonic() - started,
        ),
    })
    reporter.phase("result_write")
    _atomic_json(result_path, result)
    return 0 if result.get("success") else 1


if __name__ == "__main__":
    raise SystemExit(main())
