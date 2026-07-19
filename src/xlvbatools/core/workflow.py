"""Execute a validated ordered workflow in one owned Excel session."""

from __future__ import annotations

import time
from typing import Any, Mapping

from xlvbatools.core.session import ExcelSession
from xlvbatools.workflow import (
    WORKFLOW_SCHEMA_VERSION,
    _steps_from_payload,
    _steps_to_worker,
)


class _WorkflowAbort(Exception):
    """Leave the session context after a typed step failure."""


def _dialog_sequence(session: ExcelSession) -> int:
    return max(
        (event.sequence for event in session.dialog_events),
        default=0,
    )


def _dialog_events_after(
    session: ExcelSession,
    sequence: int,
) -> list[Mapping[str, Any]]:
    return [
        event.to_dict()
        for event in session.dialog_events
        if event.sequence > sequence
    ]


def _error(
    message: str,
    *,
    code: str,
    error_type: str | None = None,
    details: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    return {
        "message": message,
        "code": code,
        "error_type": error_type,
        "details": dict(details or {}),
    }


def _macro_step(
    session: ExcelSession,
    step: Mapping[str, Any],
    reporter: Any,
    *,
    index: int,
    count: int,
) -> tuple[dict[str, Any], str]:
    reporter.workflow_step(step, index=index, count=count, step_phase="named_range_setup")
    for name, value in (step.get("named_ranges") or {}).items():
        session.set_named_range(
            name,
            value,
            strict=bool(step.get("strict_named_ranges", True)),
        )
    reporter.workflow_step(step, index=index, count=count, step_phase="macro_execution")
    macro_result = session.run_macro(str(step["macro"]))
    envelope_keys = {
        "success", "phase", "primary_error", "error", "error_type",
        "traceback", "dialog_events", "cleanup", "com_error",
        "elapsed_seconds",
    }
    data = {
        key: value
        for key, value in macro_result.items()
        if key not in envelope_keys
    }
    data.setdefault("macro", step["macro"])
    data.setdefault("excel_pid", session.excel_pid)
    if not macro_result.get("success"):
        message = str(
            macro_result.get("primary_error")
            or macro_result.get("error")
            or f"Macro {step['macro']!r} failed"
        )
        raise _StepFailure(
            message,
            phase=str(macro_result.get("phase") or "macro_execution"),
            code="macro_failed",
            data=data,
            details={
                "com_error": macro_result.get("com_error"),
            },
        )
    return data, "complete"


def _modify_step(
    session: ExcelSession,
    step: Mapping[str, Any],
    reporter: Any,
    *,
    index: int,
    count: int,
) -> tuple[dict[str, Any], str]:
    from xlvbatools.workbook.modifier import _write_ranges_in_session

    reporter.workflow_step(step, index=index, count=count, step_phase="range_write")
    return (
        _write_ranges_in_session(
            session,
            sheet=str(step["sheet"]),
            values=step.get("values") or {},
            calculate=bool(step.get("calculate", False)),
        ),
        "complete",
    )


def _inspect_step(
    session: ExcelSession,
    workbook_path: str,
    step: Mapping[str, Any],
    reporter: Any,
    *,
    index: int,
    count: int,
) -> tuple[dict[str, Any], str]:
    from xlvbatools.workbook.dumper import _inspect_existing_session

    reporter.workflow_step(step, index=index, count=count, step_phase="inspection")
    return (
        _inspect_existing_session(
            session,
            workbook_path=workbook_path,
            sheets=list(step.get("sheets") or ()),
            output_dir=str(step.get("output_dir") or "screenshots"),
            custom_range=step.get("cell_range"),
            include_data=bool(step.get("include_data", True)),
            include_screenshots=bool(step.get("include_screenshots", False)),
            output_json=step.get("output_json"),
            output_md=step.get("output_markdown"),
            continue_on_render_error=bool(step.get("continue_on_render_error", False)),
            include_hidden_sheets=bool(step.get("include_hidden_sheets", False)),
        ),
        "complete",
    )


class _StepFailure(Exception):
    def __init__(
        self,
        message: str,
        *,
        phase: str,
        code: str,
        data: Mapping[str, Any] | None = None,
        details: Mapping[str, Any] | None = None,
    ) -> None:
        super().__init__(message)
        self.phase = phase
        self.code = code
        self.data = dict(data) if data is not None else None
        self.details = dict(details or {})


def _not_run(step: Mapping[str, Any]) -> dict[str, Any]:
    return {
        "id": step["id"],
        "kind": step["kind"],
        "status": "not_run",
        "phase": "not_run",
        "elapsed_seconds": 0.0,
        "data": None,
        "error": None,
        "dialog_events": [],
    }


def execute_workflow(arguments: Mapping[str, Any], reporter: Any) -> dict[str, Any]:
    """Run the complete request with fail-fast, explicit-save semantics."""
    allowed_arguments = {
        "workbook_path", "steps", "visible", "save_on_success",
        "workflow_schema_version",
    }
    unexpected_arguments = set(arguments) - allowed_arguments
    if unexpected_arguments:
        raise ValueError(
            "unknown workflow request fields: "
            + ", ".join(sorted(unexpected_arguments))
        )
    schema_version = arguments.get("workflow_schema_version")
    if schema_version != WORKFLOW_SCHEMA_VERSION:
        raise ValueError(
            f"Unsupported workflow schema {schema_version!r}; "
            f"expected {WORKFLOW_SCHEMA_VERSION!r}"
        )
    workbook_value = arguments.get("workbook_path")
    if not isinstance(workbook_value, str) or not workbook_value.strip():
        raise TypeError("workbook_path must be a non-empty string")
    visible = arguments.get("visible", False)
    save_requested = arguments.get("save_on_success", False)
    if not isinstance(visible, bool):
        raise TypeError("visible must be boolean")
    if not isinstance(save_requested, bool):
        raise TypeError("save_on_success must be boolean")
    typed_steps = _steps_from_payload(arguments.get("steps"))
    steps = _steps_to_worker(typed_steps)
    workbook_path = workbook_value.strip()
    session = ExcelSession(
        workbook_path,
        visible=visible,
        save_on_exit=False,
        kill_on_enter=False,
        read_only=False,
        disable_macros=False,
        on_excel_started=reporter.excel_started,
    )
    step_results: list[dict[str, Any]] = []
    failed_step_id: str | None = None
    primary_error: str | None = None
    outer_error: dict[str, Any] | None = None
    phase = "session_start"
    saved = False
    current_index = -1

    try:
        with session:
            for current_index, step in enumerate(steps):
                step_started = time.monotonic()
                event_sequence = _dialog_sequence(session)
                kind = str(step["kind"])
                try:
                    if kind == "macro":
                        data, step_phase = _macro_step(
                            session,
                            step,
                            reporter,
                            index=current_index,
                            count=len(steps),
                        )
                    elif kind == "modify":
                        data, step_phase = _modify_step(
                            session,
                            step,
                            reporter,
                            index=current_index,
                            count=len(steps),
                        )
                    else:
                        data, step_phase = _inspect_step(
                            session,
                            workbook_path,
                            step,
                            reporter,
                            index=current_index,
                            count=len(steps),
                        )
                    step_results.append({
                        "id": step["id"],
                        "kind": kind,
                        "status": "succeeded",
                        "phase": step_phase,
                        "elapsed_seconds": time.monotonic() - step_started,
                        "data": data,
                        "error": None,
                        "dialog_events": _dialog_events_after(session, event_sequence),
                    })
                except _StepFailure as error:
                    failed_step_id = str(step["id"])
                    primary_error = str(error)
                    phase = error.phase
                    step_results.append({
                        "id": step["id"],
                        "kind": kind,
                        "status": "failed",
                        "phase": error.phase,
                        "elapsed_seconds": time.monotonic() - step_started,
                        "data": error.data,
                        "error": _error(
                            str(error),
                            code=error.code,
                            error_type=type(error).__name__,
                            details=error.details,
                        ),
                        "dialog_events": _dialog_events_after(session, event_sequence),
                    })
                    outer_error = _error(
                        primary_error,
                        code="workflow_step_failed",
                        details={
                            "step_id": failed_step_id,
                            "step_kind": kind,
                            "step_index": current_index,
                            "step_error": step_results[-1]["error"],
                        },
                    )
                    raise _WorkflowAbort from error
                except Exception as error:
                    failed_step_id = str(step["id"])
                    primary_error = str(error)
                    progress = reporter.value
                    phase = str(progress.get("step_phase") or "workflow_step")
                    step_error = _error(
                        primary_error,
                        code="workflow_step_error",
                        error_type=type(error).__name__,
                    )
                    step_results.append({
                        "id": step["id"],
                        "kind": kind,
                        "status": "failed",
                        "phase": phase,
                        "elapsed_seconds": time.monotonic() - step_started,
                        "data": None,
                        "error": step_error,
                        "dialog_events": _dialog_events_after(session, event_sequence),
                    })
                    outer_error = _error(
                        primary_error,
                        code="workflow_step_failed",
                        error_type=type(error).__name__,
                        details={
                            "step_id": failed_step_id,
                            "step_kind": kind,
                            "step_index": current_index,
                            "step_error": step_error,
                        },
                    )
                    raise _WorkflowAbort from error

            if save_requested:
                reporter.phase("workbook_save")
                phase = "workbook_save"
                try:
                    session.save()
                    saved = True
                except Exception as error:
                    session.cleanup_result["workbook_save_error"] = str(error)
                    primary_error = str(error)
                    outer_error = _error(
                        primary_error,
                        code="workbook_save_failed",
                        error_type=type(error).__name__,
                    )
                    raise _WorkflowAbort from error
    except _WorkflowAbort:
        pass
    except Exception as error:
        primary_error = str(error)
        phase = session.phase
        outer_error = _error(
            primary_error,
            code="workflow_session_failed",
            error_type=type(error).__name__,
        )

    if len(step_results) < len(steps):
        step_results.extend(_not_run(step) for step in steps[len(step_results):])

    cleanup = dict(session.cleanup_result)
    success = outer_error is None
    if cleanup.get("workbook_close_error") and success:
        success = False
        phase = "workbook_close"
        primary_error = str(cleanup["workbook_close_error"])
        outer_error = _error(primary_error, code="workbook_close_failed")
    if cleanup.get("still_running"):
        success = False
        phase = "cleanup"
        primary_error = f"Owned Excel PID {session.excel_pid} remained running after cleanup"
        outer_error = _error(primary_error, code="cleanup_failed")
    if success:
        phase = "complete"

    return {
        "success": success,
        "phase": phase,
        "data": {
            "workflow_schema_version": WORKFLOW_SCHEMA_VERSION,
            "steps": step_results,
            "failed_step_id": failed_step_id,
            "save_requested": save_requested,
            "saved": saved,
        },
        "primary_error": primary_error,
        "error": outer_error,
        "dialog_events": [event.to_dict() for event in session.dialog_events],
        "cleanup": cleanup,
        "excel_pid": session.excel_pid,
    }
