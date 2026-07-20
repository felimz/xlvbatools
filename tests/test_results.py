"""Public operation-result contract tests."""

import json

import pytest


@pytest.mark.unit
def test_operation_result_serializes_and_requires_clean_shutdown():
    from xlvbatools.results import (
        CleanupReport,
        Diagnostics,
        InspectionOutput,
        OperationResult,
    )

    result = OperationResult(
        operation="inspect",
        success=True,
        phase="complete",
        data=InspectionOutput(
            workbook_data={"sheets": {}},
            screenshots={"Input": "input.png"},
        ),
        diagnostics=Diagnostics(
            cleanup=CleanupReport(
                pid=42,
                quit_requested=True,
                exited_gracefully=True,
            )
        ),
    )

    assert result.require_success() == result.data
    assert result.require_clean_shutdown().pid == 42
    payload = result.to_dict()
    assert payload["schema_version"] == "1.3"
    assert payload["data"]["screenshots"] == {"Input": "input.png"}
    json.dumps(payload)


@pytest.mark.unit
def test_failed_result_raises_public_operation_error():
    from xlvbatools.errors import OperationFailedError
    from xlvbatools.results import ErrorInfo, OperationResult

    result = OperationResult(
        operation="run_macro",
        success=False,
        phase="macro_execution",
        error=ErrorInfo(message="Execution timed out", code="timeout"),
    )

    assert result.error.code == "timeout"
    with pytest.raises(OperationFailedError, match="Execution timed out"):
        result.require_success()


@pytest.mark.unit
def test_forced_cleanup_is_not_reported_as_clean():
    from xlvbatools.errors import HeadlessCleanupError
    from xlvbatools.results import CleanupReport, Diagnostics, OperationResult

    result = OperationResult(
        operation="inspect",
        success=True,
        phase="complete",
        diagnostics=Diagnostics(
            cleanup=CleanupReport(
                pid=99,
                quit_requested=True,
                force_terminated=True,
            )
        ),
    )

    with pytest.raises(HeadlessCleanupError, match="did not exit cleanly"):
        result.require_clean_shutdown()
