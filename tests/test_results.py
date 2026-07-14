"""Public operation-result contract tests."""

import json

import pytest


@pytest.mark.unit
def test_operation_result_serializes_and_requires_clean_shutdown():
    from xlvbatools.results import InspectionOutput, OperationResult

    result = OperationResult.from_legacy(
        "inspect",
        {
            "success": True,
            "phase": "complete",
            "dialog_events": [],
            "cleanup": {
                "pid": 42,
                "quit_requested": True,
                "exited_gracefully": True,
                "force_terminated": False,
                "still_running": False,
            },
        },
        data=InspectionOutput(
            workbook_data={"sheets": {}},
            screenshots={"Input": "input.png"},
        ),
    )

    assert result.require_success() == result.data
    assert result.require_clean_shutdown().pid == 42
    payload = result.to_dict()
    assert payload["schema_version"] == "1.0"
    assert payload["data"]["screenshots"] == {"Input": "input.png"}
    json.dumps(payload)


@pytest.mark.unit
def test_failed_result_raises_public_operation_error():
    from xlvbatools.errors import OperationFailedError
    from xlvbatools.results import OperationResult

    result = OperationResult.from_legacy(
        "run_macro",
        {
            "success": False,
            "phase": "macro_execution",
            "primary_error": "Execution timed out",
            "timed_out": True,
        },
    )

    assert result.error.code == "timeout"
    with pytest.raises(OperationFailedError, match="Execution timed out"):
        result.require_success()


@pytest.mark.unit
def test_forced_cleanup_is_not_reported_as_clean():
    from xlvbatools.errors import HeadlessCleanupError
    from xlvbatools.results import OperationResult

    result = OperationResult.from_legacy(
        "inspect",
        {
            "success": True,
            "phase": "complete",
            "cleanup": {
                "pid": 99,
                "quit_requested": True,
                "exited_gracefully": False,
                "force_terminated": True,
                "still_running": False,
            },
        },
    )

    with pytest.raises(HeadlessCleanupError, match="did not exit cleanly"):
        result.require_clean_shutdown()
