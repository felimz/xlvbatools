"""Tests for the config-bound public project facade."""

import json
from unittest.mock import patch

import pytest


@pytest.mark.unit
def test_project_inspection_returns_typed_contract(tmp_path):
    from xlvbatools import XlvbaProject

    workbook = tmp_path / "book.xlsm"
    project = XlvbaProject.for_workbook(workbook)
    screenshot = str(tmp_path / "input.png")
    legacy = {
        "success": True,
        "phase": "complete",
        "screenshots": {"Input": screenshot},
        "data": {"sheets": {"Input": {"cells": {}}}},
        "dialog_events": [],
        "cleanup": {
            "pid": 42,
            "quit_requested": True,
            "exited_gracefully": True,
            "force_terminated": False,
            "still_running": False,
        },
    }
    with patch(
        "xlvbatools.workbook.dumper.inspect_workbook", return_value=legacy,
    ) as inspect:
        result = project.inspect(
            ["Input"], cell_range="B2:C3", output_dir=tmp_path,
        )

    assert result.success is True
    assert result.data.screenshots == {"Input": screenshot}
    assert result.artifacts[0].metadata["sheet"] == "Input"
    assert result.require_clean_shutdown().pid == 42
    json.dumps(result.to_dict())
    inspect.assert_called_once_with(
        str(workbook.resolve()),
        ["Input"],
        output_dir=str(tmp_path),
        custom_range="B2:C3",
        include_data=True,
        include_screenshots=True,
        output_json=None,
        output_md=None,
        continue_on_render_error=False,
        include_hidden_sheets=False,
        timeout_seconds=60.0,
    )


@pytest.mark.unit
def test_project_macro_preserves_legacy_data_inside_stable_envelope(tmp_path):
    from xlvbatools import XlvbaProject

    project = XlvbaProject.for_workbook(tmp_path / "book.xlsm")
    legacy = {
        "success": True,
        "phase": "macro_execution",
        "macro": "Calculate",
        "elapsed_seconds": 0.2,
        "dialog_events": [],
        "cleanup": {
            "pid": 10,
            "quit_requested": True,
            "exited_gracefully": True,
            "force_terminated": False,
            "still_running": False,
        },
    }
    with patch("xlvbatools.macro.runner.run_macro", return_value=legacy) as run:
        result = project.run_macro("Calculate", timeout=5)

    assert result.success is True
    assert result.data["macro"] == "Calculate"
    assert result.metadata == {"macro": "Calculate"}
    run.assert_called_once_with(
        str((tmp_path / "book.xlsm").resolve()),
        "Calculate",
        named_ranges=None,
        timeout=5,
        visible=False,
        save_on_exit=True,
        strict_named_ranges=True,
    )


@pytest.mark.unit
def test_project_lint_uses_resolved_source_path(tmp_path):
    from xlvbatools import XlvbaProject

    source = tmp_path / "vba_source"
    source.mkdir()
    (source / "modTest.bas").write_text(
        "Option Explicit\nPublic Sub Test()\nEnd Sub\n", encoding="utf-8",
    )
    project = XlvbaProject.for_workbook(
        tmp_path / "book.xlsm", vba_source=source,
    )

    result = project.lint()

    assert result.phase == "complete"
    assert result.metadata["issue_count"] >= 0
    json.dumps(result.to_dict())


@pytest.mark.unit
@pytest.mark.parametrize(
    ("method", "kwargs", "operation", "worker_data"),
    [
        ("extract", {}, "extract", {"components": []}),
        ("inject", {"backup": False}, "inject", []),
        ("diff", {}, "diff", []),
        ("modify", {"cell": "A1", "value": 4}, "modify", True),
    ],
)
def test_com_backed_project_methods_use_shared_worker(
    tmp_path, method, kwargs, operation, worker_data,
):
    from xlvbatools import XlvbaProject

    project = XlvbaProject.for_workbook(tmp_path / "book.xlsm")
    legacy = {
        "success": True,
        "phase": "complete",
        "data": worker_data,
        "worker_pid": 90,
        "excel_pid": 91,
        "cleanup": {
            "pid": 91,
            "quit_requested": True,
            "exited_gracefully": True,
            "force_terminated": False,
            "still_running": False,
        },
    }
    with patch(
        "xlvbatools.core.worker.run_isolated_operation", return_value=legacy,
    ) as worker:
        result = getattr(project, method)(**kwargs)

    assert result.success is True
    assert result.diagnostics.worker_pid == 90
    assert result.require_clean_shutdown().pid == 91
    assert worker.call_args.args[0] == operation


@pytest.mark.unit
def test_project_from_nested_directory_resolves_config_paths(tmp_path):
    from xlvbatools import XlvbaProject

    (tmp_path / "xlvbatools.toml").write_text(
        "[xlvbatools]\n"
        'workbook = "workbook/book.xlsm"\n'
        'vba_source = "workbook/vba_source"\n'
        'snapshots_dir = "workbook/snapshots"\n'
        'log_dir = "logs"\n',
        encoding="utf-8",
    )
    nested = tmp_path / "tools" / "nested"
    nested.mkdir(parents=True)

    project = XlvbaProject.from_config(nested)

    assert project.workbook_path == str(
        (tmp_path / "workbook" / "book.xlsm").resolve()
    )
    assert project.vba_source_path == str(
        (tmp_path / "workbook" / "vba_source").resolve()
    )
    assert project.config.log_dir_path == str((tmp_path / "logs").resolve())


@pytest.mark.unit
def test_public_api_imports_do_not_import_win32com():
    import sys

    for name in tuple(sys.modules):
        if name.startswith("win32com"):
            del sys.modules[name]

    from xlvbatools import OperationResult, XlvbaProject, lint_files
    from xlvbatools.analysis import VBAIssue, lint_workbook

    assert OperationResult is not None
    assert XlvbaProject is not None
    assert lint_files is not None
    assert VBAIssue is not None
    assert lint_workbook is not None
    assert "win32com" not in sys.modules


@pytest.mark.com
@pytest.mark.integration
def test_project_facade_inspection_reports_clean_owned_process(minimal_workbook):
    from xlvbatools import XlvbaProject

    result = XlvbaProject.for_workbook(minimal_workbook).inspect(
        ["Sheet1"], include_data=True, include_screenshots=False, timeout=60,
    )

    assert result.success is True, result.to_dict()
    assert result.schema_version == "1.0"
    assert result.data.workbook_data["sheets"]["Sheet1"]
    assert result.diagnostics.cleanup.is_clean, result.to_dict()
    assert result.require_clean_shutdown().still_running is False


@pytest.mark.com
@pytest.mark.integration
def test_project_facade_macro_reports_clean_owned_process(runtime_error_workbook):
    from xlvbatools import XlvbaProject

    result = XlvbaProject.for_workbook(runtime_error_workbook).run_macro(
        "CompleteNormally", timeout=60, save_on_exit=False,
    )

    assert result.success is True, result.to_dict()
    assert result.data["macro"] == "CompleteNormally"
    assert result.require_clean_shutdown().still_running is False


@pytest.mark.com
@pytest.mark.e2e
def test_project_vba_round_trip_uses_clean_sequential_workers(
    runtime_error_workbook, tmp_path,
):
    """The high-level VBA workflow never exposes COM to its caller."""
    from xlvbatools import XlvbaProject

    source = tmp_path / "isolated_source"
    project = XlvbaProject.for_workbook(
        runtime_error_workbook, vba_source=source,
    )

    extracted = project.extract(timeout=90)
    assert extracted.success is True, extracted.to_dict()
    assert extracted.data["components"]
    assert extracted.require_clean_shutdown().still_running is False

    compared = project.diff(timeout=90)
    assert compared.success is True, compared.to_dict()
    assert all(item["status"] == "identical" for item in compared.data)
    assert compared.require_clean_shutdown().still_running is False

    injected = project.inject(backup=False, timeout=90)
    assert injected.success is True, injected.to_dict()
    assert all(item["status"] == "injected" for item in injected.data)
    assert injected.require_clean_shutdown().still_running is False

    linted = project.lint(workbook=True, compile_test=False, timeout=90)
    assert linted.success is True, linted.to_dict()
    assert linted.require_clean_shutdown().still_running is False

    executed = project.run_macro(
        "CompleteNormally", timeout=90, save_on_exit=False,
    )
    assert executed.success is True, executed.to_dict()
    assert executed.require_clean_shutdown().still_running is False


@pytest.mark.com
@pytest.mark.e2e
def test_project_modify_then_inspect_across_isolated_workers(
    minimal_workbook,
):
    from xlvbatools import XlvbaProject

    project = XlvbaProject.for_workbook(minimal_workbook)
    modified = project.modify(
        sheet="Sheet1", cell="B2", value=73, timeout=90,
    )
    assert modified.success is True, modified.to_dict()
    assert modified.require_clean_shutdown().still_running is False

    inspected = project.inspect(
        ["Sheet1"], cell_range="B2", include_screenshots=False, timeout=90,
    )
    assert inspected.success is True, inspected.to_dict()
    cell = inspected.data.workbook_data["sheets"]["Sheet1"]["cells"]["B2"]
    assert cell["value"] == 73
    assert inspected.require_clean_shutdown().still_running is False
