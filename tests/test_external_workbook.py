"""Explicit downstream-workbook acceptance; never part of the library suite."""

from __future__ import annotations

from pathlib import Path
import sys

import pytest


pytestmark = [
    pytest.mark.external,
    pytest.mark.skipif(sys.platform != "win32", reason="Windows only"),
]


def _requested_workbooks(config: pytest.Config) -> list[Path]:
    return [Path(value).expanduser().resolve() for value in config.getoption(
        "--external-workbook"
    )]


def pytest_generate_tests(metafunc: pytest.Metafunc) -> None:
    if "external_workbook" not in metafunc.fixturenames:
        return
    workbooks = _requested_workbooks(metafunc.config)
    if not workbooks:
        workbooks = [pytest.param(
            None,
            marks=pytest.mark.skip(reason="pass --external-workbook PATH explicitly"),
            id="no-workbook-requested",
        )]
    metafunc.parametrize(
        "external_workbook",
        workbooks,
        ids=lambda path: Path(path).name if path is not None else None,
    )


def test_external_workbook_extracts_and_analyzes_without_project_coupling(
    external_workbook: Path,
    tmp_path: Path,
) -> None:
    """Exercise extraction and offline analysis on an explicitly owned input."""
    from xlvbatools import Project
    from xlvbatools.analysis.rules import run_all_rules
    from xlvbatools.vba.dependency import build_call_graph

    if not external_workbook.is_file():
        pytest.fail(f"External workbook does not exist: {external_workbook}")
    if external_workbook.suffix.lower() not in {".xlsm", ".xlsb", ".xls"}:
        pytest.fail(f"Unsupported external workbook type: {external_workbook}")

    source_dir = tmp_path / "extracted"
    extracted = Project.open(external_workbook, source=source_dir).extract(timeout=240)
    manifest = extracted.require_success()
    extracted.require_clean_shutdown()
    assert manifest.components, f"No VBA components extracted from {external_workbook}"

    vba_files = sorted(
        path for path in source_dir.rglob("*")
        if path.suffix.lower() in {".bas", ".cls", ".frm"}
    )
    issues = []
    for path in vba_files:
        lines = path.read_text(encoding="utf-8", errors="ignore").splitlines(True)
        issues.extend(run_all_rules(str(path.relative_to(source_dir)), lines))

    graph = build_call_graph(str(source_dir)).to_dict()
    assert isinstance(graph, dict)
    assert isinstance(issues, list)
