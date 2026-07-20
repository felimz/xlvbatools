"""Unit contract for the repository's explicit test-tier runner."""

import pytest

from scripts.test import build_pytest_arguments


pytestmark = pytest.mark.unit


def test_fast_suite_relies_on_safe_pytest_default():
    assert build_pytest_arguments("fast") == ["-m", "pytest"]


@pytest.mark.parametrize(
    ("suite", "expression"),
    [
        ("excel-smoke", "excel and smoke"),
        ("excel", "excel and not stress"),
        ("stress", "excel and stress"),
        ("distribution", "distribution"),
        ("all", "not external"),
    ],
)
def test_named_suite_maps_to_one_marker_expression(suite, expression):
    assert build_pytest_arguments(suite) == ["-m", "pytest", "-m", expression]


def test_external_suite_requires_and_resolves_explicit_workbook(tmp_path):
    workbook = tmp_path / "Model.xlsm"
    arguments = build_pytest_arguments(
        "external", external_workbooks=(str(workbook),),
    )
    assert arguments == [
        "-m", "pytest", "-m", "external",
        "tests/test_external_workbook.py",
        "--external-workbook", str(workbook.resolve()),
    ]


def test_external_workbook_is_rejected_for_library_suites():
    with pytest.raises(ValueError, match="only for the external suite"):
        build_pytest_arguments("fast", external_workbooks=("Model.xlsm",))


def test_coverage_and_extra_arguments_are_forwarded():
    arguments = build_pytest_arguments(
        "unit", coverage=True, extra=("-q", "tests/test_results.py"),
    )
    assert arguments == [
        "-m", "pytest", "-m", "unit",
        "--cov=xlvbatools",
        "--cov-report=term-missing",
        "--cov-fail-under=60",
        "-q",
        "tests/test_results.py",
    ]
