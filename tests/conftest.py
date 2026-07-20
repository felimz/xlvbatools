"""Suite taxonomy plus offline and disposable-workbook fixtures."""

from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

import pytest


ROOT = Path(__file__).resolve().parents[1]
FIXTURES_DIR = Path(__file__).resolve().parent / "fixtures"
WORKBOOK_FACTORY = Path(__file__).resolve().parent / "_workbook_factory.py"

# Every collected test belongs to exactly one execution tier. Secondary
# markers describe scheduling only. This collection-time invariant prevents a
# live Excel test from silently entering the default fast suite.
PRIMARY_TEST_TIERS = frozenset({
    "unit", "integration", "excel", "distribution", "external",
})
LIVE_EXCEL_FIXTURES = frozenset({
    "minimal_workbook",
    "runtime_error_workbook",
    "compile_error_workbook",
    "duplicate_declaration_workbook",
    "startup_event_workbook",
})


def pytest_addoption(parser):
    parser.addoption(
        "--external-workbook",
        action="append",
        default=[],
        metavar="PATH",
        help=(
            "explicit downstream workbook for the opt-in external acceptance "
            "suite; may be repeated"
        ),
    )


@pytest.hookimpl(tryfirst=True)
def pytest_collection_modifyitems(items):
    errors = []
    for item in items:
        tiers = sorted(
            marker for marker in PRIMARY_TEST_TIERS
            if item.get_closest_marker(marker) is not None
        )
        if len(tiers) != 1:
            errors.append(
                f"{item.nodeid}: expected exactly one primary test tier, got {tiers}"
            )
            continue
        tier = tiers[0]
        if item.get_closest_marker("stress") is not None and tier != "excel":
            errors.append(f"{item.nodeid}: stress tests must belong to the excel tier")
        if LIVE_EXCEL_FIXTURES.intersection(item.fixturenames) and tier != "excel":
            errors.append(
                f"{item.nodeid}: live Excel fixture is not isolated in the excel tier"
            )
    if errors:
        raise pytest.UsageError("Invalid test taxonomy:\n" + "\n".join(errors))


def _build_workbook(kind: str, output: Path, *, marker: Path | None = None) -> str:
    """Build a fixture in a child process so pytest never owns factory COM."""
    if sys.platform != "win32":
        pytest.skip("Live Excel tests require Windows")

    command = [
        sys.executable,
        "-X",
        "faulthandler",
        str(WORKBOOK_FACTORY),
        "--kind",
        kind,
        "--output",
        str(output),
    ]
    if marker is not None:
        command.extend(("--marker", str(marker)))

    with tempfile.TemporaryFile(mode="w+", encoding="utf-8", errors="replace") as log:
        try:
            completed = subprocess.run(
                command,
                cwd=ROOT,
                stdout=log,
                stderr=subprocess.STDOUT,
                timeout=120,
            )
        except subprocess.TimeoutExpired:
            pytest.fail(f"Timed out while creating {kind} Excel fixture")
        log.seek(0)
        output_text = log.read()

    if completed.returncode != 0 or not output.is_file():
        pytest.fail(
            f"Could not create {kind} Excel fixture (exit "
            f"{completed.returncode}):\n{output_text}"
        )
    return str(output)


@pytest.fixture
def fixtures_dir():
    return str(FIXTURES_DIR)


@pytest.fixture(scope="session")
def _minimal_workbook_template(tmp_path_factory):
    path = tmp_path_factory.mktemp("excel_templates") / "minimal.xlsm"
    return _build_workbook("minimal", path)


@pytest.fixture
def minimal_workbook(_minimal_workbook_template, tmp_path):
    path = tmp_path / "minimal.xlsm"
    shutil.copy2(_minimal_workbook_template, path)
    return str(path)


@pytest.fixture(scope="session")
def _runtime_error_workbook_template(tmp_path_factory):
    path = tmp_path_factory.mktemp("excel_runtime_templates") / "runtime_error.xlsm"
    return _build_workbook("runtime", path)


@pytest.fixture
def runtime_error_workbook(_runtime_error_workbook_template, tmp_path):
    path = tmp_path / "runtime_error.xlsm"
    shutil.copy2(_runtime_error_workbook_template, path)
    return str(path)


@pytest.fixture(scope="session")
def _compile_error_workbook_template(tmp_path_factory):
    path = tmp_path_factory.mktemp("excel_compile_templates") / "compile_error.xlsm"
    return _build_workbook("compile_error", path)


@pytest.fixture
def compile_error_workbook(_compile_error_workbook_template, tmp_path):
    path = tmp_path / "compile_error.xlsm"
    shutil.copy2(_compile_error_workbook_template, path)
    return str(path)


@pytest.fixture(scope="session")
def _duplicate_declaration_workbook_template(tmp_path_factory):
    path = (
        tmp_path_factory.mktemp("excel_duplicate_templates")
        / "duplicate_declaration.xlsm"
    )
    return _build_workbook("duplicate_declaration", path)


@pytest.fixture
def duplicate_declaration_workbook(
    _duplicate_declaration_workbook_template, tmp_path,
):
    path = tmp_path / "duplicate_declaration.xlsm"
    shutil.copy2(_duplicate_declaration_workbook_template, path)
    return str(path)


@pytest.fixture
def startup_event_workbook(tmp_path):
    path = tmp_path / "startup_event.xlsm"
    marker = tmp_path / "workbook_open_ran.txt"
    return _build_workbook("startup_event", path, marker=marker), marker


@pytest.fixture
def temp_vba_source(tmp_path):
    vba_dir = tmp_path / "vba_source"
    (vba_dir / "modules").mkdir(parents=True)
    (vba_dir / "classes").mkdir(parents=True)
    (vba_dir / "sheets").mkdir(parents=True)
    return vba_dir


@pytest.fixture
def sample_bas_file(temp_vba_source):
    content = '''Attribute VB_Name = "modTest"
Option Explicit

Public Sub TestSub()
    Dim x As Double
    x = 42.0
    Debug.Print x
End Sub

Public Function TestFunc(ByVal a As Long) As Long
    TestFunc = a * 2
End Function
'''
    path = temp_vba_source / "modules" / "modTest.bas"
    path.write_text(content, encoding="utf-8")
    return path


@pytest.fixture
def sample_cls_file(temp_vba_source):
    content = '''VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Value As Double

Public Property Get Value() As Double
    Value = m_Value
End Property

Public Property Let Value(ByVal newValue As Double)
    m_Value = newValue
End Property
'''
    path = temp_vba_source / "classes" / "clsTest.cls"
    path.write_text(content, encoding="utf-8")
    return path
