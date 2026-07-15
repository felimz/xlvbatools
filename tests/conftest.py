"""
Shared test fixtures and markers for xlvbatools test suite.
"""

import os
import sys
import pytest
import gc
import time

# ── Marker registration ──

def pytest_configure(config):
    config.addinivalue_line("markers", "unit: Pure Python unit tests, no COM required")
    config.addinivalue_line("markers", "integration: Requires live Excel or multi-component integration")
    config.addinivalue_line("markers", "com: Requires Excel COM automation (Windows only)")
    config.addinivalue_line("markers", "e2e: Full end-to-end pipeline tests (slowest)")


# ── Path fixtures ──

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIXTURES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures")


def _require_owned_excel_exit(excel_pid: int, *, grace_period: float = 20.0) -> None:
    """Do not let a fixture-owned Excel process leak into the next test."""
    from xlvbatools.core.process import is_process_running, kill_process_by_pid

    deadline = time.time() + grace_period
    while time.time() < deadline and is_process_running(excel_pid):
        time.sleep(0.1)
    if is_process_running(excel_pid):
        kill_process_by_pid(excel_pid)
        deadline = time.time() + 10.0
        while time.time() < deadline and is_process_running(excel_pid):
            time.sleep(0.1)
    if is_process_running(excel_pid):
        pytest.fail(f"Fixture-owned Excel PID {excel_pid} did not exit")


@pytest.fixture
def fixtures_dir():
    """Path to the test fixtures directory."""
    return FIXTURES_DIR


@pytest.fixture
def minimal_workbook(tmp_path):
    """Path to a dynamically created minimal macro-enabled workbook."""
    if sys.platform != "win32":
        pytest.skip("COM tests require Windows")

    import win32com.client
    excel = None
    wb = None
    excel_pid = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        import win32process
        _, excel_pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Add()
        
        path = os.path.join(tmp_path, "minimal.xlsm")
        # FileFormat=52 is xlOpenXMLWorkbookMacroEnabled
        wb.SaveAs(path, FileFormat=52)
        wb.Close(False)
        wb = None
        return path
    except Exception as e:
        pytest.skip(f"Could not dynamically create minimal.xlsm: {e}")
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        wb = None
        excel = None
        gc.collect()
        if excel_pid is not None:
            # Excel commonly needs more than three seconds to finish COM/VBE
            # shutdown. Killing it mid-shutdown can destabilize the next
            # DispatchEx call in the same interpreter.
            _require_owned_excel_exit(excel_pid)


@pytest.fixture
def runtime_error_workbook(tmp_path):
    """Macro workbook that raises a deterministic multiline VBA error."""
    if sys.platform != "win32":
        pytest.skip("COM tests require Windows")
    import win32com.client

    excel = None
    wb = None
    module = None
    sheet = None
    formula_cell = None
    interior = None
    excel_pid = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        import win32process
        _, excel_pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Add()
        module = wb.VBProject.VBComponents.Add(1)
        module.Name = "modReliabilityTest"
        module.CodeModule.AddFromString(
            'Option Explicit\r\n'
            'Public Sub CompleteNormally()\r\n'
            '    ThisWorkbook.Worksheets(1).Range("A1").Value = 42\r\n'
            'End Sub\r\n'
            'Public Sub ShowMessage()\r\n'
            '    MsgBox "Watchdog test message", vbInformation\r\n'
            'End Sub\r\n'
            'Public Sub ShowFilePicker()\r\n'
            '    Dim selected As Variant\r\n'
            '    selected = Application.GetOpenFilename()\r\n'
            'End Sub\r\n'
            'Public Sub LoopForever()\r\n'
            '    Do\r\n'
            '    Loop\r\n'
            'End Sub\r\n'
            'Public Sub RaiseMultilineError()\r\n'
            '    Err.Raise vbObjectError + 100, "TestMacro", _\r\n'
            '        "Diagnostic line one." & vbCrLf & "Diagnostic line two."\r\n'
            'End Sub\r\n'
        )
        sheet = wb.Worksheets(1)
        formula_cell = sheet.Range("B1")
        formula_cell.Formula = "=21*2"
        formula_cell.NumberFormat = "0.00"
        interior = formula_cell.Interior
        interior.Color = 65535
        path = os.path.join(tmp_path, "runtime_error.xlsm")
        wb.SaveAs(path, FileFormat=52)
        interior = None
        formula_cell = None
        sheet = None
        module = None
        gc.collect()
        wb.Close(False)
        wb = None
        return path
    except Exception as error:
        pytest.skip(f"Could not create VBA integration workbook: {error}")
    finally:
        if wb is not None:
            try:
                wb.Close(False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        module = None
        interior = None
        formula_cell = None
        sheet = None
        wb = None
        excel = None
        gc.collect()
        if excel_pid is not None:
            _require_owned_excel_exit(excel_pid)


@pytest.fixture
def compile_error_workbook(runtime_error_workbook, tmp_path):
    """Workbook with a deterministic Option Explicit compile failure."""
    import shutil
    from xlvbatools.core.session import ExcelSession

    path = tmp_path / "compile_error.xlsm"
    shutil.copy2(runtime_error_workbook, path)
    with ExcelSession(str(path), save_on_exit=True, kill_on_enter=False) as session:
        component = session.vb_project.VBComponents.Add(1)
        component.Name = "modCompileFailure"
        component.CodeModule.AddFromString(
            "Option Explicit\r\n"
            "Public Sub CompileFailure()\r\n"
            "    undeclaredCompileValue = 1\r\n"
            "End Sub\r\n"
        )
        del component
    return str(path)


@pytest.fixture
def temp_vba_source(tmp_path):
    """Create a temporary vba_source directory structure for testing."""
    vba_dir = tmp_path / "vba_source"
    (vba_dir / "modules").mkdir(parents=True)
    (vba_dir / "classes").mkdir(parents=True)
    (vba_dir / "sheets").mkdir(parents=True)
    return vba_dir


@pytest.fixture
def sample_bas_file(temp_vba_source):
    """Create a sample .bas file for testing."""
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
    """Create a sample .cls file for testing."""
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
