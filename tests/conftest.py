"""
Shared test fixtures and markers for xlvbatools test suite.
"""

import os
import pytest

# ── Marker registration ──

def pytest_configure(config):
    config.addinivalue_line("markers", "unit: Pure Python unit tests, no COM required")
    config.addinivalue_line("markers", "com: Requires Excel COM automation (Windows only)")
    config.addinivalue_line("markers", "e2e: Full end-to-end pipeline tests (slowest)")


# ── Path fixtures ──

PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FIXTURES_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fixtures")


@pytest.fixture
def fixtures_dir():
    """Path to the test fixtures directory."""
    return FIXTURES_DIR


@pytest.fixture
def minimal_workbook():
    """Path to the minimal test workbook, or skip if not found."""
    path = os.path.join(FIXTURES_DIR, "minimal.xlsm")
    if not os.path.exists(path):
        pytest.skip("minimal.xlsm test fixture not found")
    return path


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
