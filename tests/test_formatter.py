"""
Tests for xlvbatools.vba.formatter -- VBA code formatter.
"""

import pytest
from xlvbatools.vba.formatter import format_vba


@pytest.mark.unit
class TestVBAFormatter:
    """Test VBA code formatting."""

    def test_basic_indentation(self):
        source = """Option Explicit
Public Sub Test()
Dim x As Double
x = 42
Debug.Print x
End Sub
"""
        result = format_vba(source)
        lines = result.strip().split("\n")
        assert lines[0] == "Option Explicit"
        assert lines[2].startswith("    Dim")
        assert lines[3].startswith("    x = 42")
        assert lines[4].startswith("    Debug.Print")
        assert lines[5] == "End Sub"

    def test_if_then_block(self):
        source = """Public Sub Test()
If x > 0 Then
Debug.Print "positive"
End If
End Sub
"""
        result = format_vba(source)
        assert '        Debug.Print "positive"' in result

    def test_for_loop(self):
        source = """Public Sub Test()
Dim i As Long
For i = 1 To 10
Debug.Print i
Next i
End Sub
"""
        result = format_vba(source)
        assert "        Debug.Print i" in result

    def test_select_case(self):
        source = """Public Sub Test()
Select Case x
Case 1
Debug.Print "one"
Case Else
Debug.Print "other"
End Select
End Sub
"""
        result = format_vba(source)
        assert "End Select" in result

    def test_collapse_blank_lines(self):
        source = """Option Explicit


Public Sub Test()



Dim x As Long
End Sub
"""
        result = format_vba(source)
        # No more than one consecutive blank line
        assert "\n\n\n" not in result

    def test_preserves_vbe_headers(self):
        source = """Attribute VB_Name = "modTest"
Option Explicit

Public Sub Test()
End Sub
"""
        result = format_vba(source)
        assert 'Attribute VB_Name = "modTest"' in result

    def test_format_idempotent(self):
        source = """Option Explicit

Public Sub Test()
    Dim x As Double
    x = 42
    Debug.Print x
End Sub
"""
        first = format_vba(source)
        second = format_vba(first)
        assert first == second, "Formatter should be idempotent"


@pytest.mark.unit
class TestFormatFile:
    """Test file formatting."""

    def test_format_file_dry_run(self, sample_bas_file):
        from xlvbatools.vba.formatter import format_file
        result = format_file(str(sample_bas_file), dry_run=True)
        assert isinstance(result, dict)
        assert "changed" in result

    def test_format_directory(self, temp_vba_source, sample_bas_file):
        from xlvbatools.vba.formatter import format_directory
        results = format_directory(str(temp_vba_source), dry_run=True)
        assert len(results) >= 1
        assert all("file" in r for r in results)
