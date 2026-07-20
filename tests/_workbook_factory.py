"""Child-process-only builder for disposable live Excel test workbooks."""

from __future__ import annotations

import argparse
import gc
from pathlib import Path
import sys
import time
import traceback


RUNTIME_MODULE = '''Option Explicit
Public Sub CompleteNormally()
    ThisWorkbook.Worksheets(1).Range("A1").Value = 42
End Sub
Public Sub WorkflowRetrieve()
    ThisWorkbook.Worksheets(1).Range("A1").Value = 10
End Sub
Public Sub WorkflowCalculate()
    With ThisWorkbook.Worksheets(1)
        .Range("A3").Value = .Range("A1").Value + .Range("A2").Value
    End With
End Sub
Public Sub WorkflowFail()
    ThisWorkbook.Worksheets(1).Range("A4").Value = 99
    Err.Raise vbObjectError + 102, "WorkflowFail", _
        "Workflow failure."
End Sub
Public Sub WorkflowMustNotRun()
    ThisWorkbook.Worksheets(1).Range("A5").Value = 1
End Sub
Public Sub LeaveScreenUpdatingOff()
    With ThisWorkbook.Worksheets(1)
        .Range("A1").Value = "Visible screenshot content"
        .Range("A2").Value = 123.45
    End With
    Application.ScreenUpdating = False
End Sub
Public Sub VerifyScreenUpdatingStillOff()
    If Application.ScreenUpdating Then
        Err.Raise vbObjectError + 103, "VerifyScreenUpdatingStillOff", _
            "Screenshot capture did not restore ScreenUpdating."
    End If
End Sub
Public Sub VerifyNamedRange()
    If ThisWorkbook.Names("TestInput").RefersToRange.Value <> 42 Then
        Err.Raise vbObjectError + 101, "VerifyNamedRange", _
            "TestInput was not set to 42."
    End If
End Sub
Public Sub ShowMessage()
    MsgBox "Watchdog test message", vbInformation
End Sub
Public Sub ShowFilePicker()
    Dim selected As Variant
    selected = Application.GetOpenFilename()
End Sub
Public Sub LoopForever()
    Do
    Loop
End Sub
Public Sub RaiseMultilineError()
    Err.Raise vbObjectError + 100, "TestMacro", _
        "Diagnostic line one." & vbCrLf & "Diagnostic line two."
End Sub
'''


def _add_module(workbook, name: str, source: str):
    component = workbook.VBProject.VBComponents.Add(1)
    component.Name = name
    code_module = component.CodeModule
    code_module.AddFromString(source.replace("\n", "\r\n"))
    return component, code_module


def _build(kind: str, output: Path, marker: Path | None) -> None:
    import pythoncom
    import win32com.client
    import win32process

    from xlvbatools.core.process import is_process_running, kill_process_by_pid

    excel = None
    workbook = None
    component = None
    code_module = None
    workbook_module = None
    sheet = None
    cell = None
    interior = None
    excel_pid = None
    com_initialized = False
    pythoncom.CoInitialize()
    com_initialized = True
    try:
        output.parent.mkdir(parents=True, exist_ok=True)
        excel = win32com.client.DispatchEx("Excel.Application")
        _, excel_pid = win32process.GetWindowThreadProcessId(excel.Hwnd)
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        workbook = excel.Workbooks.Add()

        if kind in {"runtime", "compile_error", "duplicate_declaration"}:
            component, code_module = _add_module(
                workbook, "modReliabilityTest", RUNTIME_MODULE,
            )
            code_module = None
            component = None
            sheet = workbook.Worksheets(1)
            sheet.Range("C1").Value = 0
            workbook.Names.Add(Name="TestInput", RefersTo=f"={sheet.Name}!$C$1")
            cell = sheet.Range("B1")
            cell.Formula = "=21*2"
            cell.NumberFormat = "0.00"
            interior = cell.Interior
            interior.Color = 65535
            interior = None
            cell = None
            sheet = None

        if kind == "runtime":
            pass
        elif kind == "compile_error":
            component, code_module = _add_module(
                workbook,
                "modCompileFailure",
                "Option Explicit\n"
                "Public Sub CompileFailure()\n"
                "    undeclaredCompileValue = 1\n"
                "End Sub\n",
            )
        elif kind == "duplicate_declaration":
            component, code_module = _add_module(
                workbook,
                "modDuplicateDeclaration",
                "Option Explicit\n"
                "Public Sub DuplicateDeclaration(ByVal FileCount As Long)\n"
                "    Dim FileCount As Long\n"
                "End Sub\n",
            )
        elif kind == "startup_event":
            if marker is None:
                raise ValueError("startup_event requires --marker")
            workbook_module = workbook.VBProject.VBComponents.Item("ThisWorkbook")
            code_module = workbook_module.CodeModule
            marker_literal = str(marker).replace('"', '""')
            code_module.AddFromString((
                "Option Explicit\n"
                "Private Sub Workbook_Open()\n"
                "    Dim channel As Integer\n"
                "    channel = FreeFile\n"
                f'    Open "{marker_literal}" For Output As #channel\n'
                '    Print #channel, "Workbook_Open executed"\n'
                "    Close #channel\n"
                "End Sub\n"
            ).replace("\n", "\r\n"))
            code_module = None
            workbook_module = None
            component, code_module = _add_module(
                workbook,
                "modStartupSafety",
                "Option Explicit\nPublic Sub SafeProcedure()\nEnd Sub\n",
            )
        elif kind != "minimal":
            raise ValueError(f"Unknown workbook fixture kind: {kind}")

        code_module = None
        component = None
        gc.collect()
        workbook.SaveAs(str(output), FileFormat=52)
        workbook.Close(False)
        workbook = None
    finally:
        code_module = None
        del component
        workbook_module = None
        interior = None
        cell = None
        sheet = None
        if workbook is not None:
            try:
                workbook.Close(False)
            except Exception:
                pass
        workbook = None
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        excel = None
        gc.collect()
        gc.collect()
        if com_initialized:
            pythoncom.CoUninitialize()
        if excel_pid is not None:
            deadline = time.time() + 20
            while time.time() < deadline and is_process_running(excel_pid):
                time.sleep(0.1)
            if is_process_running(excel_pid):
                kill_process_by_pid(excel_pid)
                deadline = time.time() + 10
                while time.time() < deadline and is_process_running(excel_pid):
                    time.sleep(0.1)
            if is_process_running(excel_pid):
                raise RuntimeError(f"Fixture Excel PID {excel_pid} did not exit")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--kind",
        required=True,
        choices=(
            "minimal",
            "runtime",
            "compile_error",
            "duplicate_declaration",
            "startup_event",
        ),
    )
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--marker", type=Path)
    args = parser.parse_args()
    try:
        _build(args.kind, args.output.resolve(), args.marker)
    except Exception:
        traceback.print_exc()
        return 1
    print(f"WORKBOOK_FIXTURE_READY={args.output.resolve()}", flush=True)
    return 0


if __name__ == "__main__":
    sys.exit(main())
