'======================================================================
'  modMLDiag - bridge to mlfinlab diagnostics
'======================================================================
Option Private Module
Option Explicit

Private Const PYTHON_EXE As String = "python"
Private Const SCRIPT_PATH As String = "C:\\ml_scripts\\ml_diagnostics.py"

Public Function RunMLDiag(inputCSV As String, outputJSON As String) As Boolean
    Dim cmd As String
    cmd = PYTHON_EXE & " " & SCRIPT_PATH & " """ & inputCSV & """ """ & outputJSON & """"
    Shell cmd, vbHide
    RunMLDiag = (Dir(outputJSON) <> "")
End Function

Public Function ReadMLDiag(path As String) As String
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function
    Set ts = fso.OpenTextFile(path)
    ReadMLDiag = ts.ReadAll
    ts.Close
End Function
