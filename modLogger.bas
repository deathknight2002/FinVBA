'======================================================================
'  modLogger - JSON logging utilities
'======================================================================
Option Private Module
Option Explicit
Option Compare Text

Private Const LOG_SHEET As String = "_Log"

Public Sub LogWrite(msg As String, Optional level As String = "INFO")
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(LOG_SHEET, False)
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = level
    ws.Cells(r, 3).Value = msg
End Sub

Public Sub DumpLogJSON()
    Dim ws As Worksheet: Set ws = Worksheets(LOG_SHEET)
    Dim arr As Variant: arr = ws.Range("A2", ws.Cells(ws.Rows.Count, 3).End(xlUp)).Value
    If IsEmpty(arr) Then Exit Sub
    Dim json As String: json = "["
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        json = json & "{""ts"":""" & arr(i, 1) & """,""lvl"":""" & arr(i, 2) & """,""msg"":""" & Replace(arr(i, 3), """", """""") & """},"
    Next i
    json = Left$(json, Len(json) - 1) & "]"
    Dim f As Object: Set f = CreateObject("Scripting.FileSystemObject").CreateTextFile(ThisWorkbook.Path & "\log_" & Format(Now, "yyyymmdd_hhnnss") & ".json", True, True)
    f.Write json
    f.Close
End Sub
