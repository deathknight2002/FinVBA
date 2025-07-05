'======================================================================
'  modSectorTemplates – sector-specific template cloning
'======================================================================
Option Private Module
Option Explicit
Option Compare Text

' Template names must follow pattern  TPL_<Sector>
' Asset sheet name requested as   <Ticker>_<YYYY>_<Sector>

Private Const TPL_PREFIX     As String = "TPL_"

Public Sub AutoAssetDiscovery()
    Dim rgx As Object: Set rgx = CreateObject("VBScript.RegExp")
    rgx.Pattern = "^[A-Z]{1,3}_\d{4}_(\w+)$"
    Dim w As Worksheet, m, tpl

    For Each w In Worksheets
        If rgx.Test(w.Name) Then
            Set m = rgx.Execute(w.Name)
            tpl = TPL_PREFIX & UCase$(m(0).SubMatches(0))
            If SheetExists(tpl) Then
                CloneIfEmpty w, tpl
            End If
        End If
    Next w
End Sub

Private Sub CloneIfEmpty(tgt As Worksheet, tpl As String)
    If Application.WorksheetFunction.CountA(tgt.Cells) > 0 Then Exit Sub
    Worksheets(tpl).Copy Before:=tgt
    tgt.Delete
End Sub
