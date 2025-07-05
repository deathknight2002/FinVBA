'======================================================================
'  modFX  – Central FX store & helpers        GG MegaSuite v5.x add-in
'======================================================================
Option Private Module
Option Explicit
Option Compare Text

'--------------------  CONFIG  --------------------
Private Const FX_SHEET        As String = "_FX"
Private Const FX_LAST_UPDATE  As String = "Last_Update"

'--------------------  PUBLIC API  ---------------
' Put rates in _FX:  Col A=CCY (ISO-3), Col B=USD spot
Public Function GetFX(ccy As String, _
                      Optional silent As Boolean = False) As Double
    Dim ws As Worksheet: Set ws = Worksheets(FX_SHEET)
    Dim f As Range: Set f = ws.Columns(1).Find(ccy, LookAt:=xlWhole)
    If f Is Nothing Then
        If silent Then Exit Function
        Err.Raise vbObjectError + 1201, , "FX rate for '" & ccy & "' not found"
    End If
    GetFX = ws.Cells(f.Row, 2).Value
End Function

Public Function ConvertCCY(amt As Double, _
                           fromCcy As String, _
                           toCcy As String) As Double
    If fromCcy = toCcy Then
        ConvertCCY = amt
    Else
        Dim usd As Double
        usd = amt / GetFX(fromCcy, True)  'to USD
        ConvertCCY = usd * GetFX(toCcy, True)
    End If
End Function

'------------------  ADMIN MACROS  ---------------
' Paste broker feed CSV into _FX!A:B then call SyncFX
Public Sub SyncFX()
    GuardSheetExists FX_SHEET
    Dim ws As Worksheet: Set ws = Worksheets(FX_SHEET)
    ws.Range("E1").Value = Now
End Sub
