'======================================================================
'  modSpeedyTabs – bucket tab builder and rollup summarizer (2025)
'======================================================================
Option Private Module
Option Explicit

Private Const BUCKET_COL As Long = 5      'Column E
Private Const HEADER_ROW As Long = 1
Private Const BUCKET_YEAR As Long = 2025
Private BucketSheets As Variant

'---------------------------
' Build or refresh A/U/D tabs from active sheet
'---------------------------
Public Sub GGSpeedyTabs()
    Dim srcWS As Worksheet
    Dim lastRow As Long
    Dim r As Long, bucket As String
    Dim ws As Worksheet
    Dim destWS As Worksheet
    Dim destRow As Long
    BucketSheets = Array("A_" & BUCKET_YEAR, "U_" & BUCKET_YEAR, "D_" & BUCKET_YEAR)

    Set srcWS = ActiveSheet
    lastRow = srcWS.Cells(srcWS.Rows.Count, BUCKET_COL).End(xlUp).Row

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim b As Variant
    For Each b In BucketSheets
        KillSheet b
        Set ws = Worksheets.Add(After:=srcWS)
        ws.Name = b
        srcWS.Rows(HEADER_ROW).Copy ws.Rows(HEADER_ROW)
    Next b

    For r = HEADER_ROW + 1 To lastRow
        bucket = UCase$(Trim$(srcWS.Cells(r, BUCKET_COL).Value))
        If bucket <> "" Then
            For Each b In BucketSheets
                If bucket & "_" & BUCKET_YEAR = b Then
                    Set destWS = Worksheets(b)
                    destRow = destWS.Cells(destWS.Rows.Count, 1).End(xlUp).Row + 1
                    srcWS.Rows(r).Copy destWS.Rows(destRow)
                    destWS.Cells(destRow, BUCKET_COL).ClearContents
                End If
            Next b
        End If
    Next r

    For Each b In BucketSheets
        Worksheets(b).Columns.AutoFit
    Next b

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "GGSpeedyTabs completed", vbInformation
End Sub

'---------------------------
' Summarize each bucket into Rollup_YYYY
'---------------------------
Public Sub BuildSpeedyRollup()
    BucketSheets = Array("A_" & BUCKET_YEAR, "U_" & BUCKET_YEAR, "D_" & BUCKET_YEAR)
    Dim roll As Worksheet, ws As Worksheet
    Dim col As Long, lastRow As Long
    Dim destRow As Long, b As Variant

    KillSheet "Rollup_" & BUCKET_YEAR
    Set roll = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    roll.Name = "Rollup_" & BUCKET_YEAR

    Worksheets(BucketSheets(0)).Rows(HEADER_ROW).Copy roll.Rows(HEADER_ROW)
    destRow = HEADER_ROW + 1

    For Each b In BucketSheets
        Set ws = Worksheets(b)
        roll.Cells(destRow, 1).Value = b
        For col = 1 To ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column
            If col <> BUCKET_COL Then
                lastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
                roll.Cells(destRow, col).Value = Application.Sum(ws.Range(ws.Cells(HEADER_ROW + 1, col), ws.Cells(lastRow, col)))
            End If
        Next col
        destRow = destRow + 1
    Next b

    roll.Columns.AutoFit
End Sub

'---------------------------
' Delete sheet if it exists
'---------------------------
Private Sub KillSheet(nm As String)
    On Error Resume Next
    Worksheets(nm).Delete
    On Error GoTo 0
End Sub
