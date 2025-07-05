'======================================================================
'  modComps – Comparable-company & transaction multiples hub
'======================================================================
Option Private Module
Option Explicit
Option Compare Text

' COMP_SHEET layout:  A=Tkr | B=CCY | C=LTM Rev | D=LTM EBITDA | E=EV
Private Const COMP_SHEET As String = "_Comps"

Public Function GetPeerStats(ByVal metric As String) As Double()
    Dim ws As Worksheet: Set ws = Worksheets(COMP_SHEET)
    Dim rng As Range, vals() As Double, i As Long
    Dim col As Long: col = Switch(metric = "EV/EBITDA", 5, _
                                  metric = "EV/Revenue", 5, _
                                  True, 0)
    If col = 0 Then Err.Raise vbObjectError + 1301, , "Metric not supported"
    Set rng = ws.Range(ws.Cells(2, col), ws.Cells(ws.Rows.Count, col).End(xlUp))
    ReDim vals(1 To rng.Rows.Count)
    For i = 1 To rng.Rows.Count
        vals(i) = rng.Cells(i).Value / rng.Offset(0, IIf(metric = "EV/EBITDA", -1, -2)).Cells(i).Value
    Next i
    GetPeerStats = vals
End Function

Public Function PeerPercentile(p() As Double, pct As Double) As Double
    Dim tmp(): tmp = p
    Dim i As Long, j As Long, t As Double
    For i = LBound(tmp) To UBound(tmp) - 1
        For j = i + 1 To UBound(tmp)
            If tmp(j) < tmp(i) Then t = tmp(i): tmp(i) = tmp(j): tmp(j) = t
        Next j
    Next i
    PeerPercentile = tmp(Application.WorksheetFunction.RoundUp(UBound(tmp) * pct, 0))
End Function
