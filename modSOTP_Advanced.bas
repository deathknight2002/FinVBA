'======================================================================
'  modSOTP_Advanced – multi-method, FX-aware SOTP engine
'======================================================================
Option Private Module
Option Explicit
Option Compare Text

' Tag each segment sheet with named range "SEG_CCY" (ISO-3)
' and populate valuation result cells:
'   Range("DCF_EV"), Range("Trading_EV"), Range("Transaction_EV")

'------------------  CONSTANTS -------------------
Private Const DASH_SHEET   As String = "SOTP_Dashboard"
Private Const SEG_PREFIX   As String = "SOTP_"
Private Enum eMethod: mDCF = 1: mTrading: mDeal: End Enum

'------------------  ENTRY POINT -----------------
Public Sub BuildAdvancedSOTP()
    Dim segs As Collection: Set segs = GetBusinessSegments()
    Dim ws As Worksheet: Set ws = GetOrCreateSheet(DASH_SHEET, True)
    
    ws.Range("A1:F1").Value = Array("Segment", "Method", "Local EV", _
                                    "CCY", "USD EV", "EV Weight")
    Dim r As Long: r = 2
    Dim evTotal As Double, s As Worksheet, usdEV As Double
    Dim m As eMethod
    
    For Each s In segs
        For m = mDCF To mDeal
            Dim locEV As Double: locEV = ValByMethod(s, m)
            If locEV = 0 Then GoTo nxtM
            usdEV = ConvertCCY(locEV, s.Range("SEG_CCY").Value, "USD")
            ws.Cells(r, 1).Resize(1, 5).Value = _
                Array(Replace(s.Name, SEG_PREFIX, ""), _
                      MethodName(m), locEV, s.Range("SEG_CCY").Value, usdEV)
            evTotal = evTotal + usdEV
            r = r + 1
nxtM:   Next m
    Next s
    
    ws.Range("F2:F" & r - 1).FormulaR1C1 = "=RC[-1]/" & evTotal
    ws.Columns.AutoFit
End Sub

Private Function ValByMethod(seg As Worksheet, m As eMethod) As Double
    Select Case m
        Case mDCF:        ValByMethod = Nz(seg.Range("DCF_EV").Value)
        Case mTrading:    ValByMethod = Nz(seg.Range("Trading_EV").Value)
        Case mDeal:       ValByMethod = Nz(seg.Range("Transaction_EV").Value)
    End Select
End Function

Private Function MethodName(m As eMethod) As String
    MethodName = Choose(m, "", "DCF", "Trading Comp", "Deal Comp")
End Function

'------------------  PEER FLAGGING ----------------
Public Sub FlagOutliers()
    Dim segs As Collection: Set segs = GetBusinessSegments()
    Dim p() As Double: p = GetPeerStats("EV/EBITDA")
    Dim hi As Double: hi = PeerPercentile(p, 0.75)
    
    Dim s As Worksheet
    For Each s In segs
        Dim ev As Double, ebitda As Double
        ev = s.Range("Trading_EV").Value
        ebitda = s.Range("EBITDA").Value
        If ebitda <= 0 Then GoTo nxt
        If (ev / ebitda) > hi Then
            s.Range("Trading_EV").Interior.ColorIndex = 22
        End If
nxt: Next s
End Sub

'------------------  UTIL ------------------------
Private Function Nz(v) As Double: If IsNumeric(v) Then Nz = v
End Function
