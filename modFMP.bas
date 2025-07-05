'======================================================================
'  modFMP - Financial Modeling Prep API helpers
'======================================================================
Option Private Module
Option Explicit
Option Compare Text

Private Const FMP_API_KEY As String = "demo"
Private Const FMP_URL As String = "https://financialmodelingprep.com/api/v3/"

Private Function HttpGet(url As String) As String
    Dim req As Object
    Set req = CreateObject("WinHttp.WinHttpRequest.5.1")
    req.Open "GET", url, False
    req.Send
    If req.Status <> 200 Then Err.Raise vbObjectError + 1501, , "HTTP " & req.Status
    HttpGet = req.ResponseText
End Function

Public Function FMP_Quote(symbol As String) As Double
    Dim j As String
    j = HttpGet(FMP_URL & "quote/" & symbol & "?apikey=" & FMP_API_KEY)
    FMP_Quote = ExtractNumber(j, "price")
End Function

Public Function FMP_WACC(symbol As String) As Double
    Dim j As String
    j = HttpGet(FMP_URL & "wacc/" & symbol & "?apikey=" & FMP_API_KEY)
    FMP_WACC = ExtractNumber(j, "wacc")
End Function

Private Function ExtractNumber(json As String, key As String) As Double
    Dim p As Long, e As Long
    p = InStr(json, '"' & key & '"' & ":")
    If p = 0 Then Exit Function
    p = p + Len(key) + 3
    e = p
    Do While Mid$(json, e, 1) Like "[0-9.Ee+-]"
        e = e + 1
    Loop
    ExtractNumber = CDbl(Mid$(json, p, e - p))
End Function
