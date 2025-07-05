'======================================================================
'  modMegaSuite - GG MegaSuite v5.0 core
'======================================================================
Option Private Module
Option Explicit
Option Compare Text

'==============================  CONFIG  ==============================================
Public Const CONFIG_SHEET       As String = "Config"
Public Const CORPORATE_SHEET    As String = "Corporate"
Public Const WHOLECO_SHEET      As String = "WholeCo"
Public Const FX_SHEET           As String = "FX_Table"
Public Const FX_RANGE_NAME      As String = "FX_Matrix"
Public Const EV_CELL            As String = "EV_Output"
Public Const SOTP_PREFIX        As String = "SOTP_"
Public Const ALLOC_KEY_CELL     As String = "B21"
Private Const HEADER_ROW        As Long = 7
Private Const TAB_A             As String = "A"
Private Const TAB_B             As String = "B"

'==========================  ERROR CODE MAP  ==========================================
Public Const errSOTP_EVMismatch        As Long = vbObjectError + 1001
Public Const errSOTP_UnallocatedItems  As Long = vbObjectError + 1002
Public Const errSOTP_InvalidFX         As Long = vbObjectError + 1003
Public Const errSOTP_ExcessiveDilution As Long = vbObjectError + 1004

Private Enum eErr
    errMissingTemplate = 900
    errMissingMacroList = 901
    errEmptyMacroList = 902
    errNullWorksheet = 903
    errSheetNotFound = 904
    errSheetProtected = 905
    errBadTemplate = 906
    errBadLoaderStructure = 907
    errInsufficientYears = 908
    errInvalidYear = 909
End Enum

'===========================  UTILITY HELPERS  ========================================
Private Function Cfg(key As String, Optional fallback As Variant = "") As Variant
    On Error Resume Next
    Dim rng As Range
    Set rng = Worksheets(CONFIG_SHEET).Columns(1).Find(What:=key, LookIn:=xlValues, _
                LookAt:=xlWhole, MatchCase:=False)
    If Not rng Is Nothing Then
        Cfg = rng.Offset(0, 1).Value
    Else
        Cfg = fallback
    End If
End Function

Private Function Nz(val As Variant, Optional fallback As Variant = 0) As Variant
    If IsError(val) Or IsEmpty(val) Or val = "" Then
        Nz = fallback
    Else
        Nz = val
    End If
End Function

Private Function IsSOTPSegment(ws As Worksheet) As Boolean
    IsSOTPSegment = (Left(ws.Name, Len(SOTP_PREFIX)) = SOTP_PREFIX)
End Function

'=====================  SEGMENT IDENTIFICATION & TAGGING  =============================
Public Sub TagSOTPSegments()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name Like "[A-Z]{1,3}_[0-9][0-9][0-9][0-9]" Then
            On Error Resume Next
            ws.Names.Add Name:="SOTP_Flag", RefersTo:="=TRUE"
            ws.Range("Alloc_Key").Formula = "=" & CONFIG_SHEET & "!" & ALLOC_KEY_CELL
            On Error GoTo 0
        End If
    Next ws
End Sub

'=======================  CORPORATE ITEM ALLOCATION  ==================================
Private Function SumVisibleSheets(fieldName As String) As Double
    Dim ws As Worksheet, total As Double
    For Each ws In ThisWorkbook.Worksheets
        If IsSOTPSegment(ws) And ws.Visible Then total = total + Nz(ws.Range(fieldName).Value)
    Next ws
    SumVisibleSheets = total
End Function

Private Function AllocateCorporateItem(seg As Worksheet, corpField As String, segField As String) As Double
    Dim totalKey As Double: totalKey = SumVisibleSheets(segField)
    Dim segVal   As Double: segVal = Nz(seg.Range(segField).Value)
    Dim corpItem As Double: corpItem = Nz(Worksheets(CORPORATE_SHEET).Range(corpField).Value)
    If totalKey = 0 Then
        AllocateCorporateItem = 0
    Else
        AllocateCorporateItem = Application.Max(0, corpItem * segVal / totalKey)
    End If
End Function

Public Sub AllocateCorporateItems()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsSOTPSegment(ws) Then
            ws.Range("SBC_Alloc").Value  = AllocateCorporateItem(ws, "Total_SBC",      Cfg("Alloc_Key", "Revenue"))
            ws.Range("NWC_Alloc").Value  = AllocateCorporateItem(ws, "Total_NWC",      "Avg_Receivables")
            ws.Range("CapEx_Maint").Value = ws.Range("PPE").Value * Nz(Cfg("Depr_Rate", 0.03))
            ws.Range("CapEx_Growth").Value = AllocateCorporateItem(ws, "Growth_CapEx", "Revenue_Growth")
        End If
    Next ws
End Sub

'===============================  FX CONVERSION  ======================================
Public Function ConvertToReportingCCY(amount As Double, sourceCCY As String) As Double
    Dim fxTable As Range
    Set fxTable = Worksheets(FX_SHEET).Range(FX_RANGE_NAME)
    Dim rate As Double: rate = Nz(Application.VLookup(sourceCCY, fxTable, 2, False))
    If rate = 0 Then Err.Raise errSOTP_InvalidFX, , "Missing FX rate for: " & sourceCCY
    ConvertToReportingCCY = amount * rate
End Function

'=======================  DILUTION-AWARE SHARE BRIDGE  ================================
Public Sub CalculateDilutedShares()
    With Worksheets(WHOLECO_SHEET)
        Dim bs  As Double: bs  = Nz(.Range("Basic_Shares").Value)
        Dim opt As Double: opt = Nz(.Range("Outstanding_Options").Value)
        Dim k   As Double: k   = Nz(.Range("Avg_Strike_Price").Value)
        Dim px  As Double: px  = Nz(FetchFromBloomberg("LAST_PRICE"))

        Dim proceeds As Double: proceeds    = opt * k
        Dim repurch  As Double: repurch     = proceeds / px
        Dim netDil   As Double: netDil      = opt - repurch

        .Range("Diluted_Shares").Value = bs + netDil
        .Range("Dilution_Pct").Value   = netDil / bs
    End With
End Sub

'===========================  SOTP VALUATION BUILD  ===================================
Public Sub BuildSOTPValuation()
    Dim seg As Worksheet
    Dim flows As Variant, rates() As Double, rate As Double, i As Long
    For Each seg In GetBusinessSegments()
        flows = Application.Transpose(seg.Range("FCF_Stream").Value)
        rate = Nz(FMP_WACC(seg.Range("Ticker").Value) / 100, Cfg("WACC", 0.1))
        ReDim rates(LBound(flows) To UBound(flows))
        For i = LBound(flows) To UBound(flows)
            rates(i) = rate
        Next i
        seg.Range(EV_CELL).Value = DCF_NPV(flows, rates)
    Next seg
End Sub

'============================  SOTP ROLLUP  ==============================
Public Sub BuildRollups()
    Dim seg As Worksheet, ws As Worksheet
    Dim rowIdx As Long
    Set ws = GetOrCreateSheet("SOTP_Rollup", True)
    ws.Range("A1:B1").Value = Array("Segment", "EV")
    rowIdx = 2
    For Each seg In GetBusinessSegments()
        ws.Cells(rowIdx, 1).Value = seg.Name
        ws.Cells(rowIdx, 2).Value = Nz(seg.Range("Segment_EV").Value)
        rowIdx = rowIdx + 1
    Next seg
    ws.Cells(rowIdx, 1).Value = "Total"
    ws.Cells(rowIdx, 2).Formula = "=SUM(B2:B" & rowIdx - 1 & ")"
    ws.Columns.AutoFit
End Sub

'==========================  RECONCILIATION & AUDIT  ==================================
Public Sub VerifySOTPReconciliation()
    Const TOLERANCE As Double = 0.0001
    Dim segEVSum As Double: segEVSum = SumVisibleSheets("Segment_EV")
    Dim wholecoEV As Double: wholecoEV = Worksheets(WHOLECO_SHEET).Range("EntValue").Value

    If Abs(segEVSum - wholecoEV) > TOLERANCE Then
        LogError "EV MISMATCH: Δ" & Format$(Abs(segEVSum - wholecoEV), "Currency")
        Worksheets(WHOLECO_SHEET).Range("EntValue").Interior.Color = vbRed
        Err.Raise errSOTP_EVMismatch, , "SOTP EV mismatch vs WholeCo"
    End If
End Sub

Public Sub VerifyNWCTieOut()
    Dim segNWC As Double: segNWC = SumVisibleSheets("NWC_Alloc")
    Dim corpNWC As Double: corpNWC = Nz(Worksheets(CORPORATE_SHEET).Range("Total_NWC").Value)
    Dim wholecoNWC As Double: wholecoNWC = Worksheets(WHOLECO_SHEET).Range("Total_NWC").Value
    If Abs(segNWC + corpNWC - wholecoNWC) > 0.0001 Then
        LogWrite "NWC mismatch", "ERROR"
        Err.Raise errSOTP_UnallocatedItems, , "NWC tie-out failed"
    End If
End Sub
'=========================  PRIMARY SUITE INVOKER  ====================================
Public Sub RunMegaSuite()
    Dim t0 As Double
    t0 = HiResTimerStart()
    LogWrite "MegaSuite start"
    If Cfg("SOTP_ENABLED", False) Then
        TagSOTPSegments
        AllocateCorporateItems
        BuildSOTPValuation
        BuildRollups
        CalculateDilutedShares
        VerifySOTPReconciliation
        VerifyNWCTieOut
        RunMLDiag ThisWorkbook.Path & "\sotp.csv", ThisWorkbook.Path & "\ml_out.json"
    End If
    LogWrite "MegaSuite runtime: " & Format$(HiResTimerEnd(t0), "0.00") & "s"
    DumpLogJSON
End Sub
Private Sub LogError(msg As String)
    LogWrite msg, "ERROR"
End Sub

Private Sub LogWarning(msg As String)
    LogWrite msg, "WARN"
End Sub

Private Sub FlagVariance(label As String, delta As Double)
    MsgBox "Variance flagged in " & label & ": Δ" & Format$(delta, "0.00"), vbExclamation
End Sub

'===========================  EXTERNAL DATA STUBS  ====================================
Private Function FetchFromBloomberg(fieldCode As String) As Double
    FetchFromBloomberg = Nz(Worksheets("Bloomberg_PriceCache").Range("Last_Price").Value, 0)
End Function

'==========================  HIGH-RES TIMING  ============================
#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lp As LongLong) As Long
        Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lp As LongLong) As Long
    #Else
        Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lp As Currency) As Long
        Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lp As Currency) As Long
    #End If
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lp As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lp As Currency) As Long
#End If

Private Function HiResTimerStart() As Double
#If Win64 Then
    Dim t As LongLong, f As LongLong
    QueryPerformanceCounter t
    QueryPerformanceFrequency f
    HiResTimerStart = t / f
#Else
    Dim t As Currency, f As Currency
    QueryPerformanceCounter t
    QueryPerformanceFrequency f
    HiResTimerStart = t / f
#End If
End Function

Private Function HiResTimerEnd(st As Double) As Double
#If Win64 Then
    Dim t As LongLong, f As LongLong
    QueryPerformanceCounter t
    QueryPerformanceFrequency f
    HiResTimerEnd = t / f - st
#Else
    Dim t As Currency, f As Currency
    QueryPerformanceCounter t
    QueryPerformanceFrequency f
    HiResTimerEnd = t / f - st
#End If
End Function

'===========================  HELPER FUNCTIONS  =======================================
Public Function GetBusinessSegments() As Collection
    Dim seg As Worksheet, segs As New Collection
    For Each seg In ThisWorkbook.Worksheets
        If Left(seg.Name, Len(SOTP_PREFIX)) = SOTP_PREFIX Then segs.Add seg, seg.Name
    Next seg
    Set GetBusinessSegments = segs
End Function

Public Function SheetExists(nm As String) As Boolean
    On Error Resume Next
    SheetExists = Not Worksheets(nm) Is Nothing
End Function

Public Sub GuardSheetExists(nm As String)
    If Not SheetExists(nm) Then Err.Raise vbObjectError + errSheetNotFound, , "Sheet '" & nm & "' missing"
End Sub

Public Function GetOrCreateSheet(nm As String, Optional clearIt As Boolean = False) As Worksheet
    If SheetExists(nm) Then
        Set GetOrCreateSheet = Worksheets(nm)
        If clearIt Then GetOrCreateSheet.Cells.Clear
    Else
        Set GetOrCreateSheet = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        GetOrCreateSheet.Name = nm
    End If
End Function

'----------------------- Utility: column/years helpers -------------------------
Private Function ColLetter(n As Long) As String
    ColLetter = Split(Cells(1, n).Address(True, False), "$")(0)
End Function

Private Function SafeYearCol(ws As Worksheet, yr As Long) As Long
    Dim f As Range
    Set f = ws.Rows(HEADER_ROW).Find(yr, lookat:=xlWhole)
    If Not f Is Nothing Then
        SafeYearCol = f.Column
    Else
        SafeYearCol = ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Column + 1
        ws.Cells(HEADER_ROW, SafeYearCol).Value = yr
    End If
End Function

Private Function GetYears(ws As Worksheet) As Long()
    Dim col As Long, yrs()
    col = 3
    ReDim yrs(0)
    Do While IsNumeric(ws.Cells(2, col).Value)
        yrs(UBound(yrs)) = CLng(ws.Cells(2, col).Value)
        ReDim Preserve yrs(UBound(yrs) + 1)
        col = col + 1
    Loop
    ReDim Preserve yrs(UBound(yrs) - 1)
    GetYears = yrs
End Function

Private Function GetHeaderYears(ws As Worksheet) As Variant
    GetHeaderYears = Array(ws.Cells(HEADER_ROW, 3).Value, _
                           ws.Cells(HEADER_ROW, ws.Columns.Count).End(xlToLeft).Value)
End Function

Private Function Build3DSum(sheets As Collection, tgtCol As Long, tgtRow As Long) As String
    If sheets.Count = 0 Then
        Build3DSum = "0"
    Else
        Build3DSum = "'" & sheets(1).Name & ":" & sheets(sheets.Count).Name & "'!" & _
                     ColLetter(tgtCol) & tgtRow
    End If
End Function

Private Function ListSheetsLike(sfx As String) As Collection
    Dim col As New Collection, w As Worksheet
    For Each w In Worksheets
        If Right$(w.Name, Len(sfx)) = sfx Then col.Add w
    Next w
    Set ListSheetsLike = col
End Function

Private Function CorrBSheet(a As Worksheet) As Worksheet
    Dim nm As String
    nm = Replace(a.Name, " (" & TAB_A & ")", " (" & TAB_B & ")")
    If SheetExists(nm) Then Set CorrBSheet = Worksheets(nm)
End Function

