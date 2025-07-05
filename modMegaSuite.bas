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
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If IsSOTPSegment(ws) Then
            ws.Range(EV_CELL).Formula = "=XNPV(" & Cfg("WACC", 0.10) & ", FCF_Stream, Date_Stream)"
        End If
    Next ws
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

'=========================  PRIMARY SUITE INVOKER  ====================================
Public Sub RunMegaSuite()
    If Cfg("SOTP_ENABLED", False) Then
        TagSOTPSegments
        AllocateCorporateItems
        BuildSOTPValuation
        CalculateDilutedShares
        VerifySOTPReconciliation
    End If
End Sub

'=============================  LOGGING HELPERS  ======================================
Private Sub LogError(msg As String)
    Debug.Print "[ERROR] " & msg
End Sub

Private Sub LogWarning(msg As String)
    Debug.Print "[WARN] " & msg
End Sub

Private Sub FlagVariance(label As String, delta As Double)
    MsgBox "Variance flagged in " & label & ": Δ" & Format$(delta, "0.00"), vbExclamation
End Sub

'===========================  EXTERNAL DATA STUBS  ====================================
Private Function FetchFromBloomberg(fieldCode As String) As Double
    FetchFromBloomberg = Nz(Worksheets("Bloomberg_PriceCache").Range("Last_Price").Value, 0)
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

