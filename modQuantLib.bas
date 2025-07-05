'======================================================================
'  modQuantLib - QuantLib-style valuation helpers
'======================================================================
Option Private Module
Option Explicit
Option Compare Text

Public Function DCF_PV(cf As Double, rate As Double, period As Long) As Double
    DCF_PV = cf / (1 + rate) ^ period
End Function

Public Function DCF_NPV(flows As Variant, rates As Variant) As Double
    Dim i As Long, pv As Double
    If UBound(flows) <> UBound(rates) Then Err.Raise vbObjectError + 2001, , "Rate/flow mismatch"
    For i = LBound(flows) To UBound(flows)
        pv = pv + flows(i) / (1 + rates(i)) ^ i
    Next i
    DCF_NPV = pv
End Function
