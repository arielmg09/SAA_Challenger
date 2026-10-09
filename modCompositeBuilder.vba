Option Explicit
Option Base 1

Public Function CompositeBuilder(sIn() As String, sRet() As String, sCompIn() As String, sCompRet() As String, sCompWgt() As String) As Boolean
    Dim sACM() As String, sYield() As String, sCompYield() As String, sRetDaily() As String, sCompRetDaily() As String, sCompWgtDaily() As String
    Dim wsOut As Worksheet
    Const sOut = "YieldData", sIndOut$ = "CompIndexData", sCompRetOut$ = "CompReturnData", iWgt% = 3
        If CalcDailyReturns(sIn, sRetDaily) = False Then Exit Function
        If LoadAssetClassMapAndReturns(sACM, sRet) = False Then Exit Function
        If BuildCompositeReturns(sACM, sRetDaily, sCompRetDaily, sCompWgtDaily, iWgt) = False Then Exit Function
        If CalcDailyCompositeIndex(sCompRetDaily, sCompIn) = False Then Exit Function
        If BuildCompositeReturns(sACM, sRet, sCompRet, sCompWgt, iWgt) = False Then Exit Function
        If GetCompositeYield(sACM, sYield, sCompYield) = False Then Exit Function
        If AdjustCompositeNames(sACM, sCompIn, sCompRet, sCompRetDaily) = False Then Exit Function
        If AddRow2End(sYield) = False Then Exit Function
        If AppendString2String(sYield, sCompYield) = False Then Exit Function
        If CheckSheet(ThisWorkbook.Name, sCompRetOut, wsOut) = False Then Exit Function
        sCompRet(1, 1) = "Date"
        If StringArrayOut(wsOut, sCompRet, True) = False Then Exit Function
        wsOut.Cells(1, 1) = "Total Return"
        If CheckSheet(ThisWorkbook.Name, sIndOut, wsOut) = False Then Exit Function
        sCompIn(1, 1) = "Date"
        If StringArrayOut(wsOut, sCompIn, True) = False Then Exit Function
        wsOut.Cells(1, 1) = "Total Return Index"
        If CheckSheet(ThisWorkbook.Name, sOut, wsOut) = False Then Exit Function
        If StringArrayOut(wsOut, sYield, True) = False Then Exit Function
    CompositeBuilder = True
End Function
Private Function CalcDailyReturns(sIn() As String, sRetDaily() As String) As Boolean
    Dim iI%, iJ%
        ReDim sRetDaily(UBound(sIn, 1) - 1, UBound(sIn, 2))
        For iJ = 1 To UBound(sIn, 2)
            sRetDaily(1, iJ) = sIn(1, iJ)
            For iI = UBound(sIn, 1) To 3 Step -1
                Select Case iJ
                    Case 1: sRetDaily(iI - 1, iJ) = sIn(iI - 1, iJ)
                    Case Else: If IsNumeric(sIn(iI, iJ)) And IsNumeric(sIn(iI - 1, iJ)) Then If CDbl(sIn(iI, iJ)) <> 0 Then sRetDaily(iI - 1, iJ) = CStr(CDbl(sIn(iI - 1, iJ)) / CDbl(sIn(iI, iJ)) - 1)
                End Select
            Next iI
        Next iJ
    CalcDailyReturns = True
End Function
Private Function LoadAssetClassMapAndReturns(sACM() As String, sRet() As String) As Boolean
    Const sSACM$ = "AssetClassMap", sSRet$ = "ReturnData"
        If Load2StringViaVariant(ThisWorkbook.Name, sSACM, sACM) = False Then Exit Function
        If CheckColumnsMatchAndInOrder(sACM, sRet) = False Then Exit Function
    LoadAssetClassMapAndReturns = True
End Function
Private Function CheckColumnsMatchAndInOrder(sACM() As String, sRet() As String) As Boolean
    Dim iI%
        If UBound(sACM, 1) <> UBound(sRet, 2) - 1 Then Exit Function
        For iI = 1 To UBound(sACM, 1)
            If sACM(iI, 1) <> sRet(1, iI + 1) Then Exit Function
        Next iI
    CheckColumnsMatchAndInOrder = True
End Function
Private Function BuildCompositeReturns(sACM() As String, sRet() As String, sCompRet() As String, sCompWgt() As String, iWgt%) As Boolean
    Dim iI%, iJ%, iTotWgt%, iRowNum%
        If CheckWeightsColumn(sACM, iWgt) = False Then Exit Function
        ReDim sCompRet(UBound(sRet, 1), UBound(sRet, 2))
        sCompRet = sRet
        ReDim Preserve sCompRet(UBound(sCompRet, 1), 1)
        ReDim sCompWgt(UBound(sACM, 1))
        For iI = 1 To UBound(sACM, 1)
            If sACM(iI, iWgt) = 1 Then
                ReDim Preserve sCompRet(UBound(sCompRet, 1), UBound(sCompRet, 2) + 1)
                For iJ = 1 To UBound(sCompRet, 1)
                    sCompRet(iJ, UBound(sCompRet, 2)) = sRet(iJ, iI + 1)
                Next iJ
                sCompWgt(iI) = 1
            Else
                If FindNumOfRows4Composite(sACM, iRowNum, iWgt, iI) = False Then Exit Function
                If CheckCompositeWeights(sACM, iRowNum, iWgt, iI) = False Then Exit Function
                If AggregateCompositeReturnsWithRebalancing(sACM, sRet, sCompRet, iRowNum, iWgt, iI, sCompWgt) = False Then Exit Function
            End If
        Next iI
    BuildCompositeReturns = True
End Function
Private Function CheckWeightsColumn(sACM() As String, iWgt%) As Boolean
    Dim iI%
        For iI = 1 To UBound(sACM, 1)
            If IsNumeric(sACM(iI, iWgt)) = False Then Exit Function
        Next iI
    CheckWeightsColumn = True
End Function
Public Function FindNumOfRows4Composite(sACM() As String, iRowNum%, iWgt%, iX%) As Boolean
'Find number of sub assets to include in the composite
    Dim iI%
        iRowNum = 1
        For iI = iX To UBound(sACM, 1) - 1
            If sACM(iI, iWgt + 1) = sACM(iI + 1, iWgt + 1) Then iRowNum = iRowNum + 1 Else Exit For
        Next iI
    FindNumOfRows4Composite = True
End Function
Private Function CheckCompositeWeights(sACM() As String, iRowNum%, iWgt%, iX%) As Boolean
'check weights for composite add to 1. If within 4sd of 1, then adjust adjust largest weight sub-asset to make total equal 1
    Dim iI%, iMax%
    Dim dX#, dMax#
        If iX + (iRowNum - 1) > UBound(sACM, 1) Then Exit Function
        For iI = iX To iX + (iRowNum - 1)
            If IsNumeric(sACM(iI, iWgt)) Then
                sACM(iI, iWgt) = CStr(Round(CDbl(sACM(iI, iWgt)), 8))
                dX = dX + CDbl(sACM(iI, iWgt))
                If CDbl(sACM(iI, iWgt)) > dMax Then
                    dMax = CDbl(sACM(iI, iWgt))
                    iMax = iI
                End If
            Else: Exit Function
            End If
        Next iI
        If Round(dX, 4) <> 1 Then Exit Function
        If dX <> 1 Then
            sACM(iMax, iWgt) = CStr(CDbl(sACM(iMax, iWgt)) + (1 - dX))
        End If
    CheckCompositeWeights = True
End Function
Private Function AggregateCompositeReturnsWithRebalancing(sACM() As String, sRet() As String, sCompRet() As String, iRowNum%, iWgt%, iX%, sCompWgt() As String) As Boolean
    Dim iI%, iJ%, iK%, iStart%, iCount%
    Dim dWgt() As Double, dWgtFlt() As Double, dX#
        If LoadWeights2Double(sACM, dWgt, iRowNum, iWgt, iX) = False Then Exit Function
        If FindStartDate(sRet, iRowNum, iX, iStart) = False Then Exit Function
        ReDim dWgtFlt(UBound(dWgt))
        dWgtFlt = dWgt
        ReDim Preserve sCompRet(UBound(sCompRet, 1), UBound(sCompRet, 2) + 1)
        sCompRet(1, UBound(sCompRet, 2)) = sACM(iX, iWgt + 1)
        For iI = iStart To 2 Step -1
            dX = 0
            iCount = 1
            For iJ = iX + 1 To iX + iRowNum
                dX = dX + dWgtFlt(iCount) * CDbl(sRet(iI, iJ))
                dWgtFlt(iCount) = dWgtFlt(iCount) * (1 + CDbl(sRet(iI, iJ)))
                iCount = iCount + 1
            Next iJ
            sCompRet(iI, UBound(sCompRet, 2)) = CStr(dX)
            If AdjustFloatedWeights(dWgt, dWgtFlt, sRet(iI, 1)) = False Then Exit Function
        Next iI
        For iI = iX To iX + (iRowNum - 1)
            sCompWgt(iI) = CStr(dWgtFlt(iI - iX + 1))
        Next iI
        iX = iX + (iRowNum - 1)
    AggregateCompositeReturnsWithRebalancing = True
End Function
Private Function LoadWeights2Double(sACM() As String, dWgt() As Double, iRowNum%, iWgt%, iX%) As Boolean
    Dim iI%
        ReDim dWgt(iRowNum)
        For iI = iX To iX + (iRowNum - 1)
            dWgt(iI - (iX - 1)) = CDbl(sACM(iI, iWgt))
        Next iI
    LoadWeights2Double = True
End Function
Private Function FindStartDate(sRet() As String, iRowNum%, iX%, iStart%) As Boolean
    Dim iI%, iJ%
        iStart = UBound(sRet, 1)
        For iJ = iX + 1 To iX + iRowNum
            For iI = 2 To UBound(sRet, 1)
                If IsNumeric(sRet(iI, iJ)) = False Then If iI < iStart Then iStart = iI - 1
            Next iI
        Next iJ
    FindStartDate = True
End Function
Private Function AdjustFloatedWeights(dWgt() As Double, dWgtFlt() As Double, sDate$) As Boolean
    Dim iI%
    Dim dX#
        If Month(CDate(sDate)) = 12 Then
            ReDim dWgtFlt(UBound(dWgt))
            dWgtFlt = dWgt
        Else
            dX = Application.WorksheetFunction.Sum(dWgtFlt)
            For iI = 1 To UBound(dWgtFlt)
                dWgtFlt(iI) = dWgtFlt(iI) / dX
            Next iI
        End If
    AdjustFloatedWeights = True
End Function
Private Function CalcDailyCompositeIndex(sCompRetDaily() As String, sCompIn() As String) As Boolean
    Dim iI%, iJ%
        ReDim sCompIn(UBound(sCompRetDaily, 1) + 1, UBound(sCompRetDaily, 2))
        sCompIn(UBound(sCompIn, 1), 1) = CDate(sCompRetDaily(UBound(sCompRetDaily, 1), 1)) - 1
        For iJ = 1 To UBound(sCompRetDaily, 2)
            sCompIn(1, iJ) = sCompRetDaily(1, iJ)
            If iJ <> 1 Then If IsNumeric(sCompRetDaily(UBound(sCompRetDaily, 1), iJ)) Then sCompIn(UBound(sCompIn, 1), iJ) = 100
            For iI = UBound(sCompRetDaily, 1) To 2 Step -1
                Select Case iJ
                    Case 1: sCompIn(iI, iJ) = sCompRetDaily(iI, iJ)
                    Case Else
                        If IsNumeric(sCompRetDaily(iI, iJ)) And IsNumeric(sCompIn(iI + 1, iJ)) = False Then
                            sCompIn(iI + 1, iJ) = 100
                        End If
                        If IsNumeric(sCompRetDaily(iI, iJ)) And IsNumeric(sCompIn(iI + 1, iJ)) Then sCompIn(iI, iJ) = CStr(CDbl(sCompIn(iI + 1, iJ)) * (1 + CDbl(sCompRetDaily(iI, iJ))))
                End Select
            Next iI
        Next iJ
    CalcDailyCompositeIndex = True
End Function
Private Function GetCompositeYield(sACM() As String, sMedianYield() As String, sMedianCompYield() As String) As Boolean
    Dim sFSYield() As String, sHistYield() As String, sCodeName() As String, sYldComp() As String
    Dim wsOut As Worksheet
    Const sSHistYield$ = "HistoricYieldData", sSYldComp$ = "CompYieldData", iYears# = 20 '***** because of data limitation, can not set number of years to be greater than 22!!!
        If GetRawData(sFSYield, sHistYield) = False Then Exit Function
        If StoreHistoricCodesAndNames(sHistYield, sCodeName) = False Then Exit Function
        If AdjustFactSetData(sFSYield) = False Then Exit Function
        If AggregateHistoricYieldData(sFSYield, sHistYield, sYldComp, sACM, iYears) = False Then Exit Function
        If Get12MthMedianYield(sHistYield, sMedianYield) = False Then Exit Function
        If AddBackCodes4Names(sHistYield, sCodeName) = False Then Exit Function
        If Get12MthMedianYield(sYldComp, sMedianCompYield) = False Then Exit Function
        If CheckSheet(ThisWorkbook.Name, sSHistYield, wsOut) = False Then Exit Function
        If StringArrayOut(wsOut, sHistYield, True) = False Then Exit Function
        If CheckSheet(ThisWorkbook.Name, sSYldComp, wsOut) = False Then Exit Function
        sYldComp(1, 1) = "Date"
        If StringArrayOut(wsOut, sYldComp, True) = False Then Exit Function
        wsOut.Cells(1, 1) = "Yield"
        If PrepareYield4Output(sMedianYield, sMedianCompYield) = False Then Exit Function
    GetCompositeYield = True
End Function
Private Function GetRawData(sFSYield() As String, sHistYield() As String) As Boolean
    Const sSFSYield$ = "FactSet", sSHistYield$ = "HistoricYieldData"
        If Load2StringViaVariant(ThisWorkbook.Name, sSFSYield, sFSYield) = False Then Exit Function
        If Load2StringViaVariant(ThisWorkbook.Name, sSHistYield, sHistYield) = False Then Exit Function
    GetRawData = True
End Function
Private Function StoreHistoricCodesAndNames(sHistYield() As String, sCodeName() As String) As Boolean
        ReDim sCodeName(UBound(sHistYield, 1), UBound(sHistYield, 2))
        sCodeName = sHistYield
        If ResizeRowsInStringArray(sCodeName, 2) = False Then Exit Function
    StoreHistoricCodesAndNames = True
End Function
Private Function AdjustFactSetData(sFSYield() As String) As Boolean
    Dim iI%
        For iI = 3 To UBound(sFSYield, 1)
            If IsDate(sFSYield(iI, 1)) Then
                sFSYield(iI, 1) = Format(Application.WorksheetFunction.EoMonth(CDate(sFSYield(iI, 1)), 0), "dd/mm/yyyy")
            End If
        Next iI
'        If SpliceOtherData(sFSYield, "MS706427", "FI4890GB") = False Then Exit Function ' no data for MSCI UK IMI Liquid Real Estate CMBR, use FactSet
    AdjustFactSetData = True
End Function
Private Function SpliceOtherData(sYield() As String, sY$, sZ$) As Boolean
    Dim iI%, iY%, iZ%
        If FindColumnPositionInStringArray(sY, sYield, iY) = False Or iY = 0 Then Exit Function
        If FindColumnPositionInStringArray(sZ, sYield, iZ) = False Or iZ = 0 Then Exit Function
        For iI = 3 To UBound(sYield, 1)
            If IsNumeric(sYield(iI, iY)) = False And IsNumeric(sYield(iI, iZ)) Then sYield(iI, iY) = CStr(Round(CDbl(sYield(iI, iZ)), 5))
        Next iI
    SpliceOtherData = True
End Function
Private Function AggregateHistoricYieldData(sFSYield() As String, sHistYield() As String, sYldComp() As String, sWgtComp() As String, iYears%) As Boolean
    Const iWgt% = 3
        If SpliceLatestYieldData(sFSYield, sHistYield) = False Then Exit Function
        If RemoveFirstRowInString(sHistYield) = False Then Exit Function
        If AdjustLinkerYield(sHistYield) = False Then Exit Function
        If BuildCompositeYields(sWgtComp, sHistYield, sYldComp, iWgt) = False Then Exit Function
        If AdjustYieldCompositeNames(sWgtComp, sYldComp) = False Then Exit Function
        If ResizeRowsInStringArray(sYldComp, iYears * 12 + 2) = False Then Exit Function
        sYldComp(1, 1) = "Yield"
    AggregateHistoricYieldData = True
End Function
Private Function SpliceLatestYieldData(sFSYield() As String, sHistYield() As String) As Boolean
    Dim iI%, iJ%, iRow%, iCol() As Integer
        If sFSYield(3, 1) <> sHistYield(3, 1) And CDate(sFSYield(3, 1)) > CDate(sHistYield(3, 1)) Then
            If MatchColumnPositions(sHistYield, sFSYield, iCol) = False Then Exit Function
            If FindRowPositionInStringArray(sHistYield(3, 1), sFSYield, iRow) = False Or iRow < 4 Then Exit Function
            If AddRow2Start(sHistYield, iRow - 3) = False Then Exit Function
            For iJ = 1 To UBound(sHistYield, 2)
                sHistYield(1, iJ) = sHistYield(iRow - 2, iJ)
                sHistYield(2, iJ) = sHistYield(iRow - 1, iJ)
                If iJ = 1 Then
                    For iI = 3 To iRow - 1
                        sHistYield(iI, iJ) = sFSYield(iI, 1)
                    Next iI
                Else
                    For iI = 3 To iRow - 1
                        If IsNumeric(sFSYield(iI, iCol(iJ - 1))) Then sHistYield(iI, iJ) = CStr(CDbl(sFSYield(iI, iCol(iJ - 1)) / 100)) Else sHistYield(iI, iJ) = ""
                    Next iI
                End If
            Next iJ
        End If
    SpliceLatestYieldData = True
End Function
Private Function AdjustLinkerYield(sYield() As String) As Boolean
    Dim iI%, iLink%, iConv%
    Const sLink$ = "ICE BofA UK Inflation-Linked Gilts", sConv$ = "ICE BofA UK Conventional Gilts"
        If FindColumnPositionInStringArray(sLink, sYield, iLink) = False Or iLink = False Then Exit Function
        If FindColumnPositionInStringArray(sConv, sYield, iConv) = False Or iConv = False Then Exit Function
        For iI = 2 To UBound(sYield, 1)
            sYield(iI, iLink) = sYield(iI, iConv)
        Next iI
    AdjustLinkerYield = True
End Function
Private Function BuildCompositeYields(sACM() As String, sYld() As String, sYldComp() As String, iWgt%) As Boolean
    Dim iI%, iJ%, iTotWgt%, iRowNum%
        If CheckWeightsColumn(sACM, iWgt) = False Then Exit Function
        ReDim sYldComp(UBound(sYld, 1), UBound(sYld, 2))
        sYldComp = sYld
        ReDim Preserve sYldComp(UBound(sYldComp, 1), 1)
        For iI = 1 To UBound(sACM, 1) - 1
            If sACM(iI, iWgt) = 1 Then
                ReDim Preserve sYldComp(UBound(sYldComp, 1), UBound(sYldComp, 2) + 1)
                For iJ = 1 To UBound(sYldComp, 1)
                    sYldComp(iJ, UBound(sYldComp, 2)) = sYld(iJ, iI + 1)
                Next iJ
            Else
                If FindNumOfRows4Composite(sACM, iRowNum, iWgt, iI) = False Then Exit Function
                If CheckCompositeWeights(sACM, iRowNum, iWgt, iI) = False Then Exit Function
                If AggregateCompositeYield(sACM, sYld, sYldComp, iRowNum, iWgt, iI) = False Then Exit Function
            End If
        Next iI
    BuildCompositeYields = True
End Function
Private Function AggregateCompositeYield(sACM() As String, sYld() As String, sYldComp() As String, iRowNum%, iWgt%, iX%) As Boolean
    Dim iI%, iJ%, iK%, iStart%, iCount%
    Dim dWgt() As Double, dWgtFlt() As Double, dX#
        If LoadWeights2Double(sACM, dWgt, iRowNum, iWgt, iX) = False Then Exit Function
        If FindStartDate(sYld, iRowNum, iX, iStart) = False Then Exit Function
        ReDim Preserve sYldComp(UBound(sYldComp, 1), UBound(sYldComp, 2) + 1)
        sYldComp(1, UBound(sYldComp, 2)) = sACM(iX, iWgt + 1)
        For iI = iStart To 2 Step -1
            dX = 0
            iCount = 1
            For iJ = iX + 1 To iX + iRowNum
                dX = dX + dWgt(iCount) * CDbl(sYld(iI, iJ))
                iCount = iCount + 1
            Next iJ
            sYldComp(iI, UBound(sYldComp, 2)) = CStr(dX)
        Next iI
        iX = iX + (iRowNum - 1)
    AggregateCompositeYield = True
End Function
Private Function AdjustYieldCompositeNames(sACM() As String, sYldComp() As String) As Boolean
    Dim iI%, iJ%, iX%
    Const iName% = 4
        For iJ = 2 To UBound(sYldComp, 2)
            iX = 0
            If FindRowPositionInStringArray(sYldComp(1, iJ), sACM, iX) = False Then Exit Function
            If iX <> 0 Then sYldComp(1, iJ) = sACM(iX, iName)
        Next iJ
    AdjustYieldCompositeNames = True
End Function
Private Function Get12MthMedianYield(sYldComp() As String, sMedianCompYield() As String) As Boolean
    Dim iI%, iJ&
    Dim dX() As Double
        ReDim sMedianCompYield(UBound(sYldComp, 1), UBound(sYldComp, 2))
        sMedianCompYield = sYldComp
        If ResizeRowsInStringArray(sMedianCompYield, 2) = False Then Exit Function
        For iJ = 2 To UBound(sYldComp, 2)
            ReDim dX(12)
            For iI = 1 To 12
                If IsNumeric(sYldComp(iI + 1, iJ)) Then dX(iI) = CDbl(sYldComp(iI + 1, iJ)) Else dX(iI) = -999999
            Next iI
            sMedianCompYield(2, iJ) = Application.WorksheetFunction.Median(dX)
        Next iJ
    Get12MthMedianYield = True
End Function
Private Function AddBackCodes4Names(sHistYield() As String, sCodeName() As String) As Boolean
    Dim iI%, iJ%, iCol%
        If AddRow2Start(sHistYield) = False Then Exit Function
        For iJ = 2 To UBound(sHistYield, 2)
            If sHistYield(2, iJ) = sCodeName(2, iJ) Then sHistYield(1, iJ) = sCodeName(1, iJ)
        Next iJ
        sHistYield(1, 1) = "Date"
    AddBackCodes4Names = True
End Function
Private Function PrepareYield4Output(sMedianYield() As String, sMedianCompYield() As String) As Boolean
        If TransposeString(sMedianYield) = False Then Exit Function
        If TransposeString(sMedianCompYield) = False Then Exit Function
        If RemoveFirstRowInString(sMedianYield) = False Then Exit Function
        If RemoveFirstRowInString(sMedianCompYield) = False Then Exit Function
    PrepareYield4Output = True
End Function
Private Function AdjustCompositeNames(sACM() As String, sCompIn() As String, sCompRet() As String, sCompRetDaily() As String) As Boolean
    Dim iI%, iJ%, iX%
    Const iName% = 4
        If UBound(sCompIn, 2) <> UBound(sCompRet, 2) Or UBound(sCompIn, 2) <> UBound(sCompRetDaily, 2) Then Exit Function
        For iJ = 2 To UBound(sCompIn, 2)
            iX = 0
            If sCompIn(1, iJ) <> sCompRet(1, iJ) Or sCompIn(1, iJ) <> sCompRetDaily(1, iJ) Then Exit Function
            If FindRowPositionInStringArray(sCompIn(1, iJ), sACM, iX) = False Then Exit Function
            If iX <> 0 Then
                sCompIn(1, iJ) = sACM(iX, iName)
                sCompRet(1, iJ) = sACM(iX, iName)
                sCompRetDaily(1, iJ) = sACM(iX, iName)
            End If
        Next iJ
    AdjustCompositeNames = True
End Function

