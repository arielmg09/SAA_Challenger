Option Explicit
Option Base 1

Dim bOpt2(2) As Boolean, bOpt3(2) As Boolean, iYrs2%, iYrs3%, iYrsRA%, iOpt%, iRA%, iOptACM%, iWght() As Integer

Public Sub MainControl()
    Dim iM%, iMVol%, iMShort%
    Dim sS() As String, sIn() As String, sCompIn() As String, sRet() As String, sCompRet() As String, sYld() As String, sCorr() As String, sVol() As String, sVol4RA() As String _
        , sVolRet4RA() As String, sCovar() As String, sStats() As String, sPIn$, sCompCorr() As String, sCompVol() As String, sCompCovar() As String, sCompWgt() As String, sCompStats() As String
    Dim dtShort As Date
    Const sFIn$ = "Index Data_Portfolio Performance Summary.xlsx" ' Use this version of the file to load in Soc Gen CTA Index (£ hedged)
        sPIn = "\\performance\data\Performance\Index Data"
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        With frmDataAnalysis 'checked - this works
            .StartUpPosition = 0
            .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
            .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
            .Show
        End With
                
        If iOptACM = 1 Then If ImportAndOverwriteACM = False Then GoTo ErrTrap
        If LoadSheetNames(sS) = False Then GoTo ErrTrap
        If GetRawData(sFIn, sPIn, sS(1), sIn) = False Then GoTo ErrTrap 'get index data and strip out only those series required in the correct order
        
        
        If MyConvert2MonthlyReturns(sIn, sRet) = False Then Exit Sub 'checked - this works ... converts daily index data to monthly returns
        If SpliceOnMoreHistory(sS(2), sRet) = False Then Exit Sub
        If FindShortestSeries(sRet, dtShort) = False Then GoTo ErrTrap 'checked - this works
        iMShort = MyDateDif(dtShort, CDate(sRet(2, 1)), "M")
        iM = iMShort
        iMVol = iMShort
        If bOpt2(2) Then If iYrs2 * 12 < iM Then iM = iYrs2 * 12 'number of months to use for correlation and covariance matrices. Checked - this works
        If bOpt2(2) Then If iYrs2 * 12 < iMVol Then iMVol = iYrs2 * 12 'number of months to use for volatility matrix. Checked - this works ... currently set to the same as the correlation and covariance matrices
        If GetCorrCovarVol(sRet, sCorr, sCovar, sVol, iM, iMVol) = False Then GoTo ErrTrap
        If CompositeBuilder(sIn, sRet, sCompIn, sCompRet, sCompWgt) = False Then Exit Sub
        If GetCorrCovarVol(sCompRet, sCompCorr, sCompCovar, sCompVol, iM, iMVol) = False Then Exit Sub
        If CombineCorrCovarVol(sCorr, sCompCorr, sCovar, sCompCovar, sVol, sCompVol) = False Then Exit Sub
        iM = iMShort
        If bOpt3(2) Then If iYrs3 * 12 < iM Then iM = iYrs3 * 12 'number of months to use for volatility for return assumptions. Checked - this works
        If CalculateRAVolatility(sRet, sVolRet4RA, iM) = False Then GoTo ErrTrap 'checked - this works
        If CalculateRAReturn(sRet, sVolRet4RA, iM) = False Then GoTo ErrTrap 'checked - this works
        iM = iMShort
        If iOpt < 0 Then
            iM = -12 * iOpt
        Else: If iOpt = 2 Then iM = -999
        End If
        If CalculateStats(sIn, sRet, sStats, iM) = False Then GoTo ErrTrap 'checked - this works
        If CalculateStats(sCompIn, sCompRet, sCompStats, iM, True) = False Then GoTo ErrTrap
        If OutputData(sS, sRet, sCorr, sVol, sCovar) = False Then GoTo ErrTrap 'checked - this works
        If OutputStats(sStats, sCompStats) = False Then GoTo ErrTrap 'checked - this works
        If MainReturnAssumptions(sVolRet4RA, iRA, iWght) = False Then Exit Sub 'checked - this works
                
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        MsgBox "Data updated successfully"
        
        Exit Sub
ErrTrap:
        MsgBox "Error executing MainContol"
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
End Sub

Public Function DataAnalysisPassOut(bOptDA() As Boolean, bOptVC() As Boolean, bOptRA() As Boolean, iYrsDA%, iYrsVC%, iYrsRA%, iOptRA%, iACM%, iW() As Integer) As Boolean
    Dim iI%
        If bOptDA(1) Then iOpt = 1
        If bOptDA(2) Then iOpt = 2
        If bOptDA(3) Then iOpt = -1 * iYrsDA
        bOpt2(1) = bOptVC(1)
        bOpt2(2) = bOptVC(2)
        iYrs2 = iYrsVC
        bOpt3(1) = bOptRA(1)
        bOpt3(2) = bOptRA(2)
        iYrs3 = iYrsRA
        iRA = iOptRA
        ReDim iWght(UBound(iW))
        For iI = 1 To UBound(iW)
            iWght(iI) = iW(iI)
        Next iI
        iOptACM = iACM
    DataAnalysisPassOut = True
End Function
Private Function ImportAndOverwriteACM() As Boolean
    Const sPIn$ = "X:\Covent Garden\Resources\HIM\Strategic Asset Allocation", sFIn$ = "Composite Weights.xlsx", sShACM$ = "AssetClassMap"
        If FileOpen(sFIn, sPIn) = False Then Exit Function
        If CopyDataSheet2Sheet(sFIn, ThisWorkbook.Name, sShACM) = False Then Exit Function
        If FileClose(sFIn) = False Then Exit Function
    ImportAndOverwriteACM = True
End Function
Private Function LoadSheetNames(sX() As String) As Boolean
    Dim iI%, vX() As Variant
        vX = Array("IndexData", "HistoricReturnData", "CorrelationMatrix", "CovarianceMatrix", "VolatilityMatrix")
        ReDim sX(UBound(vX))
        For iI = 1 To UBound(sX, 1)
            sX(iI) = vX(iI)
        Next iI
    LoadSheetNames = True
End Function
Private Function GetRawData(sFIn$, sPIn$, sS$, sY() As String) As Boolean
    Dim iI%, iJ%, iY() As Integer
    Dim sX() As String
    Const sACM$ = "AssetClassMap"
        If FileOpen(sFIn, sPIn) = False Then Exit Function
        If Load2StringViaVariant(sFIn, sS, sX) = False Then Exit Function
        If FileClose(sFIn) = False Then Exit Function
        If RemoveFirstRowInString(sX) = False Then Exit Function
        If RemoveFirstRowInString(sX) = False Then Exit Function
        If Load2StringViaVariant(ThisWorkbook.Name, sACM, sY) = False Then Exit Function
        If AddRow2Start(sY) = False Then Exit Function
        If TransposeString(sY) = False Then Exit Function
        If MatchColumnPositions(sY, sX, iY) = False Then Exit Function
        ReDim sY(UBound(sX, 1), UBound(iY) + 1)
        For iI = 1 To UBound(sX, 1)
            sY(iI, 1) = sX(iI, 1)
            For iJ = 1 To UBound(iY)
                If iY(iJ) <> 0 Then sY(iI, iJ + 1) = sX(iI, iY(iJ))
            Next iJ
        Next iI
    GetRawData = True
End Function
Private Function MyConvert2MonthlyReturns(sIn() As String, sInMRet() As String) As Boolean
    Dim sX() As String
        ReDim sX(UBound(sIn, 1), UBound(sIn, 2))
        sX = sIn
        If Convert2Monthly(sX) = False Then Exit Function
        If GetReturns(sX, sInMRet) = False Then Exit Function
    MyConvert2MonthlyReturns = True
End Function
Private Function FindShortestSeries(sIn() As String, dtShort As Date) As Boolean
    Dim iI%, iJ%
        If CDate(sIn(2, 1)) < CDate(sIn(UBound(sIn, 1), 1)) Then dtShort = CDate(sIn(2, 1)) Else dtShort = CDate(sIn(UBound(sIn, 1), 1))
        For iJ = 2 To UBound(sIn, 2)
            For iI = 2 To UBound(sIn, 1)
                If IsNumeric(sIn(iI, iJ)) = False Or sIn(iI, iJ) = "" Then
                    If CDate(sIn(iI, 1)) > dtShort Then dtShort = CDate(sIn(iI, 1))
                    Exit For
                End If
            Next iI
        Next iJ
        If dtShort = CDate(sIn(2, 1)) Then dtShort = CDate(sIn(UBound(sIn, 1), 1)) Else dtShort = CDate(Application.WorksheetFunction.EoMonth(dtShort, 0))
    FindShortestSeries = True
End Function
Private Function MyDateDif(dtStart As Date, dtEnd As Date, sFreq$) As Integer
    Dim iX%
    Dim dtX As Date
        MyDateDif = 0
        If dtStart > dtEnd Then Exit Function
        If UCase(sFreq) = "M" Then
            dtX = dtStart
            While dtX <= dtEnd
                dtX = CDate(Application.WorksheetFunction.EoMonth(dtX, 1))
                If dtX <= dtEnd Then MyDateDif = MyDateDif + 1
                If MyDateDif > 999 Then
                    MyDateDif = 0
                    Exit Function
                End If
            Wend
        End If
End Function
Private Function GetReturns(sIn() As String, sRet() As String) As Boolean
    Dim iI%, iJ%
        ReDim sRet(UBound(sIn, 1), UBound(sIn, 2))
        sRet = sIn
        For iJ = 2 To UBound(sRet, 2)
            For iI = 2 To UBound(sRet, 1) - 1
                If IsNumeric(sRet(iI, iJ)) And IsNumeric(sRet(iI + 1, iJ)) Then
                    sRet(iI, iJ) = CStr(CDbl(sRet(iI, iJ)) / CDbl(sRet(iI + 1, iJ)) - 1)
                Else: sRet(iI, iJ) = ""
                End If
            Next iI
        Next iJ
        If TransposeString(sRet) = False Then Exit Function
        ReDim Preserve sRet(UBound(sRet, 1), UBound(sRet, 2) - 1)
        If TransposeString(sRet) = False Then Exit Function
    GetReturns = True
End Function
Private Function SpliceOnMoreHistory(sS$, sRet() As String) As Boolean
    Dim iI%, iJ%, iStart%, iX%, iZ%, iY() As Integer
    Dim sX$, sY() As String, sZ() As String
    Dim wsOut As Worksheet
        iX = UBound(sRet, 1)
        If Load2StringViaVariant(ThisWorkbook.Name, sS, sY) = False Then Exit Function
        If MatchColumnPositions(sRet, sY, iY) = False Then Exit Function
        If FindGapsAndFill(sRet, sY, iY) = False Then Exit Function
        If FindStartDate(CDate(sRet(UBound(sRet, 1), 1)), sY, iStart, iZ) = False Then Exit Function
        If TransposeString(sRet) = False Then Exit Function
        ReDim Preserve sRet(UBound(sRet, 1), UBound(sRet, 2) + iZ)
        If TransposeString(sRet) = False Then Exit Function
        For iI = iStart To UBound(sY, 1)
            iX = iX + 1
            sRet(iX, 1) = sY(iI, 1)
            For iJ = 2 To UBound(sRet, 2)
                sRet(iX, iJ) = sY(iI, iY(iJ - 1))
            Next iJ
        Next iI
        If CheckSheet(ThisWorkbook.Name, sS, wsOut) = False Then Exit Function
        sX = sRet(1, 1)
        sRet(1, 1) = "Date"
        If StringArrayOut(wsOut, sRet, True) = False Then Exit Function
        sRet(1, 1) = sX
        wsOut.Cells(1, 1) = sX
    SpliceOnMoreHistory = True
End Function
Private Function FindGapsAndFill(sRet() As String, sY() As String, iY() As Integer) As Boolean
    Dim iI%, iJ%, iX%, iZ%, iStart%
        If FindStartDate(CDate(sY(2, 1)), sRet, iStart, iZ) = False Then Exit Function
        If iStart > 1 Then iStart = iStart - 1
        iX = 1
        For iI = iStart To UBound(sRet, 1)
            iX = iX + 1
            If CDate(sRet(iI, 1)) <> CDate(sY(iX, 1)) Then Exit Function
            For iJ = 2 To UBound(sRet, 2)
                If IsNumeric(sRet(iI, iJ)) = False And IsNumeric(sY(iX, iY(iJ - 1))) Then
                    If sRet(1, iJ) <> sY(1, iY(iJ - 1)) Then Exit Function
                    sRet(iI, iJ) = sY(iX, iY(iJ - 1))
                End If
            Next iJ
        Next iI
    FindGapsAndFill = True
End Function
Private Function FindStartDate(dtY As Date, sY() As String, iStart%, iZ%) As Boolean
    Dim iI%
        If dtY > CDate(sY(2, 1)) Then GoTo MyContinue
        For iI = 2 To UBound(sY, 1)
            If dtY = CDate(sY(iI, 1)) Then
                iStart = iI + 1
                Exit For
            End If
        Next iI
        If iStart > UBound(sY, 1) Then Exit Function
        If CDate(sY(iStart, 1)) <> CDate(Application.WorksheetFunction.EoMonth(dtY, -1)) Then Exit Function
        iZ = UBound(sY, 1) - (iStart - 1)
MyContinue:
    FindStartDate = True
End Function
Private Function GetCorrCovarVol(sRet() As String, sCorr() As String, sCovar() As String, sVol() As String, iM%, iMVol%) As Boolean
    Dim iI%, iJ%
    Dim dX() As Double, dY() As Double, dVol() As Double, dZ#
        ReDim sCorr(UBound(sRet, 2), UBound(sRet, 2))
        ReDim sCovar(UBound(sRet, 2), UBound(sRet, 2))
        ReDim sVol(UBound(sRet, 2) - 1, 2)
        For iI = 2 To UBound(sRet, 2)
            If Load2Double(sRet, dX, iI, iM) = False Then Exit Function
            If Load2Double(sRet, dVol, iI, iMVol) = False Then Exit Function
            For iJ = 2 To UBound(sRet, 2)
                If Load2Double(sRet, dY, iJ, iM) = False Then Exit Function
                If UBound(dX) < UBound(dY) Then
                    ReDim Preserve dY(UBound(dX))
                ElseIf UBound(dX) > UBound(dY) Then
                    ReDim Preserve dX(UBound(dY))
                End If
                sCorr(iI, iJ) = Application.WorksheetFunction.Correl(dX, dY)
                sCovar(iI, iJ) = Application.WorksheetFunction.Covar(dX, dY)
            Next iJ
            sCorr(1, iI) = sRet(1, iI)
            sCorr(iI, 1) = sRet(1, iI)
            sCovar(1, iI) = sRet(1, iI)
            sCovar(iI, 1) = sRet(1, iI)
            sVol(iI - 1, 1) = sRet(1, iI)
            sVol(iI - 1, 2) = Application.WorksheetFunction.StDev(dVol) * Sqr(12) 'volatility matrix does NOT have matching time horizon as correlation and covariance matrices
        Next iI
        sCorr(1, 1) = Format(CDbl(UBound(dX)) / 12#, "0.0") + " years of data"
        sCovar(1, 1) = Format(CDbl(UBound(dX)) / 12#, "0.0") + " years of data"
        If AddRow2Start(sVol) = False Then Exit Function
        sVol(1, 1) = Format(CDbl(UBound(dVol)) / 12#, "0.0") + " years of data"
    GetCorrCovarVol = True
End Function
Private Function Load2Double(sX() As String, dX() As Double, iX%, iM%) As Boolean
    Dim iI%
        ReDim dX(iM)
        For iI = 2 To iM + 1
            If IsNumeric(sX(iI, iX)) Then
                dX(iI - 1) = CDbl(sX(iI, iX))
            Else
                ReDim Preserve dX(iI - 2)
                Exit For
            End If
        Next iI
    Load2Double = True
End Function
Private Function CombineCorrCovarVol(sCorr() As String, sCompCorr() As String, sCovar() As String, sCompCovar() As String, sVol() As String, sCompVol() As String) As Boolean
        If AddRow2End(sCorr) = False Then Exit Function
        If AddRow2End(sCovar) = False Then Exit Function
        If AddRow2End(sVol) = False Then Exit Function
        If AppendString2String(sCorr, sCompCorr) = False Then Exit Function
        If AppendString2String(sCovar, sCompCovar) = False Then Exit Function
        If AppendString2String(sVol, sCompVol) = False Then Exit Function
    CombineCorrCovarVol = True
End Function
Private Function CalculateRAVolatility(sRet() As String, sVolRet4RA() As String, iM%) As Boolean
    Dim iI%, iJ%, iX%
    Dim dX() As Double
        If iM < 0 Then iX = UBound(sRet, 1) Else iX = iM + 1
        If iX > UBound(sRet, 1) Then iX = UBound(sRet, 1)
        ReDim sVolRet4RA(2, UBound(sRet, 2))
        sVolRet4RA(2, 1) = "Annualised volatility for return assumptions"
        For iJ = 2 To UBound(sRet, 2)
            sVolRet4RA(1, iJ) = sRet(1, iJ)
            ReDim dX(1)
            dX(1) = CDbl(sRet(2, iJ))
            For iI = 3 To iX
                If IsNumeric(sRet(iI, iJ)) Then
                    ReDim Preserve dX(UBound(dX) + 1)
                    dX(UBound(dX)) = CDbl(sRet(iI, iJ))
                Else: Exit For
                End If
            Next iI
            sVolRet4RA(UBound(sVolRet4RA, 1), iJ) = CStr(Application.WorksheetFunction.StDev(dX) * Sqr(12))
        Next iJ
        iM = iX
        sVolRet4RA(1, 1) = Format(iM / 12, "0.0") & " years of data"
    CalculateRAVolatility = True
End Function
Private Function CalculateRAReturn(sRet() As String, sVolRet4RA() As String, iM%) As Boolean
    Dim sIndex() As String
        If CalculateAnnualisedReturn(sRet, sIndex, sVolRet4RA, iM) = False Then Exit Function
        sVolRet4RA(UBound(sVolRet4RA, 1), 1) = Replace(sVolRet4RA(UBound(sVolRet4RA, 1), 1), "Annualised return", "Annualised historic return") & " for return assumptions"
    CalculateRAReturn = True
End Function
Private Function CalculateAnnualisedReturn(sRet() As String, sIndex() As String, sStats() As String, iM%) As Boolean
    Dim iI%, iJ%
    Dim dtStart As Date, dtEnd As Date
        If AddRow2End(sStats) = False Then Exit Function
        sStats(UBound(sStats, 1), 1) = "Annualised return"
        If ConvertReturns2Index(sRet, sIndex) = False Then Exit Function
        dtEnd = CDate(sIndex(2, 1))
        For iJ = 2 To UBound(sIndex, 2)
            dtStart = dtEnd
            For iI = iM + 1 To 2 Step -1
                If IsNumeric(sIndex(iI, iJ)) Then Exit For
            Next iI
            If iI <= iM + 1 Then
                If CDate(sIndex(iI, 1)) < dtStart Then dtStart = CDate(sIndex(iI, 1)) Else Exit Function
                sStats(UBound(sStats, 1), iJ) = (sIndex(2, iJ) / sIndex(iI, iJ)) ^ (1 / Application.WorksheetFunction.YearFrac(dtStart, dtEnd)) - 1
            Else: Exit Function
            End If
        Next iJ
    CalculateAnnualisedReturn = True
End Function
Private Function ConvertReturns2Index(sRet() As String, sIndex() As String) As Boolean
    Dim bStart As Boolean
    Dim iI%, iJ%
        ReDim sIndex(UBound(sRet, 1) + 1, UBound(sRet, 2))
        For iI = 1 To UBound(sRet, 1)
            sIndex(iI, 1) = sRet(iI, 1)
        Next iI
        sIndex(UBound(sIndex, 1), 1) = Format(CDate(Application.WorksheetFunction.EoMonth(CDate(sIndex(UBound(sIndex, 1) - 1, 1)), -1)), "dd/mm/yyyy")
        For iJ = 2 To UBound(sRet, 2)
            sIndex(1, iJ) = sRet(1, iJ)
            bStart = False
            For iI = UBound(sRet, 1) To 2 Step -1
                If bStart = False And IsNumeric(sRet(iI, iJ)) Then
                    sIndex(iI + 1, iJ) = 100
                    sIndex(iI, iJ) = CStr(CDbl(sIndex(iI + 1, iJ)) * (1 + CDbl(sRet(iI, iJ))))
                    bStart = True
                ElseIf bStart And IsNumeric(sRet(iI, iJ)) Then
                    If IsNumeric(sIndex(iI + 1, iJ)) = False Then Exit Function
                    sIndex(iI, iJ) = CStr(CDbl(sIndex(iI + 1, iJ)) * (1 + CDbl(sRet(iI, iJ))))
                End If
            Next iI
        Next iJ
    ConvertReturns2Index = True
End Function
Private Function CalculateStats(sIn() As String, sRet() As String, sStats() As String, iM%, Optional bComp As Boolean) As Boolean
    Dim iJ%
    Dim sIndex() As String
    Dim dYears() As Double
        If CalculateAllVolatility(sRet, sStats, dYears, iM) = False Then Exit Function
        If CalculateAnnualisedReturn(sRet, sIndex, sStats, iM) = False Then Exit Function
        If CalculateDownVolatility(sRet, sStats, iM) = False Then Exit Function
        If CalculateSkewness(sRet, sStats, iM) = False Then Exit Function
        If CalculateKurtosis(sRet, sStats, iM) = False Then Exit Function
        If CalculateMaxDD(sIndex, sStats, iM) = False Then Exit Function
        If CalculateSharpe(sStats, bComp) = False Then Exit Function
        If CalculateSortino(sStats, bComp) = False Then Exit Function
        If CalculateCalmar(sStats, bComp) = False Then Exit Function
        If CalculateAdjustedSharpe(sStats, bComp) = False Then Exit Function
        If AddRow2End(sStats) = False Then Exit Function
        If AddRow2End(sStats) = False Then Exit Function
        sStats(UBound(sStats, 1), 1) = "Years of data"
        For iJ = 1 To UBound(dYears)
            sStats(UBound(sStats, 1), iJ + 1) = Format(dYears(iJ), "0.0")
        Next iJ
    CalculateStats = True
End Function
Private Function CalculateAllVolatility(sRet() As String, sStats() As String, dYears() As Double, iM%) As Boolean
    Dim iI%, iJ%, iX%
    Dim dX() As Double
        If iM < 0 Then iX = UBound(sRet, 1) Else iX = iM + 1
        If iX > UBound(sRet, 1) Then iX = UBound(sRet, 1)
        ReDim sStats(2, UBound(sRet, 2))
        ReDim dYears(UBound(sRet, 2) - 1)
        sStats(2, 1) = "Annualised volatility"
        For iJ = 2 To UBound(sRet, 2)
            sStats(1, iJ) = sRet(1, iJ)
            ReDim dX(1)
            dX(1) = CDbl(sRet(2, iJ))
            For iI = 3 To iX
                If IsNumeric(sRet(iI, iJ)) Then
                    ReDim Preserve dX(UBound(dX) + 1)
                    dX(UBound(dX)) = CDbl(sRet(iI, iJ))
                    dYears(iJ - 1) = dYears(iJ - 1) + 1
                Else: Exit For
                End If
            Next iI
            sStats(UBound(sStats, 1), iJ) = CStr(Application.WorksheetFunction.StDev(dX) * Sqr(12))
            dYears(iJ - 1) = dYears(iJ - 1) / 12
        Next iJ
        iM = iX
    CalculateAllVolatility = True
End Function
Private Function CalculateDownVolatility(sRet() As String, sStats() As String, iM%) As Boolean
    Dim bX As Boolean
    Dim iI%, iJ%
    Dim dX() As Double
        If AddRow2End(sStats) = False Then Exit Function
        sStats(UBound(sStats, 1), 1) = "Downside volatility"
        For iJ = 2 To UBound(sRet, 2) - 2
            ReDim dX(1)
            bX = False
            For iI = 3 To iM
                If IsNumeric(sRet(iI, iJ)) Then
                    If CDbl(sRet(iI, iJ)) < 0 Then
                        If UBound(dX) > 1 Or dX(1) <> 0 Then ReDim Preserve dX(UBound(dX) + 1)
                        dX(UBound(dX)) = CDbl(sRet(iI, iJ))
                    End If
                Else: Exit For
                End If
            Next iI
            If UBound(dX) > 1 Then sStats(UBound(sStats, 1), iJ) = CStr(Application.WorksheetFunction.StDev(dX) * Sqr(12))
        Next iJ
    CalculateDownVolatility = True
End Function
Private Function CalculateSkewness(sRet() As String, sStats() As String, iM%) As Boolean
    Dim bX As Boolean
    Dim iI%, iJ%
    Dim dX() As Double
        If AddRow2End(sStats) = False Then Exit Function
        sStats(UBound(sStats, 1), 1) = "Skewness"
        For iJ = 2 To UBound(sRet, 2) - 2
            ReDim dX(1)
            bX = False
            For iI = 3 To iM
                If IsNumeric(sRet(iI, iJ)) Then
                    If UBound(dX) > 1 Or dX(1) <> 0 Then ReDim Preserve dX(UBound(dX) + 1)
                    dX(UBound(dX)) = CDbl(sRet(iI, iJ))
                Else: Exit For
                End If
            Next iI
            If UBound(dX) > 1 Then sStats(UBound(sStats, 1), iJ) = CStr(Application.WorksheetFunction.Skew(dX))
        Next iJ
    CalculateSkewness = True
End Function
Private Function CalculateKurtosis(sRet() As String, sStats() As String, iM%) As Boolean
    Dim iI%, iJ%
    Dim dX() As Double
        If AddRow2End(sStats) = False Then Exit Function
        sStats(UBound(sStats, 1), 1) = "Kurtosis"
        For iJ = 2 To UBound(sRet, 2) - 2
            ReDim dX(1)
            For iI = 3 To iM
                If IsNumeric(sRet(iI, iJ)) Then
                    If UBound(dX) > 1 Or dX(1) <> 0 Then ReDim Preserve dX(UBound(dX) + 1)
                    dX(UBound(dX)) = CDbl(sRet(iI, iJ))
                Else: Exit For
                End If
            Next iI
            If UBound(dX) > 1 Then sStats(UBound(sStats, 1), iJ) = CStr(Application.WorksheetFunction.Kurt(dX))
        Next iJ
    CalculateKurtosis = True
End Function
Private Function CalculateMaxDD(sIndex() As String, sStats() As String, iM%) As Boolean
    Dim iI%, iJ%
    Dim dMax#, dMaxDD#
        If AddRow2End(sStats) = False Then Exit Function
        sStats(UBound(sStats, 1), 1) = "Maximum drawdown (monthly)"
        For iJ = 2 To UBound(sIndex, 2) - 2
            dMax = -100
            dMaxDD = 100
            For iI = iM To 2 Step -1
                If IsNumeric(sIndex(iI, iJ)) Then
                    If CDbl(sIndex(iI, iJ)) > dMax Then dMax = CDbl(sIndex(iI, iJ))
                    If dMax <= 0 Then Exit Function
                    If CDbl(sIndex(iI, iJ)) / dMax - 1 < dMaxDD Then dMaxDD = CDbl(sIndex(iI, iJ)) / dMax - 1
                End If
            Next iI
            sStats(UBound(sStats, 1), iJ) = CStr(dMaxDD)
        Next iJ
    CalculateMaxDD = True
End Function
Private Function CalculateSharpe(sStats() As String, bComp As Boolean) As Boolean
    Dim iJ%
    Dim sCash$
        If bComp Then sCash = "Cash" Else sCash = "1 Month SONIA (£)"
        If sStats(1, UBound(sStats, 2) - 1) <> sCash Or IsNumeric(sStats(2, UBound(sStats, 2) - 1)) = False Then Exit Function
        If AddRow2End(sStats) = False Then Exit Function
        sStats(UBound(sStats, 1), 1) = "Sharpe ratio"
        For iJ = 2 To UBound(sStats, 2) - 2
            sStats(UBound(sStats, 1), iJ) = CStr((CDbl(sStats(3, iJ)) - CDbl(sStats(3, UBound(sStats, 2) - 1))) / CDbl(sStats(2, iJ)))
        Next iJ
    CalculateSharpe = True
End Function
Private Function CalculateSortino(sStats() As String, bComp As Boolean) As Boolean
    Dim iJ%
    Dim sCash$
        If bComp Then sCash = "Cash" Else sCash = "1 Month SONIA (£)"
        If sStats(1, UBound(sStats, 2) - 1) <> sCash Or IsNumeric(sStats(2, UBound(sStats, 2) - 1)) = False Then Exit Function
        If AddRow2End(sStats) = False Then Exit Function
        sStats(UBound(sStats, 1), 1) = "Sortino ratio"
        For iJ = 2 To UBound(sStats, 2) - 2
            sStats(UBound(sStats, 1), iJ) = CStr((CDbl(sStats(3, iJ)) - CDbl(sStats(3, UBound(sStats, 2) - 1))) / CDbl(sStats(4, iJ)))
        Next iJ
    CalculateSortino = True
End Function
Private Function CalculateCalmar(sStats() As String, bComp As Boolean) As Boolean
    Dim iJ%
    Dim sCash$
        If bComp Then sCash = "Cash" Else sCash = "1 Month SONIA (£)"
        If sStats(1, UBound(sStats, 2) - 1) <> sCash Or IsNumeric(sStats(2, UBound(sStats, 2) - 1)) = False Then Exit Function
        If AddRow2End(sStats) = False Then Exit Function
        sStats(UBound(sStats, 1), 1) = "Calmar ratio"
        For iJ = 2 To UBound(sStats, 2) - 2
            sStats(UBound(sStats, 1), iJ) = CStr(CDbl(sStats(3, iJ)) / (-1 * CDbl(sStats(7, iJ))))
        Next iJ
    CalculateCalmar = True
End Function
Private Function CalculateAdjustedSharpe(sStats() As String, bComp As Boolean) As Boolean
    Dim iJ%
    Dim sCash$
        If bComp Then sCash = "Cash" Else sCash = "1 Month SONIA (£)"
        If sStats(1, UBound(sStats, 2) - 1) <> sCash Or IsNumeric(sStats(2, UBound(sStats, 2) - 1)) = False Then Exit Function
        If AddRow2End(sStats) = False Then Exit Function
        sStats(UBound(sStats, 1), 1) = "Adjusted Sharpe ratio"
        For iJ = 2 To UBound(sStats, 2) - 2
            sStats(UBound(sStats, 1), iJ) = CStr(CDbl(sStats(8, iJ)) * (1 + (CDbl(sStats(5, iJ)) / 6 * CDbl(sStats(8, iJ))) - (CDbl(sStats(6, iJ)) / 24 * CDbl(sStats(8, iJ)) ^ 2)))
        Next iJ
    CalculateAdjustedSharpe = True
End Function
Private Function OutputData(sS() As String, sRet() As String, sCorr() As String, sVol() As String, sCovar() As String) As Boolean
    Dim iI%
    Dim sX$
    Dim wsOut As Worksheet
        For iI = 3 To UBound(sS)
            If CheckSheet(ThisWorkbook.Name, sS(iI), wsOut) = False Then Exit Function
            Select Case iI
                Case 3: If StringArrayOut(wsOut, sCorr) = False Then Exit Function
                Case 4: If StringArrayOut(wsOut, sCovar) = False Then Exit Function
                Case 5: If StringArrayOut(wsOut, sVol) = False Then Exit Function
            End Select
        Next iI
    OutputData = True
End Function
Private Function OutputStats(sStats() As String, sCompStats() As String) As Boolean
    Dim wsOut As Worksheet
    Const sSOut$ = "RiskReturnStats"
        If AddRow2End(sStats) = False Then Exit Function
        If AppendString2String(sStats, sCompStats) = False Then Exit Function
        If CheckSheet(ThisWorkbook.Name, sSOut, wsOut) = False Then Exit Function
        If StringArrayOut(wsOut, sStats) = False Then Exit Function
    OutputStats = True
End Function
