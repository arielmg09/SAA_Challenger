Option Explicit
Option Base 1


Public Function MainReturnAssumptions(sVolRet4RA() As String, iRA%, iWght() As Integer) As Boolean
    Dim iX%
    Dim sSum4RA() As String
    Dim wsOut As Worksheet
    Const sSOut$ = "ReturnAssumptions"
    
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        
        If LoadInSummary4RA(sSum4RA) = False Then GoTo MyErrTrap
        If CheckTotal4RA(sSum4RA) = False Then GoTo MyErrTrap
        If CheckAssetOrder(sVolRet4RA, sSum4RA) = False Then GoTo MyErrTrap
        If AddInVol4RA(sVolRet4RA, sSum4RA) = False Then GoTo MyErrTrap
        If GenerateReturnAssumptionsFromVol(sSum4RA, iRA, iWght) = False Then GoTo MyErrTrap
        If BuildCompositeReturnAssumptions(sSum4RA) = False Then GoTo MyErrTrap
    
        If CheckSheet(ThisWorkbook.Name, sSOut, wsOut) = False Then GoTo MyErrTrap
        If StringArrayOut(wsOut, sSum4RA) = False Then GoTo MyErrTrap
                
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        
    MainReturnAssumptions = True
    Exit Function
MyErrTrap:
    iX = MsgBox("Error building return assumptions", vbExclamation)
End Function
Private Function LoadInSummary4RA(sSum4RA() As String) As Boolean
    Const sFIn$ = "Forward Looking Return Assumptions.xlsx", sPIn$ = "\\shbmain\data\Covent Garden\Resources\HIM\Strategic Asset Allocation", sS$ = "Summary"
        If FileOpen(sFIn, sPIn) = False Then Exit Function
        If Load2StringViaVariant(sFIn, sS, sSum4RA) = False Then Exit Function
        If FileClose(sFIn) = False Then Exit Function
        If TransposeString(sSum4RA) = False Then Exit Function
        ReDim Preserve sSum4RA(UBound(sSum4RA, 1), 10)
        If TransposeString(sSum4RA) = False Then Exit Function
    LoadInSummary4RA = True
End Function
Private Function CheckTotal4RA(sSum4RA() As String) As Boolean
    Dim iI%, iJ%, iMsg%
    Dim dX#
        For iJ = 2 To UBound(sSum4RA, 2)
            If IsNumeric(sSum4RA(UBound(sSum4RA, 1), iJ)) Then
                dX = 0
                For iI = 2 To UBound(sSum4RA, 1) - 1
                    If IsNumeric(sSum4RA(iI, iJ)) Then dX = dX + CDbl(sSum4RA(iI, iJ))
                Next iI
                If Round(dX, 5) <> Round(CDbl(sSum4RA(UBound(sSum4RA, 1), iJ)), 5) Then iMsg = MsgBox("Sum for column " & CStr(iJ) & " does not match summary total", vbExclamation) Else sSum4RA(UBound(sSum4RA, 1), iJ) = CStr(Round(dX, 6))
            End If
        Next iJ
    CheckTotal4RA = True
End Function
Private Function CheckAssetOrder(sVolRet4RA() As String, sSum4RA() As String) As Boolean
    Dim iI%, iJ%, iX%
    Dim sX() As String
        For iJ = 2 To UBound(sVolRet4RA, 2)
            iX = 0
            If FindColumnPositionInStringArray(sVolRet4RA(1, iJ), sSum4RA, iX) = False Or iX = 0 Then GoTo MyErrTrap
            If sVolRet4RA(1, iJ) <> sSum4RA(1, iJ) Then Exit For
        Next iJ
        If iJ <= UBound(sVolRet4RA, 2) Then
            ReDim sX(UBound(sSum4RA, 1), UBound(sVolRet4RA, 2) + 2)
            For iJ = 1 To UBound(sVolRet4RA, 2)
                iX = 0
                If iJ = 1 Then
                    iX = iJ
                Else
                    If FindColumnPositionInStringArray(sVolRet4RA(1, iJ), sSum4RA, iX) = False Or iX = 0 Then GoTo MyErrTrap
                End If
                For iI = 1 To UBound(sSum4RA, 1)
                    sX(iI, iJ) = sSum4RA(iI, iX)
                Next iI
            Next iJ
            ReDim sSum4RA(UBound(sX, 1), UBound(sX, 2))
            sSum4RA = sX
        Else
            If AddRow2End(sSum4RA, 3) = False Then Exit Function
        End If
    CheckAssetOrder = True
    Exit Function
MyErrTrap:
    MsgBox "Column headings in summary file do not asset class headings"
End Function
Private Function AddInVol4RA(sVolRet4RA() As String, sSum4RA() As String) As Boolean
    Dim iJ%
        sSum4RA(UBound(sSum4RA, 1) - 1, 1) = "Historic annualised volatility (" & sVolRet4RA(1, 1) & ")"
        sSum4RA(UBound(sSum4RA, 1), 1) = "Historic annualised return (" & sVolRet4RA(1, 1) & ")"
        For iJ = 2 To UBound(sVolRet4RA, 2)
            sSum4RA(UBound(sSum4RA, 1) - 1, iJ) = sVolRet4RA(UBound(sVolRet4RA, 1) - 1, iJ)
            sSum4RA(UBound(sSum4RA, 1), iJ) = sVolRet4RA(UBound(sVolRet4RA, 1), iJ)
        Next iJ
    AddInVol4RA = True
End Function
Private Function GenerateReturnAssumptionsFromVol(sSum4RA() As String, iRA%, iWght() As Integer) As Boolean
    Dim iJ%
    Dim dRegress() As Double
        If RunRegression(sSum4RA, dRegress, 2) = False Then Exit Function
        If RunRegression(sSum4RA, dRegress, 3) = False Then Exit Function
        If AddRow2End(sSum4RA, 3) = False Then Exit Function
        sSum4RA(UBound(sSum4RA, 1), 1) = "Intercept"
        sSum4RA(UBound(sSum4RA, 1) - 1, 1) = "Slope"
        sSum4RA(UBound(sSum4RA, 1) - 2, 2) = "HV with BBR"
        sSum4RA(UBound(sSum4RA, 1) - 2, 3) = "HV with HR"
        sSum4RA(UBound(sSum4RA, 1) - 2, 4) = "50:50 HV with BBR:HV with HR"
        If ApplyRegression2GetReturns(sSum4RA, dRegress) = False Then Exit Function
        If ApplyWeights2ReturnAssumptions(sSum4RA, iRA, iWght) = False Then Exit Function
        For iJ = 1 To UBound(sSum4RA, 2)
            sSum4RA(UBound(sSum4RA, 1) - 4, iJ) = sSum4RA(UBound(sSum4RA, 1) - 4 + (iRA - 1), iJ)
        Next iJ
        sSum4RA(UBound(sSum4RA, 1) - 4, UBound(sSum4RA, 2)) = sSum4RA(UBound(sSum4RA, 1) - 4, UBound(sSum4RA, 2) - 1) 'set inflation return assumption to equal cash return assumption
        If RemoveNonBlankRowFromString(sSum4RA, UBound(sSum4RA, 1) - 3, 2) = False Then Exit Function ' Have kept the data routines in, so can be re-included if desired!!!
        Select Case iRA
            Case 2
                sSum4RA(UBound(sSum4RA, 1) - 3, 2) = sSum4RA(UBound(sSum4RA, 1) - 3, 3)
                sSum4RA(UBound(sSum4RA, 1) - 4, 2) = sSum4RA(UBound(sSum4RA, 1) - 4, 3)
                sSum4RA(UBound(sSum4RA, 1) - 5, 2) = sSum4RA(UBound(sSum4RA, 1) - 5, 3)
            Case 3
                sSum4RA(UBound(sSum4RA, 1) - 3, 2) = sSum4RA(UBound(sSum4RA, 1) - 3, 4)
                sSum4RA(UBound(sSum4RA, 1) - 4, 2) = sSum4RA(UBound(sSum4RA, 1) - 4, 4)
                sSum4RA(UBound(sSum4RA, 1) - 5, 2) = sSum4RA(UBound(sSum4RA, 1) - 5, 4)
        End Select
        sSum4RA(UBound(sSum4RA, 1) - 3, 3) = ""
        sSum4RA(UBound(sSum4RA, 1) - 4, 3) = ""
        sSum4RA(UBound(sSum4RA, 1) - 5, 3) = ""
        sSum4RA(UBound(sSum4RA, 1) - 3, 4) = ""
        sSum4RA(UBound(sSum4RA, 1) - 4, 4) = ""
        sSum4RA(UBound(sSum4RA, 1) - 5, 4) = ""
        sSum4RA(UBound(sSum4RA, 1), UBound(sSum4RA, 2)) = sSum4RA(UBound(sSum4RA, 1), UBound(sSum4RA, 2) - 1) 'set inflation return assumption to equal cash return assumption
    GenerateReturnAssumptionsFromVol = True
End Function
Private Function RunRegression(sSum4RA() As String, dRegress() As Double, iRA%) As Boolean
    Dim dVol() As Double, dRet() As Double
        Select Case iRA
            Case 1: If LoadMaxMin2Doubles(sSum4RA, dVol, dRet) = False Then Exit Function 'use lowest and highest values from Building Block Return Assumptions and historic vol
            Case 2: If LoadAllBB2Doubles(sSum4RA, dVol, dRet) = False Then Exit Function 'use all datapoints from Building Block Return Assumptions and historic vol
            Case 3: If LoadAllHist2Doubles(sSum4RA, dVol, dRet) = False Then Exit Function 'use all datapoints from historic vol and historic return
        End Select
        If CheckDimDbl(dRegress) = False Then ReDim dRegress(2, 1) Else ReDim Preserve dRegress(2, 2)
        dRegress(1, UBound(dRegress, 2)) = Application.WorksheetFunction.Intercept(dRet, dVol)
        dRegress(2, UBound(dRegress, 2)) = Application.WorksheetFunction.Slope(dRet, dVol)
    RunRegression = True
End Function
Private Function LoadMaxMin2Doubles(sSum4RA() As String, dVol() As Double, dRet() As Double) As Boolean
    Dim iJ%
        ReDim dVol(2)
        ReDim dRet(2)
            dVol(1) = 100
            dRet(1) = 100
            For iJ = 2 To UBound(sSum4RA, 2) - 1 '*** have excluded inflation but included cash
                If IsNumeric(sSum4RA(UBound(sSum4RA, 1) - 1, iJ)) = False Then Exit Function
                If IsNumeric(sSum4RA(UBound(sSum4RA, 1) - 3, iJ)) = False Then Exit Function
                If CDbl(sSum4RA(UBound(sSum4RA, 1) - 3, iJ)) < dRet(1) Then dRet(1) = CDbl(sSum4RA(UBound(sSum4RA, 1) - 3, iJ))
                If CDbl(sSum4RA(UBound(sSum4RA, 1) - 3, iJ)) > dRet(2) Then dRet(2) = CDbl(sSum4RA(UBound(sSum4RA, 1) - 3, iJ))
                If CDbl(sSum4RA(UBound(sSum4RA, 1) - 1, iJ)) < dVol(1) Then dVol(1) = CDbl(sSum4RA(UBound(sSum4RA, 1) - 1, iJ))
                If CDbl(sSum4RA(UBound(sSum4RA, 1) - 1, iJ)) > dVol(2) Then dVol(2) = CDbl(sSum4RA(UBound(sSum4RA, 1) - 1, iJ))
            Next iJ
            If dRet(1) = 100 Or dVol(1) = 100 Or dRet(2) = 0 Or dVol(2) = 0 Then Exit Function
    LoadMaxMin2Doubles = True
End Function
Private Function LoadAllBB2Doubles(sSum4RA() As String, dVol() As Double, dRet() As Double) As Boolean
    Dim iJ%
        ReDim dVol(UBound(sSum4RA, 2) - 2)
        ReDim dRet(UBound(sSum4RA, 2) - 2)
        For iJ = 2 To UBound(sSum4RA, 2) - 1 '*** have excluded inflation but included cash
            If IsNumeric(sSum4RA(UBound(sSum4RA, 1) - 1, iJ)) = False Or IsNumeric(sSum4RA(UBound(sSum4RA, 1) - 3, iJ)) = False Then Exit Function
            dVol(iJ - 1) = CDbl(sSum4RA(UBound(sSum4RA, 1) - 1, iJ))
            dRet(iJ - 1) = CDbl(sSum4RA(UBound(sSum4RA, 1) - 3, iJ))
        Next iJ
    LoadAllBB2Doubles = True
End Function
Private Function LoadAllHist2Doubles(sSum4RA() As String, dVol() As Double, dRet() As Double) As Boolean
    Dim iJ%
        ReDim dVol(UBound(sSum4RA, 2) - 2)
        ReDim dRet(UBound(sSum4RA, 2) - 2)
        For iJ = 2 To UBound(sSum4RA, 2) - 1 '*** have excluded inflation but included cash
            If IsNumeric(sSum4RA(UBound(sSum4RA, 1) - 1, iJ)) = False Or IsNumeric(sSum4RA(UBound(sSum4RA, 1), iJ)) = False Then Exit Function
            dVol(iJ - 1) = CDbl(sSum4RA(UBound(sSum4RA, 1) - 1, iJ))
            dRet(iJ - 1) = CDbl(sSum4RA(UBound(sSum4RA, 1), iJ))
        Next iJ
    LoadAllHist2Doubles = True
End Function
Private Function ApplyRegression2GetReturns(sSum4RA() As String, dRegress() As Double) As Boolean
    Dim iI%, iJ%, iRow%
        If InStr(1, sSum4RA(UBound(sSum4RA, 1) - 4, 1), "Historic annualised volatility") = 0 Then Exit Function Else iRow = UBound(sSum4RA, 1) - 4
        ReDim Preserve dRegress(2, 3)
        dRegress(1, 3) = (dRegress(1, 1) + dRegress(1, 2)) / 2
        dRegress(2, 3) = (dRegress(2, 1) + dRegress(2, 2)) / 2
        If AddRow2End(sSum4RA, 3) = False Then Exit Function
        For iI = 1 To 3
            Select Case iI
                Case 1: sSum4RA(UBound(sSum4RA, 1) - (3 - iI), 1) = "Expected Returns from HV with BBR"
                Case 2: sSum4RA(UBound(sSum4RA, 1) - (3 - iI), 1) = "Expected Returns from HV with HR"
                Case 3: sSum4RA(UBound(sSum4RA, 1) - (3 - iI), 1) = "Expected Returns 50:50 BBR:HR (Total)"
            End Select
            sSum4RA(UBound(sSum4RA, 1) - 3, 1 + iI) = CStr(dRegress(1, iI))
            sSum4RA(UBound(sSum4RA, 1) - 4, 1 + iI) = CStr(dRegress(2, iI))
            For iJ = 2 To UBound(sSum4RA, 2)
                If IsNumeric(sSum4RA(iRow, iJ)) = False Then Exit Function
                sSum4RA(UBound(sSum4RA, 1) - (3 - iI), iJ) = CStr(Round(dRegress(2, iI) * CDbl(sSum4RA(iRow, iJ)) + dRegress(1, iI), 6))
            Next iJ
        Next iI
    ApplyRegression2GetReturns = True
End Function
Private Function ApplyWeights2ReturnAssumptions(sSum4RA() As String, iRA%, iWght() As Integer) As Boolean
    Dim iJ%, iRow%
        If sSum4RA(UBound(sSum4RA, 1) - 9, 1) <> "Expected Return BBA (Total)" Then
            MsgBox "Can't locate row 'Expected Return BBA (Total)'"
            Exit Function
        Else: iRow = UBound(sSum4RA, 1) - 9
        End If
        If AddRow2End(sSum4RA, 2) = False Then Exit Function
        sSum4RA(UBound(sSum4RA, 1) - 1, 1) = "Weight of Building Block Return"
        Select Case iRA
            Case 1: sSum4RA(UBound(sSum4RA, 1), 1) = "Weighted Expected Returns (HV/BBR)"
            Case 2: sSum4RA(UBound(sSum4RA, 1), 1) = "Weighted Expected Returns (HV/HR)"
            Case 3: sSum4RA(UBound(sSum4RA, 1), 1) = "Weighted Expected Returns (50:50 HV/BBR:HV/HR)"
        End Select
        For iJ = 2 To UBound(sSum4RA, 2) - 1
            If IsNumeric(sSum4RA(iRow, iJ)) = False Or IsNumeric(sSum4RA(iRow + 6 + iRA, iJ)) = False Then Exit Function
            sSum4RA(UBound(sSum4RA, 1), iJ) = CDbl(sSum4RA(iRow, iJ)) * (CDbl(iWght(iJ - 1)) / 100) + CDbl(sSum4RA(iRow + 6 + iRA, iJ)) * (CDbl(100 - iWght(iJ - 1)) / 100)
            sSum4RA(UBound(sSum4RA, 1) - 1, iJ) = Format(CDbl(iWght(iJ - 1)) / 100, "0%")
        Next iJ
    ApplyWeights2ReturnAssumptions = True
End Function
Private Function RemoveNonBlankRowFromString(sX() As String, iX%, Optional iY%) As Boolean
    Dim iI%, iJ%
    Dim sZ() As String
        If iY < 0 Or iY >= UBound(sX, 1) Then Exit Function
        If iY <> 0 Then iY = iY - 1
        For iI = 1 To UBound(sX, 1)
            If iI < iX Or iI > iX + iY Then
                If CheckDimStr(sZ) = False Then ReDim sZ(UBound(sX, 2), 1) Else ReDim Preserve sZ(UBound(sZ, 1), UBound(sZ, 2) + 1)
                For iJ = 1 To UBound(sX, 2)
                    sZ(iJ, UBound(sZ, 2)) = sX(iI, iJ)
                Next iJ
            End If
        Next iI
        If CheckDimStr(sZ) Then
            If TransposeString(sZ) = False Then Exit Function
            ReDim sX(UBound(sZ, 1), UBound(sZ, 2))
            sX = sZ
        Else:  Exit Function
        End If
    RemoveNonBlankRowFromString = True
End Function
Private Function BuildCompositeReturnAssumptions(sSum4RA() As String) As Boolean 'bring in sCompWgt so use floated weights, as in how it is done for yield?
    Dim iI%, iRowNum%
    Dim sACM() As String, sCompRA() As String
    Const sSACM$ = "AssetClassMap", iWgt% = 3
        If Load2StringViaVariant(ThisWorkbook.Name, sSACM, sACM) = False Then Exit Function
        For iI = 1 To UBound(sACM, 1)
            If sACM(iI, iWgt) = 1 Then
                If CheckDimStr(sCompRA) = False Then ReDim sCompRA(2, 1) Else: ReDim Preserve sCompRA(UBound(sCompRA, 1), UBound(sCompRA, 2) + 1)
                sCompRA(1, UBound(sCompRA, 2)) = sACM(iI, iWgt + 1)
                If sACM(iI, 1) = sSum4RA(1, iI + 1) Then sCompRA(2, UBound(sCompRA, 2)) = sSum4RA(UBound(sSum4RA, 1), iI + 1) Else Exit Function
            Else
                If FindNumOfRows4Composite(sACM, iRowNum, iWgt, iI) = False Then Exit Function
                If CheckDimStr(sCompRA) = False Then ReDim sCompRA(2, 1) Else: ReDim Preserve sCompRA(UBound(sCompRA, 1), UBound(sCompRA, 2) + 1)
                If sACM(iI, 1) <> sSum4RA(1, iI + 1) Then Exit Function
                sCompRA(1, UBound(sCompRA, 2)) = sACM(iI, iWgt + 1)
                If AggregateCompositeReturnAssumptions(sSum4RA, sCompRA, sACM, iRowNum, iI, iWgt) = False Then Exit Function
            End If
        Next iI
        If TransposeString(sCompRA) = False Then Exit Function
        If AddRow2End(sSum4RA) = False Then Exit Function
        If AppendString2String(sSum4RA, sCompRA) = False Then Exit Function
    BuildCompositeReturnAssumptions = True
End Function
Private Function AggregateCompositeReturnAssumptions(sSum4RA() As String, sCompRA() As String, sACM() As String, iRowNum%, iX%, iWgt%) As Boolean
    Dim iI%
    Dim dX#, dZ#
        For iI = iX To iX + (iRowNum - 1)
            If CDbl(sSum4RA(UBound(sSum4RA, 1), iI + 1)) >= 0 Then
                dX = dX + CDbl(sSum4RA(UBound(sSum4RA, 1), iI + 1)) * CDbl(sACM(iI, iWgt))
                dZ = dZ + CDbl(sACM(iI, iWgt))
            End If
        Next iI
        If dZ <> 0 Then sCompRA(2, UBound(sCompRA, 2)) = CStr(dX / dZ)
        iX = iX + (iRowNum - 1)
    AggregateCompositeReturnAssumptions = True
End Function

