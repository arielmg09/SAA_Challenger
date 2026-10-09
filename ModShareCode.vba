Option Explicit
Option Base 1

Public Function GetDrive(sX$, sZ$) As Boolean
' GetDrive - gets drive letter from filename sX
        On Error Resume Next
        sX = Workbooks(sX).Path
        If Err <> 0 Or InStr(1, sX, ":") = 0 Then Exit Function Else sZ = Mid(sX, 1, InStr(1, sX, ":"))
    GetDrive = True
End Function
Public Function FExist(sX$) As Boolean
' FExist - returns TRUE if File exists in the active path
    Dim sZ$
        On Error Resume Next
        sZ = Dir(sX)
        If Err = 0 And sZ <> "" Then FExist = True
End Function
Public Function FileOpen(sF$, sPath$, Optional bNotRO As Boolean) As Boolean
' open file sF from path sPath
    Dim sAns$
    Dim bRO As Boolean
        If bNotRO = False Then bRO = True
        sF = sPath & "\" & sF ' give the file it's full name
        If FExist(sF) = False Then Exit Function
        Workbooks.Open filename:=sF, UpdateLinks:=0, ReadOnly:=bRO ' open the data file
        FileOpen = True
End Function
Public Function FileClose(sF$, Optional bNotRO As Boolean) As Boolean
        On Error GoTo ErrTrap
        sF = FileNameOnly(sF) ' get the short file name from the full name
        If bNotRO = True Then Workbooks(sF).Save
        Workbooks(sF).Saved = True
        Workbooks(sF).Close
        FileClose = True
ErrTrap:
        Exit Function
End Function
Public Function CheckSheet(ByVal sF As String, sX$, wsX As Worksheet, Optional bX As Boolean) As Boolean
' CheckSheet - check sX exists in file sF and set as wsX. If bx true supress error message
    Dim sAns$
        sF = FileNameOnly(sF) ' get the short file name from the full name
        If SExist(sF, sX) = False Then
            If bX = False Then sAns = MsgBox("Sheet '" + sX + "' does not exist", vbExclamation)
            Exit Function
        Else: Set wsX = Workbooks(sF).Sheets(sX)
        End If
    CheckSheet = True
End Function
Public Function FileNameOnly(ByVal sPathAndName$) As String
'FileNameOnly - returns the filename from a path/filename string
        If InStr(1, sPathAndName, "\") = 0 Then
            If InStr(1, sPathAndName, ".") <> 0 Then FileNameOnly = Mid(sPathAndName, 1, InStr(1, sPathAndName, ".") - 1)
            Exit Function
        End If
        If FExist(sPathAndName) = False Then Exit Function
        If InStr(1, Dir(sPathAndName), ".") = 0 Then FileNameOnly = Dir(sPathAndName) Else FileNameOnly = Mid(Dir(sPathAndName), 1, InStr(1, Dir(sPathAndName), ".") - 1)
End Function
Private Function SExist(sFile$, sX$) As Boolean
' SExist -  TRUE if sheet exists in the this workbook
    Dim oX As Object
        On Error Resume Next
        Set oX = Workbooks(sFile).Sheets(sX)
        If Err = 0 Then SExist = True
End Function
Public Function IsStringInString(s1$, s2$) As Boolean
' IsStringInString - if s1 is found in s2 then IsStringInString = TRUE
    Dim iX%
        iX = -1
        On Error Resume Next
        If Len(s2) - Len(s1) >= 0 Then
            iX = InStr(1, s2, s1)
            If Err = 0 And iX > 0 Then IsStringInString = True
        End If
End Function
Public Function CopyDataSheet2Sheet(ByVal sFIn$, ByVal sFOut$, sSIn$, Optional bInOut As Boolean) As Boolean
' copy data from worksheet wsIn in file sFIn to sheet sSIn in this file
    Dim sX$
    Dim lR&, lC&
    Dim wsIn As Worksheet, wsOut As Worksheet
    Dim vIn() As Variant
        If bInOut Then
            sX = sFIn
            sFIn = sFOut
            sFOut = sX
        End If
        If CheckSheet(sFIn, sSIn, wsIn) = False Then Exit Function
        If CheckSheet(sFOut, sSIn, wsOut) = False Then Exit Function
        If LastRow(wsIn, lR) = False Then Exit Function
        If LastColumn(wsIn, lC) = False Then Exit Function
        wsOut.Cells.ClearContents
        Range(wsOut.Cells(1, 1), wsOut.Cells(lR, lC)).Value = Range(wsIn.Cells(1, 1), wsIn.Cells(lR, lC)).Value
        Call AdjustSedol(wsOut, "SEDOL")
    CopyDataSheet2Sheet = True
End Function
Public Function LastRow(wsIn As Worksheet, lLastRow&) As Boolean
        lLastRow = 1
        If WorksheetFunction.CountA(wsIn.Cells) > 0 Then 'Search for any entry, by searching backwards by Rows.
            lLastRow = wsIn.Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        End If
    LastRow = True
End Function
Public Function LastColumn(wsIn As Worksheet, lLastColumn&) As Boolean
        lLastColumn = 1
        If WorksheetFunction.CountA(wsIn.Cells) > 0 Then 'Search for any entry, by searching backwards by Columns.
            lLastColumn = wsIn.Cells.Find(What:="*", After:=[A1], SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        End If
    LastColumn = True
End Function
Public Function Load2StringViaVariant(sFIn$, sSIn$, sIn() As String) As Boolean
    Dim lI&, lJ&, lC&, lR&
    Dim wsIn As Worksheet
    Dim vIn() As Variant
        If CheckSheet(sFIn, sSIn, wsIn) = False Then Exit Function
        If LastRow(wsIn, lR) = False Then Exit Function
        If LastColumn(wsIn, lC) = False Then Exit Function
        ReDim vIn(lR, lC)
        ReDim sIn(lR, lC)
        vIn = Range(wsIn.Cells(1, 1), wsIn.Cells(lR, lC))
        For lI = 1 To lR
            For lJ = 1 To lC
                sIn(lI, lJ) = CStr(vIn(lI, lJ))
            Next lJ
        Next lI
    Load2StringViaVariant = True
End Function
Public Function FindColumnPositionInStringArray(sX$, sZ() As String, iZ%, Optional iX%) As Boolean
    Dim iI%
        If CheckDimStr(sZ, 2) = False Then Exit Function
        If iX = 0 Then iX = 1
        For iI = 1 To UBound(sZ, 2)
            If sX = sZ(iX, iI) Then
                iZ = iI
                Exit For
            End If
        Next iI
    FindColumnPositionInStringArray = True
End Function
Public Function FindRowPositionInStringArray(sX$, sZ() As String, iZ%, Optional iX%) As Boolean
    Dim iI%
        If CheckDimStr(sZ, 2) = False Then Exit Function
        If iX = 0 Then iX = 1
        For iI = 1 To UBound(sZ, 1)
            If sX = sZ(iI, iX) Then
                iZ = iI
                Exit For
            End If
        Next iI
    FindRowPositionInStringArray = True
End Function
Public Function CheckDimBool(bX() As Boolean, Optional iZ%) As Boolean
    Dim iX%
    On Error Resume Next
        If iZ = 0 Then iZ = 1
        iX = UBound(bX, iZ)
        If Err = 0 Then CheckDimBool = True
End Function
Public Function CheckDimDbl(dX() As Double, Optional iZ%) As Boolean
    Dim iX%
    On Error Resume Next
        If iZ = 0 Then iZ = 1
        iX = UBound(dX, iZ)
        If Err = 0 Then CheckDimDbl = True
End Function
Public Function CheckDimStr(sX() As String, Optional iZ%) As Boolean
    Dim iX%
    On Error Resume Next
        If iZ = 0 Then iZ = 1
        iX = UBound(sX, iZ)
        If Err = 0 Then CheckDimStr = True
End Function
Public Function CheckDimInt(iX() As Integer, Optional iZ%) As Boolean
    Dim iY%
    On Error Resume Next
        If iZ = 0 Then iZ = 1
        iY = UBound(iX, iZ)
        If Err = 0 Then CheckDimInt = True
End Function
Public Function CheckDimDate(dtX() As Date, Optional iZ%) As Boolean
    Dim iX%
    On Error Resume Next
        If iZ = 0 Then iZ = 1
        iX = UBound(dtX, iZ)
        If Err = 0 Then CheckDimDate = True
End Function
Public Function TransposeString(sX() As String) As Boolean
    Dim iI%, iJ%
    Dim sZ() As String
        ReDim sZ(UBound(sX, 2), UBound(sX, 1))
        For iI = 1 To UBound(sX, 1)
            For iJ = 1 To UBound(sX, 2)
                sZ(iJ, iI) = sX(iI, iJ)
            Next iJ
        Next iI
        ReDim sX(UBound(sZ, 2), UBound(sZ, 1))
        sX = sZ
    TransposeString = True
End Function
Public Function RemoveFirstRowInString(sX() As String) As Boolean
    Dim iI%, iJ%
    Dim sZ() As String
        ReDim sZ(UBound(sX, 1) - 1, UBound(sX, 2))
        For iI = 2 To UBound(sX, 1)
            For iJ = 1 To UBound(sX, 2)
               sZ(iI - 1, iJ) = sX(iI, iJ)
            Next iJ
        Next iI
        ReDim sX(UBound(sZ, 1), UBound(sZ, 2))
        sX = sZ
    RemoveFirstRowInString = True
End Function
Public Sub ReSizeString(sX() As String, iX%) 'required by Add2String
' increase 1st dimension of sX to iX, retaining data
    Dim iI%, iJ%
    Dim sY() As String
        ReDim sY(iX, UBound(sX, 2))
        For iI = 1 To UBound(sX, 1)
            For iJ = 1 To UBound(sX, 2)
                sY(iI, iJ) = sX(iI, iJ)
            Next iJ
        Next iI
        ReDim sX(UBound(sY, 1), UBound(sY, 2))
        sX = sY
End Sub
Public Function ResizeRowsInStringArray(sX() As String, iX%) As Boolean
        If TransposeString(sX) = False Then Exit Function
        ReDim Preserve sX(UBound(sX, 1), iX)
        If TransposeString(sX) = False Then Exit Function
    ResizeRowsInStringArray = True
End Function
Public Function AddColumnHeadings(sX() As String, sZ() As String)
    Dim iI%
        If AddRow2Start(sX) = False Then Exit Function
        If CheckDimStr(sZ) = False Then Exit Function
        If UBound(sX, 2) <> UBound(sZ) Then Exit Function
        For iI = 1 To UBound(sX, 2)
            sX(1, iI) = sZ(iI)
        Next iI
    AddColumnHeadings = True
End Function
Public Function AddRow2Start(sX() As String, Optional iX%) As Boolean
    Dim iI%, iJ%
    Dim sZ() As String
        If CheckDimStr(sX, 2) = False Then Exit Function
        If iX < 1 Then iX = 1
        ReDim sZ(UBound(sX, 1) + iX, UBound(sX, 2))
        For iI = 1 To UBound(sX, 1)
            For iJ = 1 To UBound(sX, 2)
                sZ(iI + iX, iJ) = sX(iI, iJ)
            Next iJ
        Next iI
        ReDim sX(UBound(sZ, 1), UBound(sZ, 2))
        sX = sZ
    AddRow2Start = True
End Function
Public Function AddRow2End(sX() As String, Optional iX%) As Boolean
        If CheckDimStr(sX, 2) = False Then Exit Function
        If iX < 1 Then iX = 1
        If TransposeString(sX) = False Then Exit Function
        ReDim Preserve sX(UBound(sX, 1), UBound(sX, 2) + iX)
        If TransposeString(sX) = False Then Exit Function
    AddRow2End = True
End Function
Public Function AddColumn2Start(sX() As String, Optional iX%) As Boolean
    Dim iI%, iJ%
    Dim sZ() As String
        If CheckDimStr(sX, 2) = False Then Exit Function
        If iX < 1 Then iX = 1
        ReDim sZ(UBound(sX, 1), UBound(sX, 2) + iX)
        For iI = 1 To UBound(sX, 1)
            For iJ = 1 To UBound(sX, 2)
                sZ(iI, iJ + iX) = sX(iI, iJ)
            Next iJ
        Next iI
        ReDim sX(UBound(sZ, 1), UBound(sZ, 2))
        sX = sZ
    AddColumn2Start = True
End Function
Public Function StringArrayOut(wsOut As Worksheet, sOut() As String, Optional bDate As Boolean) As Boolean
    Dim iI%, iDate%
    Dim lI&, lJ&
    Dim vX() As Variant
        If bDate Then
            For iI = 1 To UBound(sOut, 2)
                If UCase(sOut(1, iI)) = "DATE" Then
                    iDate = iI
                    Exit For
                End If
            Next iI
            If iDate = 0 Then bDate = False
        End If
        ReDim vX(UBound(sOut, 1), UBound(sOut, 2))
        For lI = 1 To UBound(sOut, 1)
            For lJ = 1 To UBound(sOut, 2)
                If IsNumeric(sOut(lI, lJ)) Then
                    vX(lI, lJ) = CDbl(sOut(lI, lJ))
                ElseIf sOut(lI, lJ) <> "" Then
                    vX(lI, lJ) = sOut(lI, lJ)
                End If
                If bDate And iDate = CInt(lJ) And IsDate(sOut(lI, lJ)) Then vX(lI, lJ) = CDate(sOut(lI, lJ))
            Next lJ
        Next lI
        Call VariantOut(wsOut, vX)
        Call AdjustSedol(wsOut, "Sedol")
    StringArrayOut = True
End Function
Public Sub AdjustSedol(wsOut As Worksheet, sX$, Optional iStart%)
    Dim bUpdate As Boolean
    Dim lI&, lR&, lC&
        bUpdate = Application.ScreenUpdating
        If bUpdate Then
            Application.ScreenUpdating = False
            Application.Calculation = xlManual
        End If
        If iStart = 0 Then iStart = 1
        Call LastColumn(wsOut, lC)
        Call LastRow(wsOut, lR)
        For lI = 1 To lC
            If UCase(CStr(wsOut.Cells(iStart, lI))) = UCase(sX) Then Exit For
        Next lI
        If lI <= lC Then
            lC = lI
            Range(wsOut.Cells(1, lC), wsOut.Cells(lR, lC)).NumberFormat = "@"
            For lI = 2 To lR
                If wsOut.Cells(lI, lC) <> "" Then wsOut.Cells(lI, lC) = Application.WorksheetFunction.Text(wsOut.Cells(lI, lC), "0000000")
            Next lI
        Else
            If lR > 10 Then lR = 10
            For lI = 1 To lR
                If UCase(CStr(wsOut.Cells(lI, iStart))) = UCase(sX) Then Exit For
            Next lI
            If lI <= lR Then
                lR = lI
                Range(wsOut.Cells(lR, iStart), wsOut.Cells(lR, lC)).NumberFormat = "@"
                For lI = 2 To lC
                    If wsOut.Cells(lR, lI) <> "" Then wsOut.Cells(lR, lI) = Application.WorksheetFunction.Text(wsOut.Cells(lR, lI), "0000000")
                Next lI
            End If
        End If
        If bUpdate Then
            Application.ScreenUpdating = True
            Application.Calculation = xlAutomatic
        Else
            Application.ScreenUpdating = False
            Application.Calculation = xlManual
        End If
End Sub
Public Sub VariantOut(wsOut As Worksheet, vOut As Variant, Optional bAppend As Boolean)
'VariantOut - clear worksheet wsOut then write contents of vOut
    Dim lLastRow&
        If bAppend = False Then
            wsOut.Cells.ClearContents
            Range(wsOut.Cells(1, 1), wsOut.Cells(UBound(vOut, 1), UBound(vOut, 2))) = vOut
        Else
            If LastRow(wsOut, lLastRow) = False Then Exit Sub
            Range(wsOut.Cells(lLastRow + 1, 1), wsOut.Cells(lLastRow + UBound(vOut, 1), UBound(vOut, 2))) = vOut
        End If
End Sub
Public Function SpliceStrings(sX() As String, sY() As String) As Boolean
' splices two strings and overwrites sX with the result
    Dim iI%, iJ%
    Dim lR As Long, lC As Long, lZ As Long
    Dim bX As Boolean, bY As Boolean
    Dim sZ() As String
        bX = True
        bY = True
        If UBound(sX, 1) > UBound(sY, 1) Then 'calclate which variant has the largest 1st dimension
            lR = UBound(sX, 1)
        Else
            lR = UBound(sY, 1)
        End If
        If CheckDimStr(sX, 2) = False Then
            bX = False 'sX is one dimensional array
            If CheckDimStr(sY, 2) = False Then
                bY = False 'sY is one dimensional array
                lC = 2 'second dimension of sZ
            Else: lC = lZ + UBound(sY, 2) 'second dimension of sZ
            End If
        ElseIf CheckDimStr(sY, 2) = False Then
            bY = False 'sY is one dimensional array
            lC = UBound(sX, 2) + 1 'second dimension of sZ
        Else: lC = UBound(sX, 2) + UBound(sY, 2) '2nd dimension is the sum of the 2nd dimensions for each variant
        End If
        ReDim sZ(lR, lC)
        For iI = 1 To UBound(sX, 1) 'loop through sX and write to temp variant sZ
            If bX = True Then
                For iJ = 1 To UBound(sX, 2)
                    sZ(iI, iJ) = sX(iI, iJ)
                Next iJ
            Else: sZ(iI, 1) = sX(iI)
            End If
        Next iI
        If bX = True Then
            lC = UBound(sX, 2)
        Else: lC = 1
        End If
        For iI = 1 To UBound(sY, 1) 'loop through sY and write to temp variant sZ
            If bY = True Then
                For iJ = 1 To UBound(sY, 2)
                    sZ(iI, lC + iJ) = sY(iI, iJ)
                Next iJ
            Else: sZ(iI, lC + 1) = sY(iI)
            End If
        Next iI
        ReDim sX(UBound(sZ, 1), UBound(sZ, 2)) 'overwrite sX with temp variant sZ
        sX = sZ
        SpliceStrings = True
End Function
Public Sub SortStringArray(sX() As String, iX%, Optional bX As Boolean)
' sorts a two-dimensional array stored in a string array on column iX. If bX TRUE the sort is ascending, otherwise descending
    Dim iI%, iJ%, iZ%, iStart%
    Dim sTemp() As String, sTemp2() As String
        If IsNumeric(sX(1, iX)) = True Then
            iStart = 1
        Else: iStart = 2
        End If
        For iI = iStart To UBound(sX, 1) 'count number of empty entries
            If sX(iI, 1) <> "" And sX(iI, iX) <> "" Then iZ = iZ + 1
        Next iI
        ReDim sTemp(iZ, UBound(sX, 2)) 'dimension temp variant excluding empty entries
        iZ = 0
        For iI = iStart To UBound(sX, 1) 'copy non-empty entries into temp variant
            If sX(iI, 1) <> "" And sX(iI, iX) <> "" Then
                iZ = iZ + 1
                For iJ = 1 To UBound(sX, 2)
                    sTemp(iZ, iJ) = sX(iI, iJ)
                Next iJ
            End If
        Next iI
        Call TwoDStringSort(sTemp, iX, bX) 'sort temp variant
        ReDim sTemp2(UBound(sX, 2)) 'overwrite data in sX with sorted data
        If iStart = 2 Then
            For iJ = 1 To UBound(sX, 2)
                sTemp2(iJ) = sX(1, iJ)
            Next iJ
            ReDim sX(UBound(sTemp, 1) + 1, UBound(sTemp, 2))
            For iJ = 1 To UBound(sX, 2)
                sX(1, iJ) = sTemp2(iJ)
            Next iJ
        End If
        If iStart = 1 Then ReDim sX(UBound(sTemp, 1), UBound(sTemp, 2))
        For iI = iStart To UBound(sX, 1)
            For iJ = 1 To UBound(sX, 2)
                Select Case iStart
                    Case 1: sX(iI, iJ) = sTemp(iI, iJ)
                    Case 2: sX(iI, iJ) = sTemp(iI - 1, iJ)
                End Select
            Next iJ
        Next iI
End Sub
Private Sub TwoDStringSort(sArray() As String, iSortCol1%, Optional bAsc As Boolean)
' TwoDSort - Sort a 2 dimensional string array on column iSortCol1. If bAsc TRUE the sort is ascending, otherwise descending
    Dim bCond1 As Boolean
    Dim iI%, iJ%, iY%
    Dim sT$
        For iI = LBound(sArray, 1) To UBound(sArray, 1) - 1
            For iJ = LBound(sArray, 1) To UBound(sArray, 1) - 1
                If bAsc Then
                    bCond1 = CDbl(sArray(iJ, iSortCol1)) > CDbl(sArray(iJ + 1, iSortCol1))
                Else
                    bCond1 = CDbl(sArray(iJ, iSortCol1)) < CDbl(sArray(iJ + 1, iSortCol1))
                End If
                If bCond1 Then
                    For iY = LBound(sArray, 2) To UBound(sArray, 2)
                        sT = sArray(iJ, iY)
                        sArray(iJ, iY) = sArray(iJ + 1, iY)
                        sArray(iJ + 1, iY) = sT
                    Next iY
                End If
            Next iJ
        Next iI
End Sub
Public Function InsertNewSheet(ByVal sFileIn As String, sNewSheet$, wsNewSheet As Worksheet) As Boolean
' InsertNewSheet - inserts new sheet called sNewSheet into file sFile and sets as worksheet wsNewSheet
        sFileIn = FileNameOnly(sFileIn)
        Set wsNewSheet = Workbooks(sFileIn).Sheets.Add(After:=Workbooks(sFileIn).Worksheets(Workbooks(sFileIn).Worksheets.Count))
        wsNewSheet.Name = sNewSheet
    InsertNewSheet = True
End Function
Public Function RemoveBlankRows(sX() As String, sY$, Optional bRemoveFirstRC As Boolean) As Boolean
    Dim iI%, iJ%, iY%
    Dim sZ() As String
        If bRemoveFirstRC = False Then If RemoveBlankFirstRowColumn(sX) = False Then Exit Function
        If FindColumnPositionInStringArray(sY, sX, iY) = False Then Exit Function
        If iY = 0 Then GoTo MyContinue
        ReDim sZ(UBound(sX, 2), 1)
        For iI = 1 To UBound(sX, 1)
            If sX(iI, iY) <> "" Then
                For iJ = 1 To UBound(sX, 2)
                    sZ(iJ, UBound(sZ, 2)) = sX(iI, iJ)
                Next iJ
                ReDim Preserve sZ(UBound(sZ, 1), UBound(sZ, 2) + 1)
            End If
        Next iI
        If UBound(sZ, 2) > 1 Then ReDim Preserve sZ(UBound(sZ, 1), UBound(sZ, 2) - 1)
        If TransposeString(sZ) = False Then Exit Function
        ReDim sX(UBound(sZ, 1), UBound(sZ, 2))
        sX = sZ
MyContinue:
    RemoveBlankRows = True
End Function
Private Function RemoveBlankFirstRowColumn(sX() As String) As Boolean
    Dim iI%, iJ%, iStart%
    Dim sZ() As String
        For iI = 1 To UBound(sX, 2)
            If sX(1, iI) <> "" Then Exit For
        Next iI
        If iI > UBound(sX, 2) Then iStart = iStart + 1
        ReDim sZ(UBound(sX, 1) - iStart, UBound(sX, 2))
        For iI = iStart + 1 To UBound(sX, 1)
            For iJ = 1 To UBound(sX, 2)
                sZ(iI - iStart, iJ) = sX(iI, iJ)
            Next iJ
        Next iI
        ReDim sX(UBound(sZ, 1), UBound(sZ, 2))
        sX = sZ
        iStart = 0
        For iI = 1 To UBound(sX, 1)
            If sX(iI, 1) <> "" Then Exit For
        Next iI
        If iI > UBound(sX, 1) Then iStart = iStart + 1
        ReDim sZ(UBound(sX, 1), UBound(sX, 2) - iStart)
        For iI = 1 To UBound(sZ, 1)
            For iJ = iStart + 1 To UBound(sX, 2)
                sZ(iI, iJ - iStart) = sX(iI, iJ)
            Next iJ
        Next iI
        ReDim sX(UBound(sZ, 1), UBound(sZ, 2))
        sX = sZ
    RemoveBlankFirstRowColumn = True
End Function
Public Function AppendString2String(sX() As String, sY() As String) As Boolean
' append data from string sY to string sX
    Dim iI%, iJ%, iStartRow%
        If UBound(sX, 2) < UBound(sY, 2) Then ReDim Preserve sX(UBound(sX, 1), UBound(sY, 2))
        iStartRow = UBound(sX, 1)
        Call ReSizeString(sX, UBound(sX, 1) + UBound(sY, 1))
        For iI = 1 To UBound(sY, 1)
            For iJ = 1 To UBound(sY, 2)
                sX(iStartRow + iI, iJ) = sY(iI, iJ) 'add data into sX
            Next iJ
        Next iI
        AppendString2String = True
End Function
Public Function InsertBlankRow(sX() As String, iX%, Optional iY%) As Boolean
    Dim iI%, iJ%
    Dim sZ() As String
        If CheckDimStr(sX) = False Or CheckDimStr(sX, 2) = False Or CheckDimStr(sX, 3) = True Then Exit Function
        If iX < 1 Or iX >= UBound(sX, 1) Then Exit Function
        If iY < 1 Then iY = 1
        ReDim sZ(UBound(sX, 1) + iY, UBound(sX, 2))
        For iI = 1 To iX
            For iJ = 1 To UBound(sX, 2)
                sZ(iI, iJ) = sX(iI, iJ)
            Next iJ
        Next iI
        For iI = iX + 1 To UBound(sX, 1)
            For iJ = 1 To UBound(sX, 2)
                sZ(iI + iY, iJ) = sX(iI, iJ)
            Next iJ
        Next iI
        ReDim sX(UBound(sZ, 1), UBound(sZ, 2))
        sX = sZ
    InsertBlankRow = True
End Function
Public Function IsFileOpen(sFName As String) As Boolean
    Dim iFileNum%, iErrNum%
        On Error Resume Next   ' Turn error checking off.
        iFileNum = FreeFile()   ' Get a free file number.
        Open sFName For Input Lock Read As #iFileNum ' Attempt to open the file and lock it.
        Close iFileNum          ' Close the file.
        iErrNum = Err           ' Save the error number that occurred.
        On Error GoTo 0        ' Turn error checking back on.
        Select Case iErrNum ' Check to see which error occurred.
            Case 0: IsFileOpen = False ' No error occurred, file is NOT already open by another user.
            Case 70: IsFileOpen = True ' Error number for "Permission Denied", file is already opened by another user.
            Case Else: Error iErrNum ' Another error occurred.
       End Select
End Function
Public Function IsFileOpenByMe(sFName As String) As Boolean
    Dim iFileNum%, iErrNum%
    Dim wbkZ As Workbook
        
        On Error Resume Next   ' Turn error checking off.
        Set wbkZ = Workbooks(FileNameOnly(sFName))
        iErrNum = Err           ' Save the error number that occurred.
        Select Case iErrNum ' Check to see which error occurred.
            Case 0: IsFileOpenByMe = True ' No error occurred, file is already open by me.
            Case Else: IsFileOpenByMe = False ' Error number for "Subscript out-of-range", file is NOT already opened by another user.
'            Case Else: Error iErrNum ' Another error occurred.
        End Select
        On Error GoTo 0        ' Turn error checking back on.
End Function
Public Function Convert2Monthly(sZ() As String) As Boolean
    Dim sX() As String
        If GetMonthEndDates(sZ, sX) = False Then Exit Function
        If GetMonthlyData(sZ, sX) = False Then Exit Function
        ReDim sZ(UBound(sX, 1), UBound(sX, 2))
        sZ = sX
    Convert2Monthly = True
End Function
Private Function GetMonthEndDates(sZ() As String, sX() As String) As Boolean
    Dim iI%
    Dim dtStart As Date, dtEnd As Date
        If IsDate(sZ(2, 1)) = False Then Exit Function Else dtEnd = CDate(sZ(2, 1))
        If IsDate(sZ(UBound(sZ, 1), 1)) = False Then Exit Function Else dtStart = CDate(sZ(UBound(sZ, 1), 1))
        If dtStart <> Application.WorksheetFunction.EoMonth(dtStart, 0) Then dtStart = Application.WorksheetFunction.EoMonth(dtStart, 1)
        If dtStart >= dtEnd Then Exit Function
        ReDim sX(1, 1)
        If dtEnd = Application.WorksheetFunction.EoMonth(dtEnd, 0) Then sX(1, 1) = Format(dtEnd, "dd/mm/yyyy") Else sX(1, 1) = Format(Application.WorksheetFunction.EoMonth(dtEnd, -1), "dd/mm/yyyy")
        While CDate(sX(1, UBound(sX, 2))) > dtStart
            ReDim Preserve sX(1, UBound(sX, 2) + 1)
            sX(1, UBound(sX, 2)) = Format(Application.WorksheetFunction.EoMonth(CDate(sX(1, UBound(sX, 2) - 1)), -1), "dd/mm/yyyy")
        Wend
        If CDate(sX(1, UBound(sX, 2))) < dtStart Then ReDim Preserve sX(1, UBound(sX, 2) - 1)
        If TransposeString(sX) = False Then Exit Function
    GetMonthEndDates = True
End Function
Private Function GetMonthlyData(sZ() As String, sX() As String) As Boolean
    Dim iI%, iJ%, iK%, iCount%
        ReDim Preserve sX(UBound(sX, 1), UBound(sZ, 2))
        If AddRow2Start(sX) = False Then Exit Function
        For iJ = 1 To UBound(sZ, 2)
            sX(1, iJ) = sZ(1, iJ)
        Next iJ
        sX(1, 1) = "Total Return"
        iCount = 2
        For iI = 2 To UBound(sX, 1)
            For iJ = iCount To UBound(sZ, 1)
                If CDate(sX(iI, 1)) = CDate(sZ(iJ, 1)) Then
                    For iK = 2 To UBound(sZ, 2)
                        sX(iI, iK) = sZ(iJ, iK)
                    Next iK
                    iCount = iJ
                    Exit For
                End If
            Next iJ
        Next iI
    GetMonthlyData = True
End Function
Public Function MatchColumnPositions(sX() As String, sY() As String, iY() As Integer) As Boolean
    Dim iI%
        ReDim iY(UBound(sX, 2) - 1)
        For iI = 2 To UBound(sX, 2)
            If FindColumnPositionInStringArray(sX(1, iI), sY, iY(iI - 1)) = False Then iY(iI - 1) = -1
        Next iI
    MatchColumnPositions = True
End Function
