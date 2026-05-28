Option Explicit

'--------------------------------------------
' mdlBlankMarkSheet Module
' Improved - No Color, 10 Students Per Page
'--------------------------------------------

Public Sub Generate_BlankMarkSheets()

    '-- Performance ON --
    Call PerfOn

    '-- Sheets --
    Dim wsStudent  As Worksheet
    Dim wsMadrasa  As Worksheet
    Dim wsSubject  As Worksheet
    Dim wsSettings As Worksheet

    Set wsStudent = ThisWorkbook.Sheets("Student_Database")
    Set wsMadrasa = ThisWorkbook.Sheets("Madrasa_List")
    Set wsSubject = ThisWorkbook.Sheets("Subject_Structure")
    Set wsSettings = ThisWorkbook.Sheets("Settings")

    '-- Check if Students Exist --
    Dim totalStudents As Long
    totalStudents = Application.CountA( _
        wsStudent.Range("D2:D5000"))

    If totalStudents = 0 Then
        MsgBox "No students found in Student_Database!", _
               vbExclamation, "Warning"
        Call PerfOff
        Exit Sub
    End If

    '-- Output Folder --
    Dim folderPath As String
    folderPath = ThisWorkbook.Path & "\03_BlankSheets\"

    If Not EnsureFolder(folderPath) Then
        MsgBox "Unable to create folder!" & vbCrLf & folderPath, _
               vbCritical, "Error"
        Call PerfOff
        Exit Sub
    End If

    '-- Settings Data --
    Dim boardName   As String
    Dim examName    As String
    Dim examYearEng As String
    Dim examYearBan As String
    Dim examYearHij As String

    boardName = wsSettings.Range("B10").Value
    examName = wsSettings.Range("B4").Value
    examYearEng = wsSettings.Range("B5").Value
    examYearBan = wsSettings.Range("B6").Value
    examYearHij = wsSettings.Range("B7").Value

    '-- Madrasa List --
    Dim lastMadrasa As Long
    lastMadrasa = wsMadrasa.Cells(Rows.count, 1).End(xlUp).Row

    Dim i            As Long
    Dim successCount As Long
    successCount = 0

    For i = 2 To lastMadrasa

        Dim mCode As String
        Dim mName As String
        Dim mZone As String

        mCode = Trim(wsMadrasa.Cells(i, 1).Value)
        mName = Trim(wsMadrasa.Cells(i, 2).Value)
        mZone = Trim(wsMadrasa.Cells(i, 3).Value)

        If mCode = "" Or mName = "" Then GoTo NextMadrasa

        '-- Check if Madrasa has Students --
        Dim stuCount As Long
        stuCount = Application.CountIf( _
            wsStudent.Range("K:K"), mCode)

        If stuCount = 0 Then GoTo NextMadrasa

        '-- Create New Workbook --
        Dim newWB      As Workbook
        Dim sheetAdded As Boolean

        Set newWB = Workbooks.Add
        sheetAdded = False

        '-- Loop Through Each Jamat --
        Dim jamats() As String
        Dim groups() As String
        jamats = GetJamats()
        groups = GetGroups()

        Dim j As Integer
        Dim g As Integer

        For j = 0 To UBound(jamats)

            If IsHifz(jamats(j)) Then

                For g = 0 To UBound(groups)

                    Dim grpCount As Long
                    grpCount = Application.CountIfs( _
                        wsStudent.Range("K:K"), mCode, _
                        wsStudent.Range("N:N"), jamats(j), _
                        wsStudent.Range("O:O"), groups(g))

                    If grpCount > 0 Then
                        Call CreateBlankSheet( _
                            newWB, wsStudent, wsSubject, _
                            mCode, mName, mZone, _
                            jamats(j), groups(g), _
                            boardName, examName, _
                            examYearEng, examYearBan, examYearHij, _
                            sheetAdded)
                        sheetAdded = True
                    End If

                Next g

            Else

                Dim jamatCount As Long
                jamatCount = Application.CountIfs( _
                    wsStudent.Range("K:K"), mCode, _
                    wsStudent.Range("N:N"), jamats(j))

                If jamatCount > 0 Then
                    Call CreateBlankSheet( _
                        newWB, wsStudent, wsSubject, _
                        mCode, mName, mZone, _
                        jamats(j), "", _
                        boardName, examName, _
                        examYearEng, examYearBan, examYearHij, _
                        sheetAdded)
                    sheetAdded = True
                End If

            End If

        Next j

        '-- Delete Extra Sheet --
        If sheetAdded And newWB.Sheets.count > 1 Then
            Application.DisplayAlerts = False
            newWB.Sheets(newWB.Sheets.count).Delete
            Application.DisplayAlerts = True
        End If

        '-- Save File --
        If sheetAdded Then
            Dim savePath As String
            savePath = folderPath & mCode & "_" & _
                       CleanFileName(mName) & "_Blank.xlsx"
            newWB.SaveAs savePath, xlOpenXMLWorkbook
            successCount = successCount + 1
        End If

        newWB.Close False

NextMadrasa:
    Next i

    '-- Performance OFF --
    Call PerfOff

    Shell "explorer.exe " & folderPath, vbNormalFocus

    MsgBox txtBlankSheet() & " generation complete!" & vbCrLf & _
           "Total: " & successCount & " files created.", _
           vbInformation, "Complete"

End Sub

'--------------------------------------------
' Create One Blank Sheet
'--------------------------------------------
Private Sub CreateBlankSheet( _
    newWB As Workbook, _
    wsStudent As Worksheet, _
    wsSubject As Worksheet, _
    mCode As String, _
    mName As String, _
    mZone As String, _
    jamat As String, _
    grp As String, _
    boardName As String, _
    examName As String, _
    examYearEng As String, _
    examYearBan As String, _
    examYearHij As String, _
    sheetAdded As Boolean)

    '-- New Sheet --
    Dim wsBlank As Worksheet

    If Not sheetAdded Then
        Set wsBlank = newWB.Sheets(1)
    Else
        Set wsBlank = newWB.Sheets.Add( _
            After:=newWB.Sheets(newWB.Sheets.count))
    End If

    '-- Sheet Name --
    If grp = "" Then
        wsBlank.Name = jamat
    Else
        wsBlank.Name = jamat & "_" & grp
    End If

    '-- Get Subject Info --
    Dim subInfo(1 To 5, 1 To 2) As Variant
    Call GetSubjectInfoBlank(wsSubject, jamat, subInfo)

    Dim totalFull As Variant
    Dim subCount  As Variant
    totalFull = Application.VLookup( _
        jamat, wsSubject.Range("A:M"), 12, 0)
    subCount = Application.VLookup( _
        jamat, wsSubject.Range("A:M"), 13, 0)

    If IsError(totalFull) Then totalFull = 0
    If IsError(subCount) Then subCount = 0

    '-- Count actual subjects --
    Dim actualSubCount As Integer
    actualSubCount = 0
    Dim sc As Integer
    For sc = 1 To 5
        If subInfo(sc, 1) <> "" Then
            actualSubCount = actualSubCount + 1
        End If
    Next sc

    If subCount = 0 Then subCount = actualSubCount

    '-- Collect matching student rows --
    Dim lastStudent As Long
    lastStudent = wsStudent.Cells(Rows.count, 1).End(xlUp).Row

    Dim stuRows() As Long
    Dim stuTotal  As Long
    stuTotal = 0

    ReDim stuRows(1 To lastStudent)

    Dim s As Long
    For s = 2 To lastStudent
        Dim sCode  As String
        Dim sJamat As String
        Dim sGrp   As String

        sCode = wsStudent.Cells(s, 11).Value
        sJamat = wsStudent.Cells(s, 14).Value
        sGrp = wsStudent.Cells(s, 15).Value

        Dim isMatch As Boolean
        isMatch = (sCode = mCode) And (sJamat = jamat)

        If grp <> "" Then
            isMatch = isMatch And (sGrp = grp)
        End If

        If isMatch Then
            stuTotal = stuTotal + 1
            stuRows(stuTotal) = s
        End If
    Next s

    If stuTotal = 0 Then Exit Sub
    ReDim Preserve stuRows(1 To stuTotal)

    '-- Constants --
    Const STUDENTS_PER_PAGE As Integer = 10
    Const DATA_START_ROW As Integer = 10

    With wsBlank

        '========================================
        ' FIX 1: Entire Sheet = Middle Align
        '========================================
        .Cells.Font.Name = "Kalpurush"
        .Cells.Font.Size = 11
        .Cells.VerticalAlignment = xlCenter

        '-- Column Width --
        .Columns("A").ColumnWidth = 5
        .Columns("B").ColumnWidth = 8
        .Columns("C").ColumnWidth = 16
        .Columns("D").ColumnWidth = 8
        .Columns("E").ColumnWidth = 8
        .Columns("F").ColumnWidth = 8
        .Columns("G").ColumnWidth = 10
        .Columns("H").ColumnWidth = 10
        .Columns("I").ColumnWidth = 10
        .Columns("J").ColumnWidth = 10
        .Columns("K").ColumnWidth = 10
        .Columns("L").ColumnWidth = 8
        .Columns("M").ColumnWidth = 12

        '-- Calculate dynamic columns first --
        Dim si       As Integer
        Dim startCol As Integer
        Dim totalCol As Integer
        Dim remarksCol As Integer
        Dim lastCol  As Integer

        startCol = 7
        For si = 1 To 5
            If subInfo(si, 1) <> "" Then
                startCol = startCol + 1
            End If
        Next si

        totalCol = startCol
        remarksCol = startCol + 1
        lastCol = remarksCol

        '----------------------------
        ' Row 1-3: Main Header
        '----------------------------
        .Range("A1:M1").Merge
        .Range("A1").Value = boardName
        .Range("A1").Font.Size = 16
        .Range("A1").Font.Bold = True
        .Range("A1").HorizontalAlignment = xlCenter
        .Rows(1).RowHeight = 28

        .Range("A2:M2").Merge
        .Range("A2").Value = examName & " " & _
            examYearEng & " " & ReadHelperI(18) & "/" & _
            examYearBan & " " & ReadHelperI(19) & "/" & _
            examYearHij & " " & ReadHelperI(20)
        .Range("A2").Font.Size = 13
        .Range("A2").Font.Bold = True
        .Range("A2").HorizontalAlignment = xlCenter
        .Rows(2).RowHeight = 24

        .Range("A3:M3").Merge
        .Range("A3").Value = txtBlankSheet()
        .Range("A3").Font.Size = 12
        .Range("A3").Font.Bold = True
        .Range("A3").HorizontalAlignment = xlCenter
        .Rows(3).RowHeight = 22

        .Range("A1:M3").Borders.LineStyle = xlContinuous

        '----------------------------
        ' Row 5: Madrasa Name
        '----------------------------
        .Range("A5:M5").Merge
        .Range("A5").Value = ReadHelperI(2) & " " & mName
        .Range("A5").Font.Bold = True
        .Range("A5").Font.Size = 14
        .Range("A5").HorizontalAlignment = xlCenter
        .Range("A5").WrapText = True
        .Rows(5).RowHeight = 30

        '========================================
        ' FIX 2: Row 6 & 7 = Center + Middle
        '========================================

        '-- Row 6: Code, Zone, Exam Year --
        .Cells(6, 1).Value = ReadHelperI(3)
        .Cells(6, 1).Font.Bold = True
        .Cells(6, 1).HorizontalAlignment = xlCenter

        .Range("B6:C6").Merge
        .Range("B6").Value = mCode
        .Range("B6").Font.Bold = True
        .Range("B6").HorizontalAlignment = xlCenter

        .Cells(6, 4).Value = hdrZone()
        .Cells(6, 4).Font.Bold = True
        .Cells(6, 4).HorizontalAlignment = xlCenter

        .Range("E6:G6").Merge
        .Range("E6").Value = mZone
        .Range("E6").HorizontalAlignment = xlCenter

        .Cells(6, 8).Value = lblExamYear()
        .Cells(6, 8).Font.Bold = True
        .Cells(6, 8).HorizontalAlignment = xlCenter

        .Range("I6:K6").Merge
        .Range("I6").Value = examYearEng & " " & ReadHelperI(18)
        .Range("I6").HorizontalAlignment = xlCenter

        .Rows(6).RowHeight = 20

        '-- Row 7: Jamat, Group, Total Students --
        .Cells(7, 1).Value = hdrJamat()
        .Cells(7, 1).Font.Bold = True
        .Cells(7, 1).HorizontalAlignment = xlCenter

        .Range("B7:C7").Merge
        .Range("B7").Value = jamat
        .Range("B7").Font.Bold = True
        .Range("B7").HorizontalAlignment = xlCenter

        .Cells(7, 4).Value = hdrGroup()
        .Cells(7, 4).Font.Bold = True
        .Cells(7, 4).HorizontalAlignment = xlCenter

        .Range("E7:G7").Merge
        .Range("E7").Value = IIf(grp = "", "-", grp)
        .Range("E7").HorizontalAlignment = xlCenter

        .Cells(7, 8).Value = lblTotalStudents()
        .Cells(7, 8).Font.Bold = True
        .Cells(7, 8).HorizontalAlignment = xlCenter

        .Range("I7:K7").Merge
        .Range("I7").Value = stuTotal
        .Range("I7").Font.Bold = True
        .Range("I7").HorizontalAlignment = xlCenter

        .Rows(7).RowHeight = 20

        '-- Row 5:7 Border --
        .Range("A5:M7").Borders.LineStyle = xlContinuous

        '========================================
        ' FIX 3: Row 9 Header = Height 50
        '========================================
        .Cells(9, 1).Value = hdrSerial()
        .Cells(9, 2).Value = hdrRoll()
        .Cells(9, 3).Value = hdrReg()

        .Range("D9:F9").Merge
        .Range("D9").Value = hdrName()

        '-- Subject Headers --
        Dim colIdx As Integer
        colIdx = 7

        For si = 1 To 5
            If subInfo(si, 1) <> "" Then
                .Cells(9, colIdx).Value = subInfo(si, 1) & _
                    vbCrLf & "(" & subInfo(si, 2) & ")"
                colIdx = colIdx + 1
            End If
        Next si

        '-- Total Column --
        .Cells(9, totalCol).Value = hdrTotal() & _
            vbCrLf & "(" & totalFull & ")"

        '-- Remarks Column --
        .Cells(9, remarksCol).Value = hdrRemarks()

        '-- Header Format --
        .Range(.Cells(9, 1), .Cells(9, lastCol)).Font.Bold = True
        .Range(.Cells(9, 1), .Cells(9, lastCol)).HorizontalAlignment = xlCenter
        .Range(.Cells(9, 1), .Cells(9, lastCol)).VerticalAlignment = xlCenter
        .Range(.Cells(9, 1), .Cells(9, lastCol)).WrapText = True
        .Range(.Cells(9, 1), .Cells(9, lastCol)).Borders.LineStyle = xlContinuous
        .Rows(9).RowHeight = 50

        '-- Hide unused columns --
        Dim hideCol As Integer
        For hideCol = lastCol + 1 To 13
            .Columns(hideCol).Hidden = True
        Next hideCol

        '----------------------------
        ' Student Data + Footer
        '----------------------------
        Dim currentRow As Long
        Dim serial     As Long
        Dim stuOnPage  As Integer

        currentRow = DATA_START_ROW
        serial = 1
        stuOnPage = 0

        Dim idx As Long
        For idx = 1 To stuTotal

            Dim stuRow As Long
            stuRow = stuRows(idx)

            '-- Serial --
            .Cells(currentRow, 1).Value = serial
            .Cells(currentRow, 1).HorizontalAlignment = xlCenter

            '-- Roll --
            .Cells(currentRow, 2).Value = _
                wsStudent.Cells(stuRow, 2).Value
            .Cells(currentRow, 2).HorizontalAlignment = xlCenter

            '-- Registration --
            .Cells(currentRow, 3).Value = _
                wsStudent.Cells(stuRow, 3).Value

            '-- Name (Merged) --
            .Range(.Cells(currentRow, 4), _
                   .Cells(currentRow, 6)).Merge
            .Cells(currentRow, 4).Value = _
                wsStudent.Cells(stuRow, 4).Value

            '-- Subject Mark Boxes (Empty) --
            Dim markCol As Integer
            For markCol = 7 To (7 + actualSubCount - 1)
                .Cells(currentRow, markCol).Value = ""
                .Cells(currentRow, markCol).HorizontalAlignment = xlCenter
            Next markCol

            '-- Total Column (Empty) --
            .Cells(currentRow, totalCol).Value = ""
            .Cells(currentRow, totalCol).HorizontalAlignment = xlCenter

            '-- Remarks Column (Empty) --
            .Cells(currentRow, remarksCol).Value = ""

            '-- Row Border --
            .Range(.Cells(currentRow, 1), _
                   .Cells(currentRow, lastCol)) _
                .Borders.LineStyle = xlContinuous

            '========================================
            ' FIX 4: Student Row Height = 26
            '========================================
            .Rows(currentRow).RowHeight = 26

            serial = serial + 1
            currentRow = currentRow + 1
            stuOnPage = stuOnPage + 1

            '========================================
            ' 10 students done = footer + page break
            '========================================
            If stuOnPage = STUDENTS_PER_PAGE Then

                '========================================
                ' FIX 5: ONLY Head Examiner - tight
                '========================================
                Dim fRow As Long
                fRow = currentRow  ' No gap

                .Cells(fRow, remarksCol).Value = lblHeadExaminer()
                .Cells(fRow, remarksCol).Font.Bold = True
                .Cells(fRow, remarksCol).HorizontalAlignment = xlRight
                .Cells(fRow, remarksCol).VerticalAlignment = xlCenter

                '-- Page Break if more students --
                If idx < stuTotal Then
                    .HPageBreaks.Add Before:=.Rows(fRow + 1)
                    currentRow = fRow + 1
                    stuOnPage = 0
                Else
                    currentRow = fRow + 1
                End If

            End If

        Next idx

        '========================================
        ' Last page footer (if < 10 students)
        '========================================
        If stuOnPage > 0 And stuOnPage < STUDENTS_PER_PAGE Then

        Dim lfRow As Long
        lfRow = currentRow  ' No gap - tight

        .Cells(lfRow, remarksCol).Value = lblHeadExaminer()
        .Cells(lfRow, remarksCol).Font.Bold = True
        .Cells(lfRow, remarksCol).HorizontalAlignment = xlRight
        .Cells(lfRow, remarksCol).VerticalAlignment = xlCenter

        End If

        '----------------------------
        ' Print Setup
        '----------------------------
        With .PageSetup
            .Orientation = xlLandscape
            .PaperSize = xlPaperA4
            .FitToPagesWide = 1
            .FitToPagesTall = False
            .PrintTitleRows = "$1:$9"
            .TopMargin = Application.InchesToPoints(0.5)
            .BottomMargin = Application.InchesToPoints(0.5)
            .LeftMargin = Application.InchesToPoints(0.5)
            .RightMargin = Application.InchesToPoints(0.5)
            .CenterHorizontally = True
        End With

    End With

End Sub

'--------------------------------------------
' Get Subject Info
'--------------------------------------------
Private Sub GetSubjectInfoBlank( _
    wsSubject As Worksheet, _
    jamat As String, _
    subInfo() As Variant)

    Dim si As Integer
    For si = 1 To 5
        Dim nameCol As Integer
        Dim fullCol As Integer
        nameCol = (si - 1) * 2 + 2
        fullCol = nameCol + 1

        Dim subName As Variant
        Dim subFull As Variant

        subName = Application.VLookup( _
            jamat, wsSubject.Range("A:M"), nameCol, 0)
        subFull = Application.VLookup( _
            jamat, wsSubject.Range("A:M"), fullCol, 0)

        If IsError(subName) Or subName = "" Then
            subInfo(si, 1) = ""
            subInfo(si, 2) = 0
        Else
            subInfo(si, 1) = subName
            subInfo(si, 2) = IIf(IsError(subFull), 0, subFull)
        End If
    Next si

End Sub

'--------------------------------------------
' VBA_Helper K Column Labels
'--------------------------------------------
Private Function lblTotalStudents() As String
    lblTotalStudents = ReadHelperK(2)
End Function

Private Function lblExamYear() As String
    lblExamYear = ReadHelperK(3)
End Function

Private Function lblExaminer1() As String
    lblExaminer1 = ReadHelperK(4)
End Function

Private Function lblExaminer2() As String
    lblExaminer2 = ReadHelperK(5)
End Function

Private Function lblHeadExaminer() As String
    lblHeadExaminer = ReadHelperK(6)
End Function

Private Function ReadHelperK(Row As Integer) As String
    On Error Resume Next
    ReadHelperK = Trim( _
        ThisWorkbook.Sheets("VBA_Helper").Range("K" & Row).Value)
    On Error GoTo 0
End Function

Private Function ReadHelperI(Row As Integer) As String
    On Error Resume Next
    ReadHelperI = Trim( _
        ThisWorkbook.Sheets("VBA_Helper").Range("I" & Row).Value)
    On Error GoTo 0
End Function

'--------------------------------------------
' Dashboard Button
'--------------------------------------------
Public Sub Btn_BlankMarkSheet()

    Dim wsStudent As Worksheet
    Set wsStudent = ThisWorkbook.Sheets("Student_Database")

    Dim totalStudents As Long
    totalStudents = Application.CountA( _
        wsStudent.Range("D2:D5000"))

    If totalStudents = 0 Then
        MsgBox "No students found in Student_Database!", _
               vbExclamation, "Warning"
        Exit Sub
    End If

    Dim response As VbMsgBoxResult
    response = MsgBox( _
        "Generate " & txtBlankSheet() & "?" & vbCrLf & _
        "Total Students: " & totalStudents & vbCrLf & _
        "This may take some time.", _
        vbYesNo + vbQuestion, "Confirm")

    If response = vbYes Then
        Call Generate_BlankMarkSheets
    End If

End Sub

