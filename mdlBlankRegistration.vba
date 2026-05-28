Option Explicit

'--------------------------------------------
' mdlBlankRegistration Module
' Generate Blank Registration Form
' No Color, 10 Students, Kalpurush Font
'--------------------------------------------

Public Sub Generate_BlankRegistration()

    '-- Performance ON --
    Call PerfOn

    '-- Sheets --
    Dim wsMadrasa  As Worksheet
    Dim wsSettings As Worksheet
    
    On Error GoTo ErrorHandler

    If Not SheetExists("Madrasa_List") Then
        MsgBox "Sheet 'Madrasa_List' not found!", vbCritical, "Error"
        Call PerfOff
        Exit Sub
    End If

    If Not SheetExists("Settings") Then
        MsgBox "Sheet 'Settings' not found!", vbCritical, "Error"
        Call PerfOff
        Exit Sub
    End If

    Set wsMadrasa = ThisWorkbook.Sheets("Madrasa_List")
    Set wsSettings = ThisWorkbook.Sheets("Settings")

    '-- Output Folder --
    Dim folderPath As String
    folderPath = ThisWorkbook.Path & "\01_BlankRegistration\"

    If Not EnsureFolder(folderPath) Then
        MsgBox "Unable to create folder!" & vbCrLf & folderPath, _
               vbCritical, "Error"
        Call PerfOff
        Exit Sub
    End If

    '-- Madrasa List - Better lastRow detection --
    Dim lastRow As Long
    lastRow = GetLastRow(wsMadrasa, 1)

    If lastRow < 2 Then
        MsgBox "No madrasa found in Madrasa_List!", _
               vbExclamation, "Warning"
        Call PerfOff
        Exit Sub
    End If

    '-- Settings Data --
    Dim boardName   As String
    Dim examName    As String
    Dim examYearEng As String
    Dim examYearBan As String
    Dim examYearHij As String
    Dim examFee     As String

    boardName = SafeGetValue(wsSettings.Range("B10"))
    examName = SafeGetValue(wsSettings.Range("B4"))
    examYearEng = SafeGetValue(wsSettings.Range("B5"))
    examYearBan = SafeGetValue(wsSettings.Range("B6"))
    examYearHij = SafeGetValue(wsSettings.Range("B7"))
    examFee = SafeGetValue(wsSettings.Range("B16")) & " " & txtFeeUnit()

    '-- Loop Start --
    Dim i            As Long
    Dim successCount As Long
    successCount = 0

    For i = 2 To lastRow

        '-- Madrasa Info --
        Dim mCode     As String
        Dim mName     As String
        Dim mZone     As String
        Dim mVillage  As String
        Dim mUnion    As String
        Dim mThana    As String
        Dim mDistrict As String
        Dim mPhone    As String

        mCode = Trim(SafeGetValue(wsMadrasa.Cells(i, 1)))
        mName = Trim(SafeGetValue(wsMadrasa.Cells(i, 2)))
        mZone = Trim(SafeGetValue(wsMadrasa.Cells(i, 3)))
        mVillage = Trim(SafeGetValue(wsMadrasa.Cells(i, 4)))
        mUnion = Trim(SafeGetValue(wsMadrasa.Cells(i, 5)))
        mThana = Trim(SafeGetValue(wsMadrasa.Cells(i, 6)))
        mDistrict = Trim(SafeGetValue(wsMadrasa.Cells(i, 7)))
        mPhone = Trim(SafeGetValue(wsMadrasa.Cells(i, 8)))

        '-- Skip if Empty --
        If mCode = "" Or mName = "" Then GoTo NextMadrasa

        '-- Create New Workbook --
        Dim newWB As Workbook
        Set newWB = Workbooks.Add

        '-- Registration Sheet Design --
        Dim wsReg As Worksheet
        Set wsReg = newWB.Sheets(1)
        
        On Error Resume Next
        wsReg.Name = txtRegForm()
        On Error GoTo ErrorHandler

        '-- Design the Sheet --
        Call DesignRegSheet( _
            wsReg, _
            boardName, examName, _
            examYearEng, examYearBan, examYearHij, _
            mCode, mName, mZone, _
            mVillage, mUnion, mThana, mDistrict, _
            mPhone, examFee)

        '-- Save File --
        Dim savePath As String
        Dim cleanName As String
        cleanName = CleanFileName(mName)
        savePath = folderPath & mCode & "_" & cleanName & "_Registration.xlsx"

        '-- Save with error handling --
        If SaveWorkbook(newWB, savePath) Then
            successCount = successCount + 1
        End If

        newWB.Close False

NextMadrasa:
    Next i

    '-- Performance OFF --
    Call PerfOff

    '-- Open Folder --
    If OpenFolder(folderPath) Then
        ' Folder opened successfully
    End If

    MsgBox txtRegForm() & " generation complete!" & vbCrLf & _
           "Total: " & successCount & " forms created." & vbCrLf & _
           "Location: " & folderPath, _
           vbInformation, "Complete"

    Exit Sub
ErrorHandler:
    Call PerfOff
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

'--------------------------------------------
' Registration Sheet Design
'--------------------------------------------
Private Sub DesignRegSheet( _
    ws As Worksheet, _
    boardName As String, _
    examName As String, _
    examYearEng As String, _
    examYearBan As String, _
    examYearHij As String, _
    mCode As String, _
    mName As String, _
    mZone As String, _
    mVillage As String, _
    mUnion As String, _
    mThana As String, _
    mDistrict As String, _
    mPhone As String, _
    examFee As String)

    With ws

        '========================================
        ' Entire Sheet: Kalpurush, Middle Align
        '========================================
        .Cells.Font.Name = "Kalpurush"
        .Cells.Font.Size = 12
        .Cells.VerticalAlignment = xlCenter

        '-- Column Width --
        .Columns("A").ColumnWidth = 4
        .Columns("B").ColumnWidth = 14
        .Columns("C").ColumnWidth = 12
        .Columns("D").ColumnWidth = 12
        .Columns("E").ColumnWidth = 12
        .Columns("F").ColumnWidth = 12
        .Columns("G").ColumnWidth = 10
        .Columns("H").ColumnWidth = 10
        .Columns("I").ColumnWidth = 10
        .Columns("J").ColumnWidth = 10
        .Columns("K").ColumnWidth = 10
        .Columns("L").ColumnWidth = 6
        .Columns("M").ColumnWidth = 12

        '----------------------------
        ' Row 1-3: Main Header (NO COLOR)
        '----------------------------
        .Range("A1:M1").Merge
        .Range("A1").Value = boardName
        .Range("A1").Font.Size = 18
        .Range("A1").Font.Bold = True
        .Range("A1").HorizontalAlignment = xlCenter
        .Rows(1).RowHeight = 30

        .Range("A2:M2").Merge
        .Range("A2").Value = examName & " " & _
            examYearEng & " " & txtEngYear() & "/" & _
            examYearBan & " " & txtBanYear() & "/" & _
            examYearHij & " " & txtHijYear()
        .Range("A2").Font.Size = 14
        .Range("A2").Font.Bold = True
        .Range("A2").HorizontalAlignment = xlCenter
        .Rows(2).RowHeight = 26

        .Range("A3:M3").Merge
        .Range("A3").Value = txtRegForm()
        .Range("A3").Font.Size = 15
        .Range("A3").Font.Bold = True
        .Range("A3").HorizontalAlignment = xlCenter
        .Rows(3).RowHeight = 26

        '-- Row 1-3 Border --
      '  .Range("A1:M3").Borders.LineStyle = xlContinuous

        '----------------------------
        ' Row 5: Madrasa Name (Full Row)
        '----------------------------
        .Range("A5:M5").Merge
        .Range("A5").Value = lblMadrasa() & " " & mName & _
            "  |  " & lblCode() & " " & mCode
        .Range("A5").Font.Bold = True
        .Range("A5").Font.Size = 14
        .Range("A5").HorizontalAlignment = xlCenter
        .Range("A5").WrapText = True
        .Rows(5).RowHeight = 30

        '----------------------------
        ' Row 6: Village, Union, Thana, District
        '----------------------------
        .Cells(6, 1).Value = lblVillage()
        .Cells(6, 1).Font.Bold = True
        .Cells(6, 1).HorizontalAlignment = xlCenter

        .Range("B6:C6").Merge
        .Range("B6").Value = mVillage
        .Range("B6").HorizontalAlignment = xlCenter

        .Cells(6, 4).Value = lblUnion()
        .Cells(6, 4).Font.Bold = True
        .Cells(6, 4).HorizontalAlignment = xlCenter

        .Range("E6:F6").Merge
        .Range("E6").Value = mUnion
        .Range("E6").HorizontalAlignment = xlCenter

        .Cells(6, 7).Value = lblThana()
        .Cells(6, 7).Font.Bold = True
        .Cells(6, 7).HorizontalAlignment = xlCenter

        .Range("H6:I6").Merge
        .Range("H6").Value = mThana
        .Range("H6").HorizontalAlignment = xlCenter

        .Cells(6, 10).Value = lblDistrict()
        .Cells(6, 10).Font.Bold = True
        .Cells(6, 10).HorizontalAlignment = xlCenter

        .Range("K6:M6").Merge
        .Range("K6").Value = mDistrict
        .Range("K6").HorizontalAlignment = xlCenter

        .Rows(6).RowHeight = 22

        '----------------------------
        ' Row 7: Jamat, Zone, Phone, Fee
        '----------------------------
        .Cells(7, 1).Value = hdrJamat()
        .Cells(7, 1).Font.Bold = True
        .Cells(7, 1).HorizontalAlignment = xlCenter

        '-- Jamat Box (Empty) --
        .Range("B7:C7").Merge
        .Range("B7").Value = ""
        .Range("B7").HorizontalAlignment = xlCenter
        .Range("B7").Borders.LineStyle = xlContinuous

        .Cells(7, 4).Value = hdrZone()
        .Cells(7, 4).Font.Bold = True
        .Cells(7, 4).HorizontalAlignment = xlCenter

        .Range("E7:F7").Merge
        .Range("E7").Value = mZone
        .Range("E7").HorizontalAlignment = xlCenter

        .Cells(7, 7).Value = lblPhone()
        .Cells(7, 7).Font.Bold = True
        .Cells(7, 7).HorizontalAlignment = xlCenter

        .Range("H7:I7").Merge
        .Range("H7").Value = mPhone
        .Range("H7").HorizontalAlignment = xlCenter

        .Cells(7, 10).Value = lblFee()
        .Cells(7, 10).Font.Bold = True
        .Cells(7, 10).HorizontalAlignment = xlCenter

        .Range("K7:M7").Merge
        .Range("K7").Value = examFee
        .Range("K7").HorizontalAlignment = xlCenter

        .Rows(7).RowHeight = 22

        '-- Row 5:7 Border --
        .Range("A5:M7").Borders.LineStyle = xlContinuous

        '----------------------------
        ' Row 9: Table Header (NO COLOR)
        '----------------------------
        .Cells(9, 1).Value = hdrSerial()
        .Cells(9, 2).Value = lblRegNo()

        .Range("C9:D9").Merge
        .Range("C9").Value = hdrName()

        .Range("E9:F9").Merge
        .Range("E9").Value = hdrFather()

        .Cells(9, 7).Value = lblDOB()
        .Cells(9, 8).Value = lblVillage()
        .Cells(9, 9).Value = lblPostOffice()
        .Cells(9, 10).Value = lblThana()
        .Cells(9, 11).Value = lblDistrict()
        .Cells(9, 12).Value = lblFee()
        .Cells(9, 13).Value = lblSignature()

        '-- Header Format (Bold, Center, Border ONLY) --
        .Range("A9:M9").Font.Bold = True
        .Range("A9:M9").Font.Size = 12
        .Range("A9:M9").HorizontalAlignment = xlCenter
        .Range("A9:M9").VerticalAlignment = xlCenter
        .Range("A9:M9").WrapText = True
        .Range("A9:M9").Borders.LineStyle = xlContinuous
        .Rows(9).RowHeight = 50

        '----------------------------
        ' Row 10-19: Empty Table (10 rows)
        ' NO COLOR, Height 28
        '----------------------------
        Dim r As Integer
        For r = 10 To 19

            .Cells(r, 1).Value = r - 9  ' Serial
            .Cells(r, 1).HorizontalAlignment = xlCenter

            '-- Name Merge --
            .Range(.Cells(r, 3), .Cells(r, 4)).Merge
            '-- Father Merge --
            .Range(.Cells(r, 5), .Cells(r, 6)).Merge

            '-- Border --
            .Range(.Cells(r, 1), .Cells(r, 13)) _
                .Borders.LineStyle = xlContinuous

            '-- Row Height 28 --
            .Rows(r).RowHeight = 28

        Next r

        '----------------------------
        ' Row 20: Footer - Only Collector Signature
        ' Tight with row 19, Right side
        '----------------------------
        Dim footerRow As Long
        footerRow = 20  ' Right after last student

        .Cells(footerRow, 13).Value = lblCollector()
        .Cells(footerRow, 13).Font.Bold = True
        .Cells(footerRow, 13).HorizontalAlignment = xlRight
        .Cells(footerRow, 13).VerticalAlignment = xlCenter

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
' Get Label Text from VBA_Helper I Column
'--------------------------------------------
Private Function lblMadrasa() As String
    lblMadrasa = ReadHelperI(2)
End Function

Private Function lblCode() As String
    lblCode = ReadHelperI(3)
End Function

Private Function lblVillage() As String
    lblVillage = ReadHelperI(4)
End Function

Private Function lblUnion() As String
    lblUnion = ReadHelperI(5)
End Function

Private Function lblThana() As String
    lblThana = ReadHelperI(6)
End Function

Private Function lblDistrict() As String
    lblDistrict = ReadHelperI(7)
End Function

Private Function lblPhone() As String
    lblPhone = ReadHelperI(8)
End Function

Private Function lblFee() As String
    lblFee = ReadHelperI(9)
End Function

Private Function lblRegNo() As String
    lblRegNo = ReadHelperI(10)
End Function

Private Function lblDOB() As String
    lblDOB = ReadHelperI(11)
End Function

Private Function lblPostOffice() As String
    lblPostOffice = ReadHelperI(12)
End Function

Private Function lblSignature() As String
    lblSignature = ReadHelperI(13)
End Function

Private Function lblCollector() As String
    lblCollector = ReadHelperI(14)
End Function

Private Function lblDate() As String
    lblDate = ReadHelperI(15)
End Function

Private Function lblSealSign() As String
    lblSealSign = ReadHelperI(16)
End Function

Private Function lblBoardUse() As String
    lblBoardUse = ReadHelperI(17)
End Function

Private Function hdrSerial() As String
    hdrSerial = ReadHelperI(22)
End Function

Private Function hdrName() As String
    hdrName = ReadHelperI(23)
End Function

Private Function hdrFather() As String
    hdrFather = ReadHelperI(24)
End Function

Private Function hdrJamat() As String
    hdrJamat = ReadHelperI(25)
End Function

Private Function hdrZone() As String
    hdrZone = ReadHelperI(26)
End Function

Private Function txtRegForm() As String
    txtRegForm = ReadHelperI(27)
End Function

Private Function txtEngYear() As String
    txtEngYear = ReadHelperI(18)
End Function

Private Function txtBanYear() As String
    txtBanYear = ReadHelperI(19)
End Function

Private Function txtHijYear() As String
    txtHijYear = ReadHelperI(20)
End Function

Private Function txtFeeUnit() As String
    txtFeeUnit = ReadHelperI(21)
End Function

'--------------------------------------------
' Improved Helper Functions
'--------------------------------------------

'-- Safe Helper I Column Reader with Error Handling --
Private Function ReadHelperI(Row As Integer) As String
    Dim wsHelper As Worksheet
    Dim result As String
    
    On Error GoTo ErrorHandler
    
    If Not SheetExists("VBA_Helper") Then
        ReadHelperI = ""
        Exit Function
    End If
    
    Set wsHelper = ThisWorkbook.Sheets("VBA_Helper")
    result = Trim(wsHelper.Range("I" & Row).Value)
    
    If result = "" Then
        result = ""  ' Return empty string if value is empty
    End If
    
    ReadHelperI = result
    Exit Function
    
ErrorHandler:
    ReadHelperI = ""
End Function

'-- Clean File Name (Remove invalid characters) --
Private Function CleanFileName(fileName As String) As String
    Dim invalidChars As String
    Dim i As Integer
    Dim result As String
    
    invalidChars = "<>:/\|?*" & Chr(34)
    result = fileName
    
    For i = 1 To Len(invalidChars)
        result = Replace(result, Mid(invalidChars, i, 1), "_")
    Next i
    
    ' Limit length
    If Len(result) > 50 Then
        result = Left(result, 50)
    End If
    
    CleanFileName = result
End Function

'-- Check if Sheet Exists --
Private Function SheetExists(sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

'-- Get Last Row in Column --
Private Function GetLastRow(ws As Worksheet, columnNum As Long) As Long
    Dim lastRow As Long
    On Error Resume Next
    lastRow = ws.Cells(ws.Rows.Count, columnNum).End(xlUp).Row
    On Error GoTo 0
    
    If lastRow = 0 Then
        lastRow = 1
    End If
    
    GetLastRow = lastRow
End Function

'-- Safe Get Cell Value --
Private Function SafeGetValue(cell As Range) As String
    On Error Resume Next
    SafeGetValue = CStr(cell.Value)
    If Err.Number <> 0 Then
        SafeGetValue = ""
    End If
    On Error GoTo 0
End Function

'-- Save Workbook with Error Handling --
Private Function SaveWorkbook(wb As Workbook, filePath As String) As Boolean
    On Error GoTo ErrorHandler
    wb.SaveAs filePath, xlOpenXMLWorkbook
    SaveWorkbook = True
    Exit Function
    
ErrorHandler:
    MsgBox "Error saving file: " & filePath & vbCrLf & Err.Description, vbExclamation
    SaveWorkbook = False
End Function

'-- Open Folder (Cross-platform compatible) --
Private Function OpenFolder(folderPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    #If Win64 Or Win32 Then
        Shell "explorer.exe " & folderPath, vbNormalFocus
    #ElseIf Mac Then
        Shell "open """ & folderPath & """"
    #End If
    
    OpenFolder = True
    Exit Function
    
ErrorHandler:
    OpenFolder = False
End Function

'--------------------------------------------
' Dashboard Button
'--------------------------------------------
Public Sub Btn_BlankRegistration()

    Dim wsMadrasa As Worksheet
    
    On Error GoTo ErrorHandler
    
    If Not SheetExists("Madrasa_List") Then
        MsgBox "Sheet 'Madrasa_List' not found!", vbCritical, "Error"
        Exit Sub
    End If

    Set wsMadrasa = ThisWorkbook.Sheets("Madrasa_List")

    Dim madrasaCount As Long
    madrasaCount = GetLastRow(wsMadrasa, 1) - 1

    If madrasaCount <= 0 Then
        MsgBox "No madrasa found in Madrasa_List!", _
               vbExclamation, "Warning"
        Exit Sub
    End If

    Dim response As VbMsgBoxResult
    response = MsgBox( _
        "Generate " & txtRegForm() & "?" & vbCrLf & _
        "Total Madrasa: " & madrasaCount, _
        vbYesNo + vbQuestion, "Confirm")

    If response = vbYes Then
        Call Generate_BlankRegistration
    End If
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub
