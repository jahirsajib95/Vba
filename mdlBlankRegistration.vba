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

    '-- Madrasa List --
    Dim lastRow As Long
    lastRow = wsMadrasa.Cells(Rows.Count, 1).End(xlUp).Row

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

    boardName = wsSettings.Range("B10").Value
    examName = wsSettings.Range("B4").Value
    examYearEng = wsSettings.Range("B5").Value
    examYearBan = wsSettings.Range("B6").Value
    examYearHij = wsSettings.Range("B7").Value
    examFee = wsSettings.Range("B16").Value & " " & GetFeeUnit()

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

        mCode = Trim(wsMadrasa.Cells(i, 1).Value)
        mName = Trim(wsMadrasa.Cells(i, 2).Value)
        mZone = Trim(wsMadrasa.Cells(i, 3).Value)
        mVillage = Trim(wsMadrasa.Cells(i, 4).Value)
        mUnion = Trim(wsMadrasa.Cells(i, 5).Value)
        mThana = Trim(wsMadrasa.Cells(i, 6).Value)
        mDistrict = Trim(wsMadrasa.Cells(i, 7).Value)
        mPhone = Trim(wsMadrasa.Cells(i, 8).Value)

        '-- Skip if Empty --
        If mCode = "" Or mName = "" Then GoTo NextMadrasa

        '-- Create New Workbook --
        Dim newWB As Workbook
        Set newWB = Workbooks.Add

        '-- Registration Sheet Design --
        Dim wsReg As Worksheet
        Set wsReg = newWB.Sheets(1)
        wsReg.Name = txtRegForm()

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
        savePath = folderPath & mCode & "_" & _
                   CleanFileName(mName) & "_Registration.xlsx"

        On Error Resume Next
        newWB.SaveAs savePath, xlOpenXMLWorkbook
        If Err.Number = 0 Then
            successCount = successCount + 1
        Else
            MsgBox "Error saving: " & mName & vbCrLf & Err.Description, vbExclamation
        End If
        On Error GoTo ErrorHandler

        newWB.Close False

NextMadrasa:
    Next i

    '-- Performance OFF --
    Call PerfOff

    '-- Open Folder --
    On Error Resume Next
    Shell "explorer.exe " & folderPath, vbNormalFocus
    On Error GoTo 0

    MsgBox txtRegForm() & " generation complete!" & vbCrLf & _
           "Total: " & successCount & " forms created." & vbCrLf & _
           "Location: " & folderPath, _
           vbInformation, "Complete"

    Exit Sub
ErrorHandler:
    Call PerfOff
    MsgBox "Error: " & Err.Description, vbCritical, "Error in Generate_BlankRegistration"
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
        .Cells.Font.Size = 11
        .Cells.VerticalAlignment = xlCenter

        '-- Column Width --
        .Columns("A").ColumnWidth = 3.5
        .Columns("B").ColumnWidth = 12
        .Columns("C").ColumnWidth = 10
        .Columns("D").ColumnWidth = 10
        .Columns("E").ColumnWidth = 10
        .Columns("F").ColumnWidth = 10
        .Columns("G").ColumnWidth = 8
        .Columns("H").ColumnWidth = 8
        .Columns("I").ColumnWidth = 8
        .Columns("J").ColumnWidth = 8
        .Columns("K").ColumnWidth = 8
        .Columns("L").ColumnWidth = 5
        .Columns("M").ColumnWidth = 10

        '----------------------------
        ' Row 1-3: Main Header (NO COLOR)
        '----------------------------
        .Range("A1:M1").Merge
        .Range("A1").Value = boardName
        .Range("A1").Font.Size = 14
        .Range("A1").Font.Bold = True
        .Range("A1").HorizontalAlignment = xlCenter
        .Rows(1).RowHeight = 20

        .Range("A2:M2").Merge
        .Range("A2").Value = examName & " " & _
            examYearEng & " " & GetEngYear() & "/" & _
            examYearBan & " " & GetBanYear() & "/" & _
            examYearHij & " " & GetHijYear()
        .Range("A2").Font.Size = 11
        .Range("A2").Font.Bold = True
        .Range("A2").HorizontalAlignment = xlCenter
        .Rows(2).RowHeight = 18

        .Range("A3:M3").Merge
        .Range("A3").Value = txtRegForm()
        .Range("A3").Font.Size = 12
        .Range("A3").Font.Bold = True
        .Range("A3").HorizontalAlignment = xlCenter
        .Rows(3).RowHeight = 18

        '----------------------------
        ' Row 4: Empty Space
        '----------------------------
        .Rows(4).RowHeight = 3

        '----------------------------
        ' Row 5: Madrasa Name (Full Row)
        '----------------------------
        .Range("A5:M5").Merge
        .Range("A5").Value = GetLblMadrasa() & " " & mName & _
            "  |  " & GetLblCode() & " " & mCode
        .Range("A5").Font.Bold = True
        .Range("A5").Font.Size = 11
        .Range("A5").HorizontalAlignment = xlCenter
        .Range("A5").WrapText = True
        .Rows(5).RowHeight = 20

        '----------------------------
        ' Row 6: Village, Union, Thana, District
        '----------------------------
        .Cells(6, 1).Value = GetLblVillage()
        .Cells(6, 1).Font.Bold = True
        .Cells(6, 1).Font.Size = 9
        .Cells(6, 1).HorizontalAlignment = xlCenter

        .Range("B6:C6").Merge
        .Range("B6").Value = mVillage
        .Range("B6").HorizontalAlignment = xlCenter
        .Range("B6").Font.Size = 10

        .Cells(6, 4).Value = GetLblUnion()
        .Cells(6, 4).Font.Bold = True
        .Cells(6, 4).Font.Size = 9
        .Cells(6, 4).HorizontalAlignment = xlCenter

        .Range("E6:F6").Merge
        .Range("E6").Value = mUnion
        .Range("E6").HorizontalAlignment = xlCenter
        .Range("E6").Font.Size = 10

        .Cells(6, 7).Value = GetLblThana()
        .Cells(6, 7).Font.Bold = True
        .Cells(6, 7).Font.Size = 9
        .Cells(6, 7).HorizontalAlignment = xlCenter

        .Range("H6:I6").Merge
        .Range("H6").Value = mThana
        .Range("H6").HorizontalAlignment = xlCenter
        .Range("H6").Font.Size = 10

        .Cells(6, 10).Value = GetLblDistrict()
        .Cells(6, 10).Font.Bold = True
        .Cells(6, 10).Font.Size = 9
        .Cells(6, 10).HorizontalAlignment = xlCenter

        .Range("K6:M6").Merge
        .Range("K6").Value = mDistrict
        .Range("K6").HorizontalAlignment = xlCenter
        .Range("K6").Font.Size = 10

        .Rows(6).RowHeight = 16

        '----------------------------
        ' Row 7: Jamat, Zone, Phone, Fee
        '----------------------------
        .Cells(7, 1).Value = hdrJamat()
        .Cells(7, 1).Font.Bold = True
        .Cells(7, 1).Font.Size = 9
        .Cells(7, 1).HorizontalAlignment = xlCenter

        '-- Jamat Box (Empty) --
        .Range("B7:C7").Merge
        .Range("B7").Value = ""
        .Range("B7").HorizontalAlignment = xlCenter
        .Range("B7").Borders.LineStyle = xlContinuous

        .Cells(7, 4).Value = hdrZone()
        .Cells(7, 4).Font.Bold = True
        .Cells(7, 4).Font.Size = 9
        .Cells(7, 4).HorizontalAlignment = xlCenter

        .Range("E7:F7").Merge
        .Range("E7").Value = mZone
        .Range("E7").HorizontalAlignment = xlCenter
        .Range("E7").Font.Size = 10

        .Cells(7, 7).Value = GetLblPhone()
        .Cells(7, 7).Font.Bold = True
        .Cells(7, 7).Font.Size = 9
        .Cells(7, 7).HorizontalAlignment = xlCenter

        .Range("H7:I7").Merge
        .Range("H7").Value = mPhone
        .Range("H7").HorizontalAlignment = xlCenter
        .Range("H7").Font.Size = 10

        .Cells(7, 10).Value = GetLblFee()
        .Cells(7, 10).Font.Bold = True
        .Cells(7, 10).Font.Size = 9
        .Cells(7, 10).HorizontalAlignment = xlCenter

        .Range("K7:M7").Merge
        .Range("K7").Value = examFee
        .Range("K7").HorizontalAlignment = xlCenter
        .Range("K7").Font.Size = 10

        .Rows(7).RowHeight = 16

        '-- Row 5:7 Border --
        .Range("A5:M7").Borders.LineStyle = xlContinuous

        '----------------------------
        ' Row 8: Empty Space
        '----------------------------
        .Rows(8).RowHeight = 2

        '----------------------------
        ' Row 9: Table Header (NO COLOR)
        '----------------------------
        .Cells(9, 1).Value = hdrSerial()
        .Cells(9, 1).Font.Size = 9

        .Cells(9, 2).Value = GetLblRegNo()
        .Cells(9, 2).Font.Size = 9

        .Range("C9:D9").Merge
        .Range("C9").Value = hdrName()
        .Range("C9").Font.Size = 9

        .Range("E9:F9").Merge
        .Range("E9").Value = hdrFather()
        .Range("E9").Font.Size = 9

        .Cells(9, 7).Value = GetLblDOB()
        .Cells(9, 7).Font.Size = 8

        .Cells(9, 8).Value = GetLblVillage()
        .Cells(9, 8).Font.Size = 8

        .Cells(9, 9).Value = GetLblPostOffice()
        .Cells(9, 9).Font.Size = 8

        .Cells(9, 10).Value = GetLblThana()
        .Cells(9, 10).Font.Size = 8

        .Cells(9, 11).Value = GetLblDistrict()
        .Cells(9, 11).Font.Size = 8

        .Cells(9, 12).Value = GetLblFee()
        .Cells(9, 12).Font.Size = 8

        .Cells(9, 13).Value = GetLblSignature()
        .Cells(9, 13).Font.Size = 8

        '-- Header Format (Bold, Center, Border ONLY) --
        .Range("A9:M9").Font.Bold = True
        .Range("A9:M9").HorizontalAlignment = xlCenter
        .Range("A9:M9").VerticalAlignment = xlCenter
        .Range("A9:M9").WrapText = True
        .Range("A9:M9").Borders.LineStyle = xlContinuous
        .Rows(9).RowHeight = 25

        '----------------------------
        ' Row 10-19: Empty Table (10 rows)
        ' NO COLOR, Height Reduced
        '----------------------------
        Dim r As Integer
        For r = 10 To 19

            .Cells(r, 1).Value = r - 9  ' Serial
            .Cells(r, 1).HorizontalAlignment = xlCenter
            .Cells(r, 1).Font.Size = 10

            '-- Name Merge --
            .Range(.Cells(r, 3), .Cells(r, 4)).Merge
            '-- Father Merge --
            .Range(.Cells(r, 5), .Cells(r, 6)).Merge

            '-- Border --
            .Range(.Cells(r, 1), .Cells(r, 13)) _
                .Borders.LineStyle = xlContinuous

            '-- Row Height Reduced to 20 --
            .Rows(r).RowHeight = 20

        Next r

        '----------------------------
        ' Row 20: Footer - Only Collector Signature
        '----------------------------
        .Cells(20, 13).Value = GetLblCollector()
        .Cells(20, 13).Font.Bold = True
        .Cells(20, 13).Font.Size = 8
        .Cells(20, 13).HorizontalAlignment = xlRight
        .Cells(20, 13).VerticalAlignment = xlCenter
        .Rows(20).RowHeight = 15

        '----------------------------
        ' Print Setup - সব কিছু এক পেইজে
        '----------------------------
        With .PageSetup
            .Orientation = xlLandscape
            .PaperSize = xlPaperA4
            .FitToPagesWide = 1
            .FitToPagesTall = 1  '← এক পেইজে ফিট করবে
            .PrintTitleRows = "$1:$9"
            .TopMargin = Application.InchesToPoints(0.3)
            .BottomMargin = Application.InchesToPoints(0.3)
            .LeftMargin = Application.InchesToPoints(0.3)
            .RightMargin = Application.InchesToPoints(0.3)
            .HeaderMargin = Application.InchesToPoints(0.2)
            .FooterMargin = Application.InchesToPoints(0.2)
            .CenterHorizontally = True
            .CenterVertically = False
            .PrintHeadings = False
            .PrintGridlines = False
        End With

        '-- Set Print Area --
        .PageSetup.PrintArea = "$A$1:$M$20"

    End With

End Sub

'--------------------------------------------
' Helper Labels - নাম ভেবে যাবে এখান থেকে
' VBA_Helper Sheet-এর মাধ্যমে
'--------------------------------------------

Private Function GetLblMadrasa() As String
    On Error Resume Next
    GetLblMadrasa = Trim(ThisWorkbook.Sheets("VBA_Helper").Range("I2").Value)
    If GetLblMadrasa = "" Then GetLblMadrasa = "মাদরাসা"
    On Error GoTo 0
End Function

Private Function GetLblCode() As String
    On Error Resume Next
    GetLblCode = Trim(ThisWorkbook.Sheets("VBA_Helper").Range("I3").Value)
    If GetLblCode = "" Then GetLblCode = "কোড"
    On Error GoTo 0
End Function

Private Function GetLblVillage() As String
    On Error Resume Next
    GetLblVillage = Trim(ThisWorkbook.Sheets("VBA_Helper").Range("I4").Value)
    If GetLblVillage = "" Then GetLblVillage = "গ্রাম"
    On Error GoTo 0
End Function

Private Function GetLblUnion() As String
    On Error Resume Next
    GetLblUnion = Trim(ThisWorkbook.Sheets("VBA_Helper").Range("I5").Value)
    If GetLblUnion = "" Then GetLblUnion = "ইউনিয়ন"
    On Error GoTo 0
End Function

Private Function GetLblThana() As String
    On Error Resume Next
    GetLblThana = Trim(ThisWorkbook.Sheets("VBA_Helper").Range("I6").Value)
    If GetLblThana = "" Then GetLblThana = "থানা"
    On Error GoTo 0
End Function

Private Function GetLblDistrict() As String
    On Error Resume Next
    GetLblDistrict = Trim(ThisWorkbook.Sheets("VBA_Helper").Range("I7").Value)
    If GetLblDistrict = "" Then GetLblDistrict = "জেলা"
    On Error GoTo 0
End Function

Private Function GetLblPhone() As String
    On Error Resume Next
    GetLblPhone = Trim(ThisWorkbook.Sheets("VBA_Helper").Range("I8").Value)
    If GetLblPhone = "" Then GetLblPhone = "ফোন"
    On Error GoTo 0
End Function

Private Function GetLblFee() As String
    On Error Resume Next
    GetLblFee = Trim(ThisWorkbook.Sheets("VBA_Helper").Range("I9").Value)
    If GetLblFee = "" Then GetLblFee = "বেতন"
    On Error GoTo 0
End Function

Private Function GetLblRegNo() As String
    On Error Resume Next
    GetLblRegNo = Trim(ThisWorkbook.Sheets("VBA_Helper").Range("I10").Value)
    If GetLblRegNo = "" Then GetLblRegNo = "রেজ নং"
    On Error GoTo 0
End Function

Private Function GetLblDOB() As String
    On Error Resume Next
    GetLblDOB = Trim(ThisWorkbook.Sheets("VBA_Helper").Range("I11").Value)
    If GetLblDOB = "" Then GetLblDOB = "জন্ম"
    On Error GoTo 0
End Function

Private Function GetLblPostOffice() As String
    On Error Resume Next
    GetLblPostOffice = Trim(ThisWorkbook.Sheets("VBA_Helper").Range("I12").Value)
    If GetLblPostOffice = "" Then GetLblPostOffice = "ডাকঘর"
    On Error GoTo 0
End Function

Private Function GetLblSignature() As String
    On Error Resume Next
    GetLblSignature = Trim(ThisWorkbook.Sheets("VBA_Helper").Range("I13").Value)
    If GetLblSignature = "" Then GetLblSignature = "স্বাক্ষর"
    On Error GoTo 0
End Function

Private Function GetLblCollector() As String
    On Error Resume Next
    GetLblCollector = Trim(ThisWorkbook.Sheets("VBA_Helper").Range("I14").Value)
    If GetLblCollector = "" Then GetLblCollector = "সংগ্রহকারী"
    On Error GoTo 0
End Function

Private Function GetEngYear() As String
    On Error Resume Next
    GetEngYear = Trim(ThisWorkbook.Sheets("VBA_Helper").Range("I18").Value)
    If GetEngYear = "" Then GetEngYear = "ইং"
    On Error GoTo 0
End Function

Private Function GetBanYear() As String
    On Error Resume Next
    GetBanYear = Trim(ThisWorkbook.Sheets("VBA_Helper").Range("I19").Value)
    If GetBanYear = "" Then GetBanYear = "বাং"
    On Error GoTo 0
End Function

Private Function GetHijYear() As String
    On Error Resume Next
    GetHijYear = Trim(ThisWorkbook.Sheets("VBA_Helper").Range("I20").Value)
    If GetHijYear = "" Then GetHijYear = "হিজ"
    On Error GoTo 0
End Function

Private Function GetFeeUnit() As String
    On Error Resume Next
    GetFeeUnit = Trim(ThisWorkbook.Sheets("VBA_Helper").Range("I21").Value)
    If GetFeeUnit = "" Then GetFeeUnit = "টাকা"
    On Error GoTo 0
End Function

'--------------------------------------------
' Dashboard Button
'--------------------------------------------
Public Sub Btn_BlankRegistration()

    Dim wsMadrasa As Worksheet
    Set wsMadrasa = ThisWorkbook.Sheets("Madrasa_List")

    Dim madrasaCount As Long
    madrasaCount = wsMadrasa.Cells(Rows.Count, 1).End(xlUp).Row - 1

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

End Sub

