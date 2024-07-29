Attribute VB_Name = "Module3"
Option Explicit 'Mail Merge to Outlook
Function FileExists(TenFile As String) As Boolean
    Dim path As String
    Dim filePath As String
    path = "C:\Users\Sammy\OneDrive - Industrial University of HoChiMinh City\Desktop\Mail Merge in Excel\Files Attachment"
    filePath = path + "\" & TenFile & ".pdf"
    FileExists = (Dir(filePath) <> "")
End Function
Sub CapNhatCongThuc()
    Range("K2:K65000").ClearContents
    Range("L2:L65000").ClearContents
    Dim lr1 As Long
    Dim lr2 As Long
    lr1 = Range("I" & Rows.Count).End(xlUp).Row
    lr2 = Range("J" & Rows.Count).End(xlUp).Row
    Range("K2:K" & lr1).Formula = "=FileExists(RC[-2])"
    Range("L2:L" & lr2).Formula = "=FileExists(RC[-2])"
End Sub
Function LayDuongDan()
    LayDuongDan = "C:\Users\Sammy\OneDrive - Industrial University of HoChiMinh City\Desktop\Mail Merge in Excel\Files Attachment"
End Function
Sub ChonThuMau(ByRef control As Office.IRibbonControl)
    Dim FD As FileDialog
    Set FD = Application.FileDialog(msoFileDialogFilePicker)
    Dim strFileName As String

    With FD
        .AllowMultiSelect = False
        .InitialFileName = "C:\Users\Sammy\OneDrive - Industrial University of HoChiMinh City\Desktop\Mail Merge in Excel\Outlook Templates"
        .Title = "Please select one file"
        .Filters.Clear
        .Filters.Add "All files", "*.*"

        If .Show = True Then
            strFileName = Dir(.SelectedItems(1))
            Range("A2").Value = strFileName
        Else
            Range("A2").Value = "No files selected."
        End If
    End With
End Sub
Sub GuiMailHangLoat(ByRef control As Office.IRibbonControl)
    'Tao ten A1:G1
    Rows("1:1").RowHeight = 25
    Columns("A:A").ColumnWidth = 20
    Columns("B:B").ColumnWidth = 8
    Columns("C:C").ColumnWidth = 11
    Columns("D:D").ColumnWidth = 23
    Columns("E:E").ColumnWidth = 20
    Columns("F:F").ColumnWidth = 29
    Columns("G:G").ColumnWidth = 29
    Columns("H:H").ColumnWidth = 29
    Columns("I:I").ColumnWidth = 13
    Columns("J:J").ColumnWidth = 13
    Columns("K:K").ColumnWidth = 13
    Columns("L:L").ColumnWidth = 13
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "Outlook Template"
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "STT"
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "Subject"
    Range("D1").Select
    ActiveCell.FormulaR1C1 = "Name"
    Range("E1").Select
    ActiveCell.FormulaR1C1 = "MSSV"
    Range("F1").Select
    ActiveCell.FormulaR1C1 = "Mail To"
    Range("G1").Select
    ActiveCell.FormulaR1C1 = "CC"
    Range("H1").Select
    ActiveCell.FormulaR1C1 = "BCC"
    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Attach File 1"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Attach File 2"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "File 1 Check"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "File 2 Check"
        'To dam va mau
    With Range("A1:L1").Interior
        .Pattern = xlSolid
        .ColorIndex = 22
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
    With Range("A1:L1").Font
        .Bold = True
    End With
    Range("A1:L1").HorizontalAlignment = xlLeft
    Range("A1:L1").VerticalAlignment = xlCenter

    ' Luu b?ng tính
    ' ThisWorkbook.Save
    ' Cap nhat lai bang tinh
    ActiveSheet.UsedRange
    'Co dinh hang 1 va gian dong cot F
    Range("A1").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.AutoFilter
    Rows("2:2").Select
    ActiveWindow.FreezePanes = True
    ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, 400, 100, 100, _
        25).Select
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "Check Attachment Existence"
    Selection.OnAction = "CapNhatCongThuc"
End Sub
Sub LuuMail()
CapNhatCongThuc
If MsgBox("Confirm Saving?", vbYesNo) = vbYes Then
    Dim olApp As Object
    Dim olMail As Object
    Dim i As Long, lr As Long
    Dim ThuMau As String
    Dim MailNhan As String
    Dim TieuDe As String
    Dim CC As String
    Dim BCC As String
    Dim TenFile1 As String
    Dim TenFile2 As String
    Dim FileDinhKem1 As String
    Dim FileDinhKem2 As String
    Dim path As String
    Dim olInsp As Object
    Dim wdDoc As Object
    Dim oRng As Object
    Dim Name As String
    Dim MSSV As String

    With ActiveSheet
        On Error Resume Next
        .ShowAllData
        On Error GoTo 0

        lr = .Range("F" & Rows.Count).End(xlUp).Row 'dong cuoi

        For i = 2 To lr
            ThuMau = .Range("A2").Value
            If .Range("F" & i).Value <> "" Then 'neu co file dinh kem
                Set olApp = CreateObject("Outlook.Application") 'gan bien
                Set olMail = olApp.CreateItemFromTemplate("C:\Users\Sammy\OneDrive - Industrial University of HoChiMinh City\Desktop\Mail Merge in Excel\Outlook Templates" + "\" & ThuMau) 'goi thu mau
                olMail.Display
                MailNhan = .Range("F" & i).Value
                TieuDe = .Range("C" & i).Value
                CC = .Range("G" & i).Value
                BCC = .Range("H" & i).Value
                TenFile1 = .Range("I" & i).Value
                TenFile2 = .Range("J" & i).Value
                Name = .Range("D" & i).Value
                MSSV = .Range("E" & i).Value
                If .Range("K" & i).Value = True Then
                    path = "C:\Users\Sammy\OneDrive - Industrial University of HoChiMinh City\Desktop\Mail Merge in Excel\Files Attachment"
                    FileDinhKem1 = path + "\" & TenFile1 & ".pdf" 'duong dan file dinh kem
                    FileDinhKem2 = path + "\" & TenFile2 & ".pdf"
                End If
                With olMail
                    'Cau truc gui mail
                    .To = MailNhan ' Email nguoi nhan
                    .Subject = TieuDe ' Tieu de mail
                    .CC = CC
                    .BCC = BCC
                    If Len(Dir(FileDinhKem1)) > 0 Then .Attachments.Add FileDinhKem1 ' Ki?m tra và dính kèm file n?u t?n t?i
                    If Len(Dir(FileDinhKem2)) > 0 Then .Attachments.Add FileDinhKem2
                    Set olInsp = .GetInspector
                    Set wdDoc = olInsp.WordEditor
                    Set oRng = wdDoc.Range
                    With oRng.Find
                        Do While .Execute(FindText:="{{Name}}")
                            oRng.Text = Name
                            Exit Do
                        Loop ' Thay the tat ca các {{Name}} bang giá tri cua Name
                        Do While .Execute(FindText:="{{MSSV}}")
                            oRng.Text = MSSV
                            Exit Do
                        Loop ' Thay th? t?t c? các {{Name}} b?ng giá tr? c?a Name
                    End With
                    .Save 'lenh luu mail
                End With
                Set olMail = Nothing
                Set wdDoc = Nothing
                Set oRng = Nothing
            End If
        Next i
        Set olApp = Nothing
        MsgBox "Done!"
    End With
End If
End Sub
Sub GuiMail()
CapNhatCongThuc
If MsgBox("Confirm Sending?", vbYesNo) = vbYes Then
    Dim olApp As Object
    Dim olMail As Object
    Dim i As Long, lr As Long
    Dim ThuMau As String
    Dim MailNhan As String
    Dim TieuDe As String
    Dim CC As String
    Dim BCC As String
    Dim TenFile1 As String
    Dim TenFile2 As String
    Dim FileDinhKem1 As String
    Dim FileDinhKem2 As String
    Dim path As String
    Dim olInsp As Object
    Dim wdDoc As Object
    Dim oRng As Object
    Dim Name As String
    Dim MSSV As String

    With ActiveSheet
        On Error Resume Next
        .ShowAllData
        On Error GoTo 0

        lr = .Range("F" & Rows.Count).End(xlUp).Row 'dong cuoi

        For i = 2 To lr
            ThuMau = .Range("A2").Value
            If .Range("F" & i).Value <> "" Then 'neu co file dinh kem
                Set olApp = CreateObject("Outlook.Application") 'gan bien
                Set olMail = olApp.CreateItemFromTemplate("C:\Users\Sammy\OneDrive - Industrial University of HoChiMinh City\Desktop\Mail Merge in Excel\Outlook Templates" + "\" & ThuMau) 'goi thu mau
                olMail.Display
                MailNhan = .Range("F" & i).Value
                TieuDe = .Range("C" & i).Value
                CC = .Range("G" & i).Value
                BCC = .Range("H" & i).Value
                TenFile1 = .Range("I" & i).Value
                TenFile2 = .Range("J" & i).Value
                Name = .Range("D" & i).Value
                MSSV = .Range("E" & i).Value
                If .Range("K" & i).Value = True Then
                    path = "C:\Users\Sammy\OneDrive - Industrial University of HoChiMinh City\Desktop\Mail Merge in Excel\Files Attachment"
                    FileDinhKem1 = path + "\" & TenFile1 & ".pdf" 'duong dan file dinh kem
                    FileDinhKem2 = path + "\" & TenFile2 & ".pdf"
                End If
                With olMail
                    'Cau truc gui mail
                    .To = MailNhan ' Email nguoi nhan
                    .Subject = TieuDe ' Tieu de mail
                    .CC = CC
                    .BCC = BCC
                    If Len(Dir(FileDinhKem1)) > 0 Then .Attachments.Add FileDinhKem1 ' Ki?m tra và dính kèm file n?u t?n t?i
                    If Len(Dir(FileDinhKem2)) > 0 Then .Attachments.Add FileDinhKem2
                    Set olInsp = .GetInspector
                    Set wdDoc = olInsp.WordEditor
                    Set oRng = wdDoc.Range
                    With oRng.Find
                        Do While .Execute(FindText:="{{Name}}")
                            oRng.Text = Name
                            Exit Do
                        Loop ' Thay the tat ca các {{Name}} bang giá tri cua Name
                        Do While .Execute(FindText:="{{MSSV}}")
                            oRng.Text = MSSV
                            Exit Do
                        Loop ' Thay th? t?t c? các {{Name}} b?ng giá tr? c?a Name
                    End With
                    .Send 'lenh luu mail
                End With
                Set olMail = Nothing
                Set wdDoc = Nothing
                Set oRng = Nothing
            End If
        Next i
        Set olApp = Nothing
        MsgBox "Sent!"
    End With
End If
End Sub


