Attribute VB_Name = "deletemovenames"
Sub warmup()
    Application.Workbooks("names.xlsm").Worksheets("Sheet1").Activate
    Dim countrows As Integer
    Range("A1").Select
    countrows = Range("A1", Range("A1").End(xlDown)).Count
    Range("B1").Value = countrows
    For Each cell In Range("A1:A9")
        If cell.Value = "Raymond" Then
            Selection.Font.Bold = True
        End If
        cell.Offset(1, 0).Select
    Next
    Range("A1").Select
    For Each cell In Range("A1:A" & countrows)
        If cell.Value = "Raymond" Then
            Selection.Font.Color = vbBlue
        End If
        ActiveCell.Offset(1, 0).Select
    Next
End Sub

Sub extramovedelete()
    Application.Workbooks("names.xlsm").Worksheets("Sheet1").Activate
    Dim countrows As Integer
    countrows = Range("A1", Range("A1").End(xlDown)).Count
    For Each cell In Range("A1:A" & countrows)
        If cell.Value = "Raymond" Then
        End If
    Next
End Sub

Sub extramovedelete2()
    Application.Workbooks("names.xlsm").Worksheets("Sheet1").Activate
    Dim countrows As Integer
    countrows = Range("A1", Range("A1").End(xlDown)).Count
    Range("A1").Select
    For n = countrows To 1 Step -1
        'move rows with raymond to worksheet 2
        If InStr(1, Cells(n, 1), "Raymond") > 0 Then
            Rows(n).Copy
            Worksheets("raymond").Activate
            ActiveCell.PasteSpecial xlPasteAll
            'Application.CutCopyMode = False
            ActiveCell.Offset(1, 0).Select
        ElseIf InStr(1, Cells(n, 1), "James") > 0 Then
            Rows(n).Copy
            Worksheets("james").Activate
            ActiveCell.PasteSpecial xlPasteAll
            'Application.CutCopyMode = False
            ActiveCell.Offset(1, 0).Select
        ElseIf InStr(1, Cells(n, 1), "Michelle") > 0 Then
            Rows(n).Copy
            Worksheets("michelle").Activate
            ActiveCell.PasteSpecial xlPasteAll
            'Application.CutCopyMode = False
            ActiveCell.Offset(1, 0).Select
        End If
        Worksheets(1).Activate
        Rows(n).Delete
    Next
End Sub

Sub extramovedelete3()
    Application.Workbooks("names.xlsm").Worksheets("Sheet1").Activate
    Dim countrows As Integer, column As Integer
    countrows = Range("A1", Range("A1").End(xlDown)).Count
    column = 6
    Range("A1").Select
    For n = countrows To 1 Step -1
        'move rows with raymond to worksheet 2
        If InStr(1, Cells(n, column), "Raymond") > 0 Then
            Rows(n).Copy
            Worksheets("raymond").Activate
            ActiveCell.PasteSpecial xlPasteAll
            'Application.CutCopyMode = False
            ActiveCell.Offset(1, 0).Select
        ElseIf InStr(1, Cells(n, column), "James") > 0 Then
            Rows(n).Copy
            Worksheets("james").Activate
            ActiveCell.PasteSpecial xlPasteAll
            'Application.CutCopyMode = False
            ActiveCell.Offset(1, 0).Select
        ElseIf InStr(1, Cells(n, column), "Michelle") > 0 Then
            Rows(n).Copy
            Worksheets("michelle").Activate
            ActiveCell.PasteSpecial xlPasteAll
            'Application.CutCopyMode = False
            ActiveCell.Offset(1, 0).Select
        End If
        Worksheets(1).Activate
        Rows(n).Delete
    Next
End Sub

