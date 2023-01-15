Attribute VB_Name = "forloop"
Sub quickguideforloop()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("forloop").Activate
    Dim i As Long
    For i = 1 To 20
        Cells(i, 4).Value = i
    Next i
    For i = 2 To 10 Step 2
    Next
    For i = 10 To 1 Step -1
    Next
    Dim ncollection As Range
    Set ncollection = Range("A1:A10")
    For Each n In ncollection
        n.Value = "okay"
    Next n
    Range("B6").Value = "found"
    For i = 1 To 10
        If Cells(i, 2) = "found" Then
            Exit For
        Else
            Cells(i, 2) = "somewhere else"
        End If
    Next i
End Sub

Sub forloop()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("forloop").Activate
    Dim firstrow, lastrow, orangessold As Long
    firstrow = 22
    lastrow = Range("A" & firstrow, Range("A" & firstrow).End(xlDown)).count
    orangessold = 0
    For i = firstrow To (firstrow + lastrow)
        Range("A" & i).Select
        If Range("A" & i).Value = "Oranges" Then
            orangessold = orangessold + Range("B" & i).Value
        End If
    Next i
    Range("A42") = orangessold & " oranges sold."
End Sub

Sub exitforloop()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("forloop").Activate
    Dim firstrow, lastrow, orangessold As Long
    firstrow = 22
    lastrow = Range("A" & firstrow, Range("A" & firstrow).End(xlDown)).count
    orangessold = 0
    For i = firstrow To (firstrow + lastrow)
        Range("A" & i).Select
        If Range("A" & i).Value = "Oranges" Then
            orangessold = orangessold + Range("B" & i).Value
        End If
        If Range("A" & i).Value = "Pears" Then
            MsgBox "Pears are not for me"
            Exit For
        End If
    Next i
    Range("A42") = orangessold & " oranges sold."
End Sub

Sub collectionforloop()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("forloop").Activate
    'The For loop can read items in a Collection.  Example below is display open workbooks
    Dim i As Long
    Range("D22").Select
    For i = 1 To Workbooks.count
        ActiveCell.Value = Workbooks(i).FullName 'print G:\Raymond\Excel Files 2GB Backup 072218\VBA Macros Round Two\excelmacromastery.xlsm
        ActiveCell.offset(1, 0).Select
        ActiveCell.Value = Workbooks(i).name     'print excelmacromastery.xlsm
        ActiveCell.offset(1, 0).Select
    Next i
End Sub

Sub nestedforloop()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("forloop").Activate
    Dim i, j As Long
    Range("E22").Select
    For i = 1 To Workbooks.count
        For j = 1 To Workbooks(i).Worksheets.count
            ActiveCell.Value = Workbooks(i).FullName + ":" + Worksheets(j).name
            ActiveCell.offset(1, 0).Select
        Next j
    Next i
End Sub

Sub foreachloop()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("forloop").Activate
    'For Each loop read items from a collection or an array; e.g. access open workbooks
    'Dim openworkbooks As Workbook
    'For Each openworkbooks In Workbooks
    '    MsgBox openworkbooks.FullName
    '    MsgBox openworkbooks.name
    'Next openworkbooks

    'Another collection example is sheets
    Range("F22").Select
    Dim allsheets As Variant
    For Each allsheets In ThisWorkbook.Sheets
        ActiveCell.Value = allsheets.name
        ActiveCell.offset(1, 0).Select
    Next allsheets
    'RM:  skipped arrays
End Sub

Sub nestedforeachloop()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("forloop").Activate
    Dim i As Long, j As Long
    Range("G22").Select
    ' First Loop goes through all workbooks
    For i = 1 To Workbooks.count
        ' Second loop goes through all the worksheets of workbook(i)
        For j = 1 To Workbooks(i).Worksheets.count
            ActiveCell.Value = Workbooks(i).name + ": " + Worksheets(j).name
            ActiveCell.offset(1, 0).Select
        Next j
    Next i

    Range("H22").Select
    Dim wk As Workbook, sh As Worksheet
    ' Read each workbook
    For Each wk In Workbooks
        ' Read each worksheet in the wk workbook
        For Each sh In wk.Worksheets
            ' Print workbook name and worksheet name
            ActiveCell.Value = wk.name + ": " + sh.name
            ActiveCell.offset(1, 0).Select
        Next sh
    Next wk
End Sub

