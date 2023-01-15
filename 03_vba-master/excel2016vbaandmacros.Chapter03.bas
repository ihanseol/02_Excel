Attribute VB_Name = "Chapter03"
Sub ranges()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("3").Activate
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("3").Range("A5").Value = "Range Hierarchy"
    Worksheets(2).Range("B6").Value = "Worksheet 3 Is the second worksheet Then another way reference Range"
    Range("B5").Range("C3").Select        'Range D7 selected
    'Range("B5").Range("C3").Select selects cell D7. Think about cell C3, which is located two rows below
    'and two columns to the right of cell A1. The preceding line of code starts at cell B5. If we
    'assume that B5 is in the A1 position, VBA finds the cell that would be in the C3 position
    'relative to B5. In other words, VBA finds the cell that is two rows below and two columns
    'to the right of B5, which is D7.
    Range("A1", "B5").Select
    'same as
    Range("A1:B5").Select
    Range("C9").Select
    Range(ActiveCell, "F5").Select
    Range("D1").Select
    Range(ActiveCell, ActiveCell.Offset(2, 4)).Select
    Range("A12:B14, D12:F15").Select        'select multiple cell ranges
    Range("A1").Select
    Range("E4").Select
    Range("A1").Select
    Range(ActiveCell, Range("E4")).Select        'highlight cells A1:E4
End Sub

Sub cellscells()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("3").Activate
    Cells(1, 1).Select
    'same as
    Cells(1, "A").Select
    finalrowe = Cells(Rows.Count, "E").End(xlUp).Row
    For i = 1 To finalrowe Step 1
        Cells(i, "E").Font.Bold = True
    Next i
    Range(Cells(1, 1), Cells(5, 5)).Select        'highlight cells A1:E5
End Sub


Sub offsets()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("3").Activate
    'Range.Offset(RowOffset, ColumnOffset)Range("A1").Select
    Range("A1").Offset(10, 2).Select
    Range("A1").Offset(0, 1).Select
    Range("A1:C3").Offset(1, 1).Select        'highlight cells A1:C3 shifted one row down one column right
    With Range("B20:B23")
        Set pricecolumn = .Find(What:="1", lookat:=xlWhole, LookIn:=xlValues)
        If Not pricecolumn Is Nothing Then
            firstAddress = pricecolumn.Address
            Do
                pricecolumn.Offset(0, 1).Value = "Low"
                Set pricecolumn = .FindNext(pricecolumn)
            Loop While Not pricecolumn Is Nothing And pricecolumn.Address <> firstAddress
        End If
    End With
    Range("A20").CurrentRegion.Select        'highlight all fruits cells
End Sub


Sub isemptycell()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("3").Activate
    LastRow = 10
    For i = 1 To LastRow
        If IsEmpty(Range("E" & i)) Then
            Range("F" & i).Value = "Empty"
        End If
    Next i
End Sub


Sub referencetables()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("3").Activate
    Worksheets("3").ListObjects("tableH10J13").Range.Select        'doesn't work
End Sub


