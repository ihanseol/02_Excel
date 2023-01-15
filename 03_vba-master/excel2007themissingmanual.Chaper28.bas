Attribute VB_Name = "Chaper28"
Sub formatrow()
    Application.Workbooks("excel2007themissingmanual.xlsm").Worksheets("28").Activate
    'ActiveCell.Rows.EntireRow.Select  'select an entire row
    'Range(Selection, Selection.End(xlToRight)).Select 'select row contiguous cells
    Selection.Interior.ColorIndex = 35
    Selection.Interior.Pattern = xlSolid
    'also with statement
    'With Selection.Interior
    '    .ColorIndex = 35
    '    .Pattern = xlSolid
    'End With

    ActiveCell.Value = "Hello World"
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = "Add me " & ActiveCell.Offset(-1, 0).Value
    Range("E1").Value = 5
    Range("E2").Value = (Range("E1").Value * 2) - 1
    Range("E4:E14").Value = "Hello"              'print Hello cells E4:E14
    Range("F1").Value = "Excel file created by " & Application.UserName
End Sub

Sub formatcells()
    Application.Workbooks("excel2007themissingmanual.xlsm").Worksheets("28").Activate
    Range("A20:A24").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "Arial"
        .FontStyle = "Bold"
        .Size = 14
    End With

    With Range("E4:E10").Font
        .Name = "Arial"
        .FontStyle = "Bold"
        .Size = 14
    End With
End Sub

Sub selectrelativecells()
    Application.Workbooks("excel2007themissingmanual.xlsm").Worksheets("28").Activate
    'select a relative group of cells
    'select current cell and the cell immediately to the right
    ActiveCell.Range("A1:A2").Select
End Sub

Function commissionbonus(sales) As Double
    'ifs formula in Excel may be better
    If sales > 1000 Then
        commissionbonus = 100
    ElseIf sales > 500 Then
        commissionbonus = 50
    ElseIf sales > 100 Then
        commissionbonus = 10
    Else
        commissionbonus = 1
    End If
End Function

Sub dountilactivecellblank()
    Application.Workbooks("excel2007themissingmanual.xlsm").Worksheets("28").Activate
    
    'No do until
    Range("K1", Range("K1").End(xlDown)).Select
    With Selection.Interior
        .ColorIndex = 35
        .Pattern = xlSolid
    End With
    
    Selection.ClearFormats
    'yes do until
    
    Range("K1").Select
    Do Until ActiveCell.Value = ""
        With Selection.Interior
            .ColorIndex = 35
            .Pattern = xlSolid
        End With
        ActiveCell.Offset(1, 0).Select
    Loop
End Sub

Sub easyforloopselectioncells()
    Application.Workbooks("excel2007themissingmanual.xlsm").Worksheets("28").Activate
    
    Range("K1", Range("K1").End(xlDown)).Select
    For Each cell In Selection
        Selection.Font.FontStyle = "Bold"
    Next
End Sub

