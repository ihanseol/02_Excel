Attribute VB_Name = "Module1"
Sub MacroNoRelative()
    '
    ' MacroNoRelative Macro
    '

    '
    Range("C1").Select
    Selection.FormulaR1C1 = "42"
End Sub

Sub MacroYesRelative()
    '
    ' MacroYesRelative Macro
    '

    '
    ActiveCell.Offset(0, 2).Select
    Selection.FormulaR1C1 = "42"
End Sub

Sub InsertHeader()
    '
    ' InsertHeader Macro
    '

    '
    Range("A1").Value = "Sales Report"
    Range("A1:C1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Range("A2").Select
    Selection.FormulaR1C1 = "=TODAY()"
End Sub

Sub Macro5()
    '
    ' Macro5 Macro
    '

    '
    ActiveCell.Range("A1:E1").Select
End Sub

Sub Macro6()
    '
    ' Macro6 Macro
    '

    '
    ActiveCell.Offset(5, 0).Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
End Sub

Sub Macro7()
    '
    ' Macro7 Macro
    '

    '
    Selection.ClearFormats
End Sub

