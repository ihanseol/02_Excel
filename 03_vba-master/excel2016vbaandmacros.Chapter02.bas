Attribute VB_Name = "Chapter02"
Sub selectionprimer()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("2").Activate
    Range("C1").Clear
    Range("J1", Range("J1").End(xlDown)).Clear
    
    Range("A1").Select
    selection.End(xlDown).Select
    Range("A1").Select
    selection.End(xlDown).Offset(1, 0).Select
    Range("A1").Select
    ActiveCell.End(xlDown).Select
    Range("A1").Select
    Range("A1", Range("A1").End(xlDown)).Select        'highlight all cells
    Range("A1").Select
    ActiveCell.Offset(0, 1).Select
    Range("A1").Select
    Range("A1").Copy Range("C1")
    Range("A1").Select
    Range("A1", Range("A1").End(xlDown)).Copy
    Application.CutCopyMode = False
    Range("K1").Select
    Range("A1", Range("A1").End(xlDown)).Copy Range("J1")
End Sub


Sub withendwithfontintroduction()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("2").Activate
    Range("A31:A35").Value = "abc"
    With Range("A31:A35").Font
        .Bold = True
        .Size = 12
        .Name = Arial
    End With
    'or
    Range("C31:C35").Select
    With selection.Font
        .Bold = True
        .Size = 24
        .Name = TimesNewRoman
    End With
    selection.ClearFormats
End Sub
