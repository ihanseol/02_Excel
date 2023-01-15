Attribute VB_Name = "Chapter04"
Sub objectsandtheirmethods()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("4").Activate
    Range("A1").Select
    Range("A2").Value = "abcdefghijklmnop"
    Range("A2").ClearContents
    Range("A3").Value = "qrstuvwyxz"
    Range("A3").Copy
    Range("A4").PasteSpecial xlPasteValues
    Range("A5").Value = "cut"
    'Range("A5").Cut
    'Range("A6").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    Worksheets("3").Select
    Worksheets("4").Select
    Worksheets.Add before:=Worksheets(1)         'add sheet farthest left or add sheet left worksheet index 1
    ActiveSheet.name = "firstworksheet"
    Worksheets.Add After:=Sheets(Sheets.Count)   'add sheet farthest right
    ActiveSheet.name = "lastworksheet"
    Worksheets("firstworksheet").Delete
    Application.DisplayAlerts = False
    Worksheets("lastworksheet").Delete
    Application.DisplayAlerts = True
End Sub

Sub moreactives()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("4").Activate
    Range("B1").Select
    Range("B1").Value = "Yes"
    If ActiveCell.Value = "Yes" Then
        Range("B2").Value = "Yes, too"
    End If
    Range("B7").Select
    ActiveCell.Value = 4
    ActiveWorkbook.Worksheets("4").Range("B10").Value = 1
    ActiveCell.Offset(1, 0).Formula = "=sum(b7,b10)" 'cell B8 is =SUM(B7,B10)
    Range("B11").Formula = "=sum(b7:b10)"        'cell b11 is =SUM(B7:B10)
    Range("C1").Select
    ActiveCell.Offset(1, 0).Select               'active cell focus is cell C2
    ActiveCell.Offset(1, 0).Select               'active cell focus is cell C3
    ActiveCell.Offset(1, 0).Select               'active cell focus is cell C4
    ActiveCell.Offset(1, 0).Select               'active cell focus is cell C5
    Range("A3").Copy
    ActiveCell.Offset(1, 0).PasteSpecial xlPasteValues 'paste A3 to cell C6, yet active cell _
                                                       focus is cell C5
    Application.CutCopyMode = False
    Range("D1").Value = ActiveCell.Address       'print $C$6
    Range("D2").Select
End Sub

Sub withendwith()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("4").Activate
    Range("D4").Select
    With Selection.Interior
        .ColorIndex = 3
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
    End With
End Sub

Sub foreachnextloopquickintro()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("4").Activate
    counter = 1
    For Each Cell In Range("F1:F10")
        Range("F" & counter).Value = counter
        counter = counter + 1
    Next
End Sub

Sub chapter4vbacodecombine()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("4").Activate
    Dim frange As Range
    Set frange = Range("F1:F20")
    For Each Cell In frange
        Cell.Value = 1                           'it works.  Cells F1:F20 value is 1
    Next
    counter = 1
    For Each Cell In frange
        Range("F" & counter).Value = counter
        counter = counter + 1
    Next
    Range("F9").Value = "abc"
    For Each Cell In frange                      'value is non number color is red
        If Not IsNumeric(Cell) Then
            MsgBox "Please enter a number in cell " & Cell.Address
            With Cell.Interior
                .ColorIndex = 3
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
            End With
        ElseIf Cell.Value Mod 2 = 0 Then         'value is even color is green
            With Cell.Interior
                .ColorIndex = 4
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
            End With
        End If
    Next
End Sub

Sub randomandsort()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("4").Activate
    Dim randomrange As Range
    Set randomrange = Range("H1:H10")
    For Each Cell In randomrange:
        Cell.Value = Int(Rnd() * 20 + 1)
    Next
    'sort simple sorting
    randomrange.Sort key1:=randomrange
End Sub

Sub insertrow()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("4").Activate
    Range("A30").Value = "dog fox"
    Range("A31").Value = "cat cow"
    Range("A32").Value = "fish bear"
    Range("A31").Select
    Selection.EntireRow.Insert                   'insert row at row 31 cursor at cell A31
    Selection.EntireRow.Delete                   'delete row at row 31 cursor at cell A31
End Sub


