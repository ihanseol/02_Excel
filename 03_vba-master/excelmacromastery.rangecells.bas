Attribute VB_Name = "rangecells"
Sub basicsrangecells()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("rangecells").Activate
    Range("A1") = "Range(""A1"")"
    Range("A2:A4") = "Range(""A2:A4"")"
    Range("A5:C5, D1:D2") = "Range(""A5:C5, D1:D2"")"
    Cells(6, 2) = "Cells(6,2) row six column two"
    Rows(8) = "rows(8) rows 8"
    Columns(5) = "columns(5) column 5"
    Rows("10:12") = "rows(""10:12"") rows 10, 11, 12"
    Columns("I:K") = "columns(""I:K"") columns i, j, k"
    'clear contents lower memory size
    Range("E13", Range("E13").End(xlDown)).ClearContents
    Range("I13", Range("I13").End(xlDown).End(xlToRight)).ClearContents
    Range("N8", Range("N8").End(xlToRight)).ClearContents
    Range("N10", Range("N10").End(xlDown).End(xlToRight)).ClearContents
End Sub

Sub writetocell()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("rangecells").Activate
    'write a date
    ThisWorkbook.Worksheets("rangecells").Range("A14") = #11/21/2017#
    'write multiple cells
    Range("A16:C20") = "John Smith"
    'write multiple cells different areas multiple locations
    Range("A22:A24, C22:D36") = "John Smith II"
    'use Cells() in Range()
    Range(Cells(38, 1), Cells(40, 3)) = "A38 to C40"
    'print cell address
    Range("A42") = Range("A42").Address          'print $A$42
End Sub

Sub copypaste()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("rangecells").Activate
    Range("D22").Copy Destination:=Range("F22")  'print John Smith II cell F22
    Range("C38").Copy
    Range("F38").PasteSpecial xlPasteValues      'print A38 to C40
    Range("C39").Font.Bold = True
    Range("C39").Copy
    Range("H38").PasteSpecial xlPasteAll         'print A38 to C40 bold
    Range("I38").PasteSpecial xlPasteValues      'print A38 to C40 not bold
    Application.CutCopyMode = False              'stop copy or stop dancing ants
End Sub

Sub formatcells()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("rangecells").Activate
    Range("A44").Value = "Range A44"
    Range("A44").Font.Bold = True
    Range("A44").Font.Underline = True
    Range("A44").Font.Color = rgbNavy
    Range("A45").Value = 2
    Range("A45").NumberFormat = "0.00"           'display cell A45 2.00
    Range("A46").Value = #8/27/2018#
    Range("A46").NumberFormat = "mm/dd/yy"       'display cell 08/27/18
    Range("A46").Interior.Color = rgbysandybrown
    Range("A46").Font.Color = rgbWhite
    Range("A46").Borders.LineStyle = xlDash
    Range("A46").Borders.Color = rgbYellow
    Range("A46").Borders.Weight = xlThick
    'Range("A47").numberformat = "General"
    'Range("A47").numberformat = "Text"
End Sub

