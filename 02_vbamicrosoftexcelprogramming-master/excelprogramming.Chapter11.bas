Attribute VB_Name = "Chapter11"
Sub rangeproperty()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("11").Activate
    Range("A1:A3").Value = "Range A1:A3"
    Range("B11", "D11").Value = "Range B11, D11" 'print Range B11, Dll at B11, C11, D11
    Range("D3, D1, D5").Value = "Range D3, D1, D5" 'print Range D3, D1, D5 at D1, D3, D5
    Range("A7, C5, F10:I10").Value = "acfi"      'print acfi A7, C5 F10, G10, H10, I10
    Range("A3:F3, D2:G5, I5").Select             'activecell is at A3
    Application.CutCopyMode = False
    Range("A3:F3, D2:G5, I5").Activate           'activecell is at A3
    Range("A1").Select
End Sub

Sub cellspropertyquickfontproperty()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("11").Activate
    Dim rownumber As Integer
    For rownumber = 20 To 30:
        Cells(rownumber, 1).Value = rownumber
        With Cells(rownumber, 1).Font
            .Bold = True
            .Italic = True
            .Underline = False
            .Color = RGB(200, 125, 170)
            .Size = 15
            .Name = "Comic Sans MS"
        End With
    Next
End Sub

Sub rangevariable()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("11").Activate
    Dim number20range As Range
    Set number20range = Range("A20:A30")
    number20range.ClearFormats
End Sub

Sub offset()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("11").Activate
    Dim offsetrange As Range
    Set offsetrange = Range("C13")
    offsetrange.offset(1, 0).Value = "I'm down one row from C13 to C14"
    Range("C15").offset().Value = "No Offset, no moving around at C15"
End Sub

Sub deleterangeshiftcellsup()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("11").Activate
    Range("A35").Value = "1"
    Range("A36").Value = "delete me"
    Range("A37").Value = "2 I'm going up"
    Range("A36").Delete shift:=xlShiftUp
    Rows(35).Delete                              'automatically shift up
    Cells(11, 6).Value = "f"
    Cells(11, 7).Value = "g"
    Cells(11, 8).Value = "h"
    Columns(7).Delete                            'automatically shift left
End Sub

Sub insertrangeshiftcellsdown()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("11").Activate
    Range("A34").Value = "A34"
    Range("A35").Value = "A35"
    Range("A36").Value = "A36"
    Range("A35").Insert shift:=xlShiftDown
    Range("A35").Value = "A35 Part II"
    Rows(37).Insert                              'automatically shift down
    Range("A37").Value = "auto shift down row 37"
    Range("A34", Range("A34").End(xlDown)).clearcontents
    Range("J1").Value = "J1"
    Range("K1").Value = "K1"
    Range("L1").Value = "L1"
    Range("K1").Insert shift:=xlShiftToRight
    Range("K1").Value = "K1 Part II"
    Range("J1", Range("J1").End(xlToRight)).clearcontents
End Sub

Sub rowscountcolumnscount()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("11").Activate
    Range("H1").Value = Range("A1:A10").Rows.count 'print 10
    Range("H2").Value = Range("A1:E1").Columns.count 'print 5
    Range("H3").Value = Range("A20", Range("A20").End(xlDown)).Rows.count 'print 11
End Sub

