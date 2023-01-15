Attribute VB_Name = "Introduction"
Public Const filename As String = "MicrosoftExcelProgrammingChapter03Introduction.xlsm"

Sub AssignValueToCell()
Workbooks(filename).Worksheets("Sheet1").Activate
Range("A1").Value = 100
[A2].Value = 200
Range("A3") = 0.2
Cells(4, 1).Value = Bonus(Range("A1"), Range("A3"))
Cells(10, "A").Value = "A10"
End Sub

Function Bonus(Salary, Percent)
Bonus = Salary * Percent
End Function

Sub CellReference()
    Workbooks(filename).Worksheets("Sheet1").Activate
    Range("C4").Interior.Color = rgbLightBlue
    Range("B1:B7").Interior.Color = rgbLightGreen
    Range("D1:D8, F1:H2, F7:H8, G2:G6").Interior.Color = rgbLightGrey
    Range("J:J").Interior.Color = rgbLightSalmon
    Range("11:11").Interior.Color = rgbLightYellow
    Range("L:M").Interior.Color = rgbLightSeaGreen
    Range("14:16").Interior.Color = rgbLightCyan
End Sub
