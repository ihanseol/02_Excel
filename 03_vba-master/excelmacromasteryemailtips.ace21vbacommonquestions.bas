Attribute VB_Name = "ace21vbacommonquestions"
Sub immediatewindow()
    'list output in Immediate Windows View-->Immediate Windows or Ctrl+G
    Debug.Print "this is a test"
End Sub

Sub openworkbook()
    'formal way declare a variable
    Dim wk As Workbook
    Set wk = Workbooks.Open("G:\Raymond\Excel Files 2GB Backup 072218\YouTube Self Job Review And Training.xlsx")
    Workbooks.Open ("G:\Raymond\Excel Files 2GB Backup 072218\YouTube Self Job Review And Training.xlsx")
End Sub

Sub getlastrow()
    'only works when data starts at Range("A1")
    ThisWorkbook.Worksheets("ace21commonVBA").Activate
    Dim lastrow As Long, lastcolumn As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).row
    MsgBox lastrow
    lastcolumn = Cells(1, Columns.Count).End(xlToLeft).Column
    MsgBox lastcolumn

    Range("E20").Select
    lastrow = Cells(Rows.Count, 5).End(xlUp).row
    MsgBox lastrow                               'return 28 which is row 28
    lastcolumn = Cells(20, Columns.Count).End(xlToLeft).Column
    MsgBox lastcolumn                            'return 8 which is column H
    Range("E20").currentregion.Select
End Sub

Function addition(a As Long, b As Long)
    addition = (a + b) + 10
End Function

Sub callfunction()
    Dim result As Long
    result = addition(5, 6)
    MsgBox result                                'return 21
End Sub

Sub addformulatocell()
    Range("A20").Formula = "=sum(A15:a19)"
End Sub

Sub printworksheetnames()
    ThisWorkbook.Worksheets("ace21commonVBA").Activate
    Dim sh As Worksheet
    Range("A22").Select
    For Each sh In ThisWorkbook.Worksheets()
        ActiveCell.Value = sh.Name
        ActiveCell.Offset(1, 0).Select
    Next
    'worksheets from right to left
    Dim i As Long
    For i = ThisWorkbook.Worksheets.Count To 1
        Debug.Print ThisWorkbook.Worksheets(i).Name
    Next
    'worksheets at a specific starting point
    Dim i As Long
    For i = 1 To 5
        Debug.Print ThisWorkbook.Worksheets(i).Name
    Next
End Sub

Sub copycells()
    Range("A1:A4").Copy Destination:=Range("I1")
    'also
    Range("A1:A4").Copy
    Range("J1").PasteSpecial xlPasteValues
    Range("J1").PasteSpecial xlPasteFormats
    Range("J1").PasteSpecial xlPasteFormulas
End Sub

'start tip #11
Sub findtextincells()
    Dim rg As Range
    Set rg = Range("A22:A26").Find("co")
    MsgBox rg.Address, rg.row, rg.Column
End Sub

Sub sumrange()
    Dim rg As Range
    Set rg = Range("A15:A20")
    MsgBox WorksheetFunction.Sum(rg)
    Range("C15").Value = WorksheetFunction.Sum(rg)
End Sub

Sub formatcells()
    With Range("A1:A10")
        .Font.Bold = True
        .Font.Size = 10
        .Font.Color = rgbRed
        .Interior.Color = rgbLightBlue
        .Borders.LineStyle = xlDouble
        .Borders.Color = rgbGreen
    End With
    Range("A1:A10").ClearFormats
End Sub

Sub hiderows()
    Rows(1).Hidden = True
    Rows(1).Hidden = False
    Rows("1:3").Hidden = True
    Rows("1:3").Hidden = False
    'hide more than one row
    For i = 1 To 20 Step 2
        Rows(i).Hidden = True
    Next
    For i = 1 To 20 Step 2
        Rows(i).Hidden = False
    Next
End Sub

Sub hidecolumns()
    Columns(4).Hidden = True
    Columns(4).Hidden = False
    Columns("A:C").Hidden = True
    Columns("A:C").Hidden = False
End Sub

Sub copyworksheet()
    'copy worksheet to last position
    Worksheets("ace21commonVBA").Copy After:=Worksheets(Worksheets.Count)
    Worksheets(Worksheets.Count).Name = "Copy Temp Delete"
    Worksheets(Worksheets.Count).Delete
    'copy worksheet to first position
    Worksheets("ace21commonVBA").Copy Before:=Worksheets(1)
    Worksheets(1).Name = "Copy Temp Delete"
    Worksheets(1).Delete
End Sub

Sub addworksheet()
    'add worksheet to first position
    Dim sh As Worksheet
    Set sh = Worksheets.Add(Before:=Worksheets(1))
    sh.Name = "Accounts"
    Worksheets("Accounts").Delete
    'add two worksheets to first positions
    Worksheets.Add Before:=Worksheets(1), Count:=2
    Worksheets(1).Name = "First New Worksheet"
    Worksheets(2).Name = "Second New Worksheet"
    Worksheets(1).Delete
    Worksheets(1).Delete                         'The Second New Worksheet worksheet is the first worksheet after First _
                                                 New Worksheet was deleted
End Sub

Sub addworkbookornewexcelfile()
    Dim bk As Workbook
    Set bk = Workbooks.Add
    bk.SaveAs "C:\Users\Mar\Desktop\temp.xlsx"
End Sub

Sub insertrowinsertcolumn()
    Worksheets("ace21commonVBA").Activate
    Rows(1).Insert
    Rows(1).Delete
    Columns("D").Insert
    Columns("D").Delete
    Rows("10:12").Insert
    Rows("10:12").Delete
    Columns("I:L").Insert
    Columns("I:L").Delete
    'insert more than one row not sequential.  Notice deleting requires lower count.
    For i = 1 To 21 Step 3
        Rows(i).Insert
    Next i
    For i = 1 To 14 Step 2
        Rows(i).Delete
    Next i
End Sub

Sub rangecellsoffset()
    Worksheets("ace21commonVBA").Activate
    Range("A30").Value = 6
    Range("C30:C37, E35").Value = 99.99
    Cells(30, 1).Value = 6
    'same as
    Cells("30,1").Value = 6
    'RM it appears cells can't assign multiple cell ranges
    Range("E35").Select
    Range("C37").Offset(1, 0).Value = "Down one cell"
    Range("C37").Offset(0, 2).Value = "Right two cells"
    Range("C30:C32").Offset(-1, 4).Interior.Color = rgbRed

End Sub


