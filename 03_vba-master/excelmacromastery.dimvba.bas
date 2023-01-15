Attribute VB_Name = "dimvba"
Sub dimvba()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("dim").Activate
    'The dim keyword is Dimension.  dim declares variables.
    'If the type is not declared, VBA assigns Variant as the type.
    'You can use a variable without using the Dim Statement.  VBA assigns the variable Variant.
    'Intellisense doesn't work with variables declared Variant.

    'dim variablename as Long <--use for integers
    'dim variablename as Double <--use for decimals
    'dim variablename as Currency
    'dim variablename as String
    'dim variablename as String * 10 <--size 10.  Size is always the same
    'dim variablename as Date
    'dim variablename as Boolean
    'dim variablename as Variant
    'dim variablename as New objecttype
    'dim variablename as objecttype2
    'set variablename = objecttype2
    'dim variablename(1 to 6) as Long <--static array
    'dim variablename() as Long  <--dynamic array 1 of 2
    'ReDim variablename(1 to 6)  <--dynamic array 2 of 2
    'dim countnumber6 as long: count = 6

    Dim namestring As String
    Dim count As Long
    Dim decimalnumber As Double
    Dim amount As Currency
    Dim eventdate As Date
    Dim userid As String * 8
    Dim sh As Worksheet
    Dim wk As Workbook
    Dim rg As Range
    Dim collection1 As Collection
    Dim object1 As New Class1
    Dim collection2 As Collection
    Set collection2 = New Collection
    Dim object2 As Class1
    Set object2 = New Class1
    Dim arrayScores(1 To 5) As Long
    Dim arrayCountries(0 To 9) As String
    Dim dynamicarrayMarks As Long
    ReDim dynamicarrayMarks(1 To 10) As Long
    Dim dynamicarrayNames As String
    ReDim dynamicarrayNames(5 To 15) As String
    Dim nameplease As String, ageplease As Long, countplease As Long
End Sub

Sub dimobjects()
    Dim i As Long, count As Long
    Dim wk2 As Workbook, sh2 As Worksheet, rg2 As Range
    Set wk2 = Workbooks.Open("G:\Raymond\Excel Files 2GB Backup 072218\VBA Macros Round Two\" _
                           & "saveasfilename.xlsx")
    Set sh2 = wk2.Worksheets(1)
    Set rg2 = sh2.Range("A1:A10")
    Range("A1").Select
    For i = 1 To rg2.Rows.count
        MsgBox ActiveCell.Value
        ActiveCell.offset(1, 0).Select
    Next i
End Sub

Sub dimobjectsbetterdeclaration()
    Dim wk3 As Workbook
    Set wk3 = Workbooks.Open("G:\Raymond\Excel Files 2GB Backup 072218\VBA Macros Round Two\" _
                           & "saveasfilename.xlsx")
    Dim sh3 As Worksheet
    Set sh3 = wk3.Worksheets(1)
    Dim rg3 As Range
    Set rg3 = sh3.Range("A1:A10")

    Dim i As Long, count As Long
    sh3.Cells(1, 1).Select
    For i = 1 To rg3.Rows.count
        MsgBox ActiveCell.Value
        ActiveCell.offset(1, 0).Select
    Next i
End Sub

Sub fixedstring()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("dim").Activate
    Dim string1 As String * 4, string2 As String * 4
    string1 = "John Smith"
    string2 = "Tomm"
    Range("A1").Value = string1                  'print John
    Range("A2").Value = string2                  'print Tomm
End Sub

Sub dimobjectsbasics()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("dim").Activate
    Dim wk As Workbook
    Set wk = Workbooks.Open("G:\Raymond\Excel Files 2GB Backup 072218\VBA Macros Round Two\" _
                          & "saveasfilename.xlsx") 'assign variable wk and open workbook saveasfilename.xlsx
    wk.Close                                     'Close workbook
    'Set wk = Workbooks.Add 'assign variable wk and add new workbook
    'Set wk = Workbooks(1) 'assign variable wk to first workbook opened
    'Set wk = Workbooks("filename.xlsx") 'assign variable wk to workbook filename.xlsx
    'Set wk = ActiveWorkbook 'assign variable wk to active workbook
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets.Add         'assign variable sh and add new worksheet
    sh.Delete                                    'Delete worksheet assigned to variable sh
    'Set sh = ThisWorkbook.Worksheets(1) 'assign variable sh to leftmost worksheet
    'Set sh = ThisWokrbook.Worksheet("worksheetname") 'assign variable sh to worksheet worksheetname
    'Set sh = ActiveSheet 'assign variable sh to active worksheet
    Set sh = ThisWorkbook.Worksheets("Sheet1")   'assign variable sh to worksheet Sheet1
    sh.Activate                                  'make worksheet Sheet1 active

    'Go to dim worksheet
    Dim dimworksheet As Worksheet
    Set dimworksheet = Workbooks("excelmacromastery.xlsm").Worksheets("dim")
    dimworksheet.Activate
    Dim rg As Range
    Set rg = dimworksheet.Range("A1")
    rg.Activate                                  'go to cell A1
    Set rg = dimworksheet.Range("B4:F7")
    'rg.Activate 'selected cells B4:F7
    'same as
    rg.Select                                    'selected cells B4:F7
End Sub

Sub dimarrays()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("dim").Activate
    'Both dim staticarraylong stores 7 longs
    Dim staticarraylong(0 To 6) As Long
    Dim staticarraylong(6) As Long
    'Both dim dynamicarraylong are dynamic
    Dim dynamicarraylong() As Long
    ReDim dynamicarraylong(0 To 6) As Long
End Sub


