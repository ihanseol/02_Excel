Attribute VB_Name = "vbatutorial1"
Sub writevalue()
    Sheet1.Range("A1").Value = 5
    Sheet1.Range("B1").Value = "some text"
    Range("C3:E5") = 5.55
    Range("F1") = Now                            'print 6/26/2018 20:19
    Range("C2") = Range("A1")
    Range("A4") = Range("C2") + Range("C3")      'print 10.55
End Sub

Sub copyvalues()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("Sheet1").Activate
    Sheet2.Range("A1") = Sheet1.Range("C2")      'print 5
    Sheet2.Range("A2") = Sheet1.Range("A1") * Sheet1.Range("C3") 'print 27.75
    ThisWorkbook.Worksheets("Sheet2").Range("A5").Value = "Full VBA reference"
    Sheet1.Range("C7", Range("C7").End(xlDown).End(xlToRight)).Select
    'same as
    Range("C7").CurrentRegion.Select
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("Sheet2").Activate
    Sheet2.Range("D2:F4") = Sheet1.Range("C7").CurrentRegion.Value
    'Sheet2.Range("D2") = Sheet1.Range("C7").CurrentRegion.Value 'doesn't paste all values
End Sub

Sub sheetnamecodename()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("Codename").Activate
    Sheetname.Range("A1").Value = "Sheetname (Codename) under Project - VBAProject under " _
                                & "Microsoft Excel Objects"
    Range("A2") = "dddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddddd" _
                & "*space*eeeeeeeeeeeeeeeeeeeeeeeeeeee"
    Range("A5") = 500 _
                + 80 _
                + 90
End Sub

Sub withkeyword()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("Sheet1").Activate
    For i = 10 To 20
        Cells(i, 1).Value = i
    Next i
    With Range("A10:A20").Font
        .Bold = True
        .Underline = True
        .Color = rgbYellow
        .Size = 25
    End With

    Range("A10", Range("A10").End(xlDown)).Select
    Range("A10", Range("A10").End(xlDown)).ClearFormats 'RM: no need to .select. Added for reference
    Range("A10", Range("C20").End(xlDown).End(xlToRight)).Select
    Range("A1").Select
End Sub

Sub transpose()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("Sheet2").Activate
    Range("D7:F9") = WorksheetFunction.transpose(Range("D2:F4").CurrentRegion.Value)
    'RM it appears the starting transpose must match the ending transpose range references
    Range("D11:F16") = WorksheetFunction.transpose(Range("D2:F4").CurrentRegion.Value) 'returns #N/A
End Sub

Sub variables()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("Sheet2").Activate
    'Boolean true or false
    'Currency four decimal places
    'Date
    'Double decimals
    'Long integers
    'String text
    'Variant VBA decide the type at runtime
    'VBA doesn't make you declare variables by default.  Activate default, type Option Explicit at the top
    'dim declarevariablename as type
    
    Dim variablename As String
    Dim integernumber As Long
    'Set integernumber = 100  'Set keyword for objects only?
    integernumber = 100
    Range("A30").Value = integernumber
    
    Dim price As Currency
    price = 29.99
    Range("A31").Value = price
    
    Dim startdate As Date
    startdate = #1/21/2018#
    Range("A32") = startdate
    
    Dim customername As String
    customername = "John Smith"
    Range("A33") = customername
End Sub

Sub immediatewindow()
    'Use Debug.Print to write values, text and results of calculations or to check output.
    'Immediate Window Ctrl+G or View-->Immediate Window
    Debug.Print "This is a test print line"
    Dim integernumber As Long
    integernumber = 5678
    Debug.Print integernumber
End Sub


