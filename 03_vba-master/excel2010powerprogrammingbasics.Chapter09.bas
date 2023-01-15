Attribute VB_Name = "Chapter09"
Sub subprocedures()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("9").Activate
End Sub

Sub passargumentstoprocedures255()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("9").Activate
    x = Range("A1").Value
    y = Range("A2").Value
    operation = Range("A3").Value
    Call arithemeticprocedure(x, y, operation)
    myvalue = 10
    Call process(myvalue)
    Range("A6").Value = myvalue                  'print 100
End Sub

Sub arithemeticprocedure(x, y, operation)
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("9").Activate
    If operation = "add" Then
        Range("A4").Value = x + y
    ElseIf operation = "subtract" Then
        Range("A4").Value = x - y
    ElseIf operation = "multiply" Then
        Range("A4").Value = x * y
    ElseIf operation = "divide" Then
        Range("A4").Value = x / y
    End If
End Sub

Sub process(yourvalue As Integer)
    'If you don’t want the called procedure to modify any variables passed as arguments, you can modify
    'the called procedure’s argument list so that arguments are passed to it by value rather than by
    'reference. To do so, precede the argument with the ByVal keyword Sub process(ByVal yourvalue)
    yourvalue = yourvalue * 10
End Sub

Sub sortsheets269()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("9").Activate
    'loop all worksheets in open workbook print to a worksheet
    sheetcount = ActiveWorkbook.Sheets.Count
    Range("C1").Value = sheetcount
    ReDim sheetnames(1 To sheetcount)            'created an array of sheetnames(1) to sheetnames(n)
    For i = 1 To sheetcount
        sheetnames(i) = ActiveWorkbook.Sheets(i).Name
        Range("D" & i).Value = sheetnames(i)
    Next i
End Sub

Sub screenupdatingpractice275()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("9").Activate
    Application.ScreenUpdating = False
    Range("B1").Value = 1
    For i = 2 To 101 Step 1
        Range("B" & i).Value = i + Range("B" & i - 1).Value
    Next i
    Application.ScreenUpdating = True
End Sub

