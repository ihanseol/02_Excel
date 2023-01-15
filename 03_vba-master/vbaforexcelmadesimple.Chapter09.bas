Attribute VB_Name = "Chapter09"
Sub functions()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("9").Activate
    Range("A1").Value = Range("B1:B20").Count    'print 20
    Range("A3").Formula = "=SUM(B1:B20)"         'print 15
    Range("A4").Value = Application.WorksheetFunction.sum(Range("B1:B20")) 'print 15
    Range("A10").Value = LCase("THESE LETTERS ARE LOWERCASE")
    Range("A11").Value = Val("three")            'convert string to numeric print 0
    Range("A12").Value = Str(345)                'convert numeric to string print 345
End Sub

Function doublenumber(x As Integer) As Integer
    doublenumber = x * 2
End Function

Function addtwonumbers(x As Integer, y As Integer) As Integer
    addtwonumbers = x + y
End Function

Sub countoccupiedcells()
    Dim rangecells As Range
    Dim rangecellscount As Integer
    Dim rangecellssum As Integer
    Set rangecells = Range("E1:E5")
    rangecells.Select
    rangecellscount = countvalidcells(rangecells)
    MsgBox rangecellscount
    rangecellssum = addvalidcells(rangecells)
    MsgBox rangecellssum
End Sub

Function countvalidcells(rangecells As Range) As Integer
    Dim mycount As Integer
    mycount = Application.CountA(Selection)      'CountA counts occupied cells
    countvalidcells = mycount
End Function

Function addvalidcells(rangecells As Range) As Integer
    Dim mysum As Integer
    mysum = Application.sum(Selection)
    addvalidcells = mysum
End Function

Sub multiplefunction()
    Dim mathoperation As String
    Dim number1 As Integer
    Dim number2 As Integer
    Dim solution As Integer
    mathoperation = Range("G1").Value
    number1 = Range("G2").Value
    number2 = Range("G3").Value
    solution = mathcalculation(mathoperation, number1, number2)
    Range("G4").Value = solution
End Sub

Function mathcalculation(operation As String, n1 As Integer, n2 As Integer) As Integer
    Dim answer As Integer
    If operation = "add" Then
        answer = n1 + n2
    ElseIf operation = "subtract" Then
        answer = n1 - n2
    ElseIf operation = "multiply" Then
        answer = n1 * n2
    ElseIf operation = "divide" Then
        answer = n1 / n2
    End If
    mathcalculation = answer
End Function

Sub byrefsquarenumberbefore()
    'passing a parameter by reference
    Dim number As Integer
    number = 10
    Range("K1").Value = "Number variable passed to square function is " & number 'print 10
    Range("K2").Value = "Function value number squared is " & byrefsquarenumberafter(number) 'print 100
    Range("K3").Value = "Number variable received from square function is " & number 'print 222
End Sub

Function byrefsquarenumberafter(ByRef number As Integer)
    byrefsquarenumberafter = number * number
    number = 222
End Function

Sub byvalsquarenumberbefore()
    'passing a parameter by value
    Dim number As Integer
    number = 10
    Range("K5").Value = "Number variable passed to square function is " & number 'print 10
    Range("K6").Value = "Function value number squared is " & byvalsquarenumberafter(number) 'print 100
    Range("K7").Value = "Number variable received from square function is " & number 'print 10
End Sub

Function byvalsquarenumberafter(ByVal number As Integer)
    byvalsquarenumberafter = number * number
    number = 222
End Function

