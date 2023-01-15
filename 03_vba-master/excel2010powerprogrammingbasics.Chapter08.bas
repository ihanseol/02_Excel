Attribute VB_Name = "Chapter08"
Public examplepublicvariable As String
Public Const appname As String = "Budget Application"

Sub vbademop193()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("8").Activate
    Dim total As Long, i As Long
    'Although VBA can take care of data typing automatically, it does so at a cost: slower execution
    'and less efficient use of memory. An advantage of explicitly declaring your variables is VBA
    'performs additional error checking at the compile stage. e.g. boolean, integer, long, double,
    'currency, decimal, date, string, byte 0 to 255
    total = 0
    For i = 1 To 100 Step 1
        Range("A1").Value = i
        Range("B" & i).Value = i
    Next
    Range("B" & i).Select
    total = 0                                    'rm:  I don't need total rewriting the book example
    For i = 1 To 100
        Range("C" & i).Value = i
    Next i
    Range("C" & i).Select
    ActiveCell.Offset(1, 0).Select
    ActiveCell.Value = TypeName(i)               'print Long. find variable data type
    ActiveCell.Offset(1, 0).Select
    examplepublicvariable = "yes"
    ActiveCell.Value = examplepublicvariable     'print yes
End Sub

Sub publicvariablequickexample()
    examplepublicvariable = "yes2"
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("8").Activate
    Range("E100").Value = examplepublicvariable  'print yes2
End Sub

Sub declareconstantsp206()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("8").Activate
    Dim total As Long, i As Long
    'You declare constants with the Const statement
    Const numquarters As Integer = 4
    Const rate = 0.0725, period = 12             'No data type declaration. VBA determines the data type from the value.
    Const modname As String = "Budget Macros"
    Range("E101").Value = 17 Mod 3               'print 2
    Range("E102").Value = 18 Mod 3               'print 0
End Sub

Sub objectvariablesp215()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("8").Activate
    'An object variable is a variable that represents an entire object, such as a range or a worksheet.
    'Object variables are declared with the Dim or Public statement.
    'Use the Set keyword to assign an object to the variable.
    Dim inputareaobject As Range
    Set inputareaobject = Range("E16:G16")
    inputareaobject.Value = 124
    inputareaobject.Font.Bold = True
    inputareaobject.Font.Italic = True
    inputareaobject.Font.Size = 14
    inputareaobject.Font.Name = Cambria
    'ROMAN function converts a decimal number into a Roman numeral
    Range("E19").Value = Application.WorksheetFunction.Roman(2000)
End Sub

Sub withendwithpagep220()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("8").Activate
    'no with-end with
    Range("E20").Select
    Selection.Font.Name = “Cambria?
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    Selection.Font.Size = 12
    Selection.Font.Underline = xlUnderlineStyleSingle
    Selection.Font.ThemeColor = xlThemeColorAccent1
    Selection.Value = "Okay dokay"
    'yes with-end with
    Range("E21").Select
    With Selection.Font
        .Name = “Cambria?
        .Bold = True
        .Italic = True
        .Size = 12
        .Underline = xlUnderlineStyleSingle
        .ThemeColor = xlThemeColorAccent1
    End With
    Selection.Value = "Okay dokay2"
End Sub

Sub foreachnextp221()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("8").Activate
    'syntax for each-next
    'For Each element In Collection
    '   [instructions]
    '   [Exit For]
    '   [instructions]
    'Next [element]
    For Each Item In Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets
        MsgBox Item.Name
    Next Item
    'same as
    For Each Item In ActiveWorkbook.Worksheets
        MsgBox Item.Name
    Next Item
    'A common use for the For Each-Next construct is to loop through all cells in a range.
    For Each cell In Range("E2:E5")
        cell.Font.Bold = True
        cell.Value = UCase(cell.Value)
    Next cell
End Sub

Sub vbaflowcodep223()
    Application.Workbooks("excel2010powerprogrammingbasics.xlsm").Worksheets("8").Activate
    'GoTo Statements.  the only time sse a GoTo statement is for error handling
    person = "George"
    Range("F1").Value = person
    If person <> "Ron" Then GoTo wrongperson
    Range("F2").Value = "Welcome Ron"
    Exit Sub
wrongperson:
    Range("F2").Value = "Sorry you're not Ron"
    
    'If Then syntax
    'If condition Then
    '    [true_instructions]
    '[ElseIf condition-n Then
    '    [alternate_instructions]]
    '[Else
    '    [default_instructions]]
    'End If
    thetime = 2
    If thetime < 1 Then
        Range("F5").Value = "Good morning"
    ElseIf thetime >= 1 And thetime <= 2 Then
        Range("F5").Value = "Good afternoon"
    Else
        Range("F5").Value = "Good evening"
    End If

    'Select Case syntax
    'Select Case testexpression
    '   [Case expressionlist-n
    '       [instructions-n]]
    '   [Case Else
    '       [default_instructions]]
    'End Select
    'VBA exits a Select Case construct as soon as a True case is found.
    'Select Case can be nested like If Then
    thetime = 2
    Select Case thetime
    Case Is < 1
        Range("F6").Value = "Case good morning"
    Case 1 To 2
        Range("F6").Value = "Case good afternoon"
    Case Else
        Range("F6").Value = "Case good evening"
    End Select
    'WeekDay function determine whether the current day is a weekend returns 1 or 7
    weekdaynumber = Weekday(Now)
    Select Case weekdaynumber
    Case 1, 7
        Range("F7").Value = "This is the weekend"
    Case 4
        Range("F7").Value = "Today is Wednesday"
    Case Else
        Range("F7").Value = "This is the weekday"
    End Select
    quantity = 55
    Select Case quantity
    Case "":
        Exit Sub
    Case 0 To 24:
        discount = 0.1
    Case 25 To 49:
        discount = 0.15
    Case 50 To 74:
        discount = 0.2
    Case Is >= 75:
        discount = 0.25
    End Select
    Range("F8").Value = quantity * discount

    'For-Next Syntax
    'For counter = start To end [Step stepval]
    '   [instructions]
    '   [Exit For]
    '   [instructions]
    'Next [counter]
    numbersum = 0
    For Count = 1 To 100 Step 2
        numbersum = numbersum + Sqr(Count)
        If numbersum > 100 Then
            Exit For
        End If
        Range("F9").Value = numbersum
    Next Count

    'Do While
    'Do [While condition]
    '   [instructions]
    '   [Exit Do]
    '   [instructions]
    'Loop
    'or
    'Do
    '   [instructions]
    '   [Exit Do]
    '   [instructions]
    'Loop [While condition]
    counter = 1
    Do While counter <= 10
        Range("I" & counter).Value = counter
        counter = counter + 1
    Loop
    counter = 1
    Do
        Range("J" & counter).Value = counter
        counter = counter + 1
        If counter = 5 Then
            Range("J" & counter).Value = "We're done"
            Exit Do
        End If
    Loop While counter <= 10

    'Do Until
    'Do [Until condition]
    '   [instructions]
    '   [Exit Do]
    '   [instructions]
    'Loop
    'or
    'Do
    '   [instructions]
    '   [Exit Do]
    '   [instructions]
    'Loop [Until condition]
End Sub

