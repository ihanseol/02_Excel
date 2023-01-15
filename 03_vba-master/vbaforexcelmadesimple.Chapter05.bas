Attribute VB_Name = "Chapter05"
Sub variables()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("5").Activate
    'most common data types in my experience:  Boolean, Integer, Long, Single for decimals, _
    Currency, Date, String, Array, Variant store any data type String or Double
    Dim decimalnumber As Single, nectarpoints As Integer, daughtersname As String
    Dim lastdayofterm As Date
    Dim topcell As Range                         'Range object declaration
    Dim thecell As Object                        'genric Object declaration
    decimalnumber = 3.55                         'set keyword not required
    'assigning a value to an object variable and a non-object variable is that the object _
    assignment must begin with the keyword Set.
    nectarpoints = 5000
    daughtersname = "Rhiannon"
    lastdayofterm = #4/3/2003#
    Set topcell = Range("A1")                    'set keyword required
    Range("A1").Value = decimalnumber            'print 3.55
    Range("A2").Value = lastdayofterm            'print 4/3/2003
End Sub

Sub variableexample()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("5").Activate
    Dim firstnumber, secondnumber, sum As Integer
    firstnumber = Range("C1").Value
    secondnumber = Range("C2").Value
    sum = firstnumber + secondnumber
    Range("C3").Value = sum
End Sub

Sub constants()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("5").Activate
    Const speedoflight As Double = 186000
    Const pi As Double = 3.142
End Sub

Sub userdefinedtypes()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("5").Activate
    'Type declaration is above Sub userdefinedtypes()
    'Type employeerecord
    '    firstname As String
    '    lastname As String
    '    telephonenumber As String
    '    salary As Currency
    '    startdate As Date
    'End Type
    Dim employee As employeerecord
    employee.firstname = "John"
    'RM:  is user define types like classes in Python?
End Sub

Sub userdefinedtypesexample()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("5").Activate
    'Type declaration is above Sub userdefinedtypesexample()
    'Type Course
    '    CourseName As String
    '    Unit As String
    '    numberofstudents As Integer
    'End Type
    Dim course1 As Course
    Dim course2 As Course
    course1.CourseName = "math"
    course1.Unit = "calculus"
    course1.numberofstudents = 40
    Range("A10").Value = course1.CourseName      'print math
    Range("A11").Value = course1.Unit            'print calculus
    Range("A12").Value = course1.numberofstudents 'print 40
    course2.CourseName = "business"
    course2.Unit = "accounting"
    course2.numberofstudents = 50
    Range("A13").Value = course2.CourseName      'print business
    Range("A14").Value = course2.Unit            'print accounting
    Range("A15").Value = course2.numberofstudents 'print 50
End Sub

Sub arrays()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("5").Activate
    'The index numbering starts at 0, though we tend to think ofpositions as starting at 1.
    'Thus, the element at position 1 is written as expenses(0).
    Const n As Integer = 5
    Dim d(n) As Integer
    Dim i As Integer
    For i = 0 To n Step 1
        d(i) = i ^ 2
        Range("E" & i + 1).Value = d(i)          'print 0
        Range("F" & i + 1).Value = "d(" & i & ") for which d is the array and i = " & i
        'print d(0) for which d is the array and i = 0
    Next i
End Sub

Sub dynamicarrays()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("5").Activate
    Dim daysales(5) As Integer
    'daysales(0) to daysales(4) random number between 10 and 100
    daysales(0) = Int((100 - 10 + 1) * Rnd + 10)
    daysales(1) = Int((100 - 10 + 1) * Rnd + 10)
    daysales(2) = Int((100 - 10 + 1) * Rnd + 10)
    daysales(3) = Int((100 - 10 + 1) * Rnd + 10)
    daysales(4) = Int((100 - 10 + 1) * Rnd + 10)
    'daysales(0) = Range("E10").Value
    'daysales(1) = Range("E11").Value
    'daysales(2) = Range("E12").Value
    'daysales(3) = Range("E13").Value
    'daysales(4) = Range("E14").Value
    'print each daysales() array random number
    Range("E10").Value = daysales(0)
    Range("E11").Value = daysales(1)
    Range("E12").Value = daysales(2)
    Range("E13").Value = daysales(3)
    Range("E14").Value = daysales(4)
    totals = 0
    For i = 0 To 4 Step 1
        totals = totals + daysales(i)
    Next i
    'print the totals from for loop summing the daysales() array
    Range("E16").Value = totals
End Sub


