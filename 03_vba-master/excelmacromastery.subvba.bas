Attribute VB_Name = "subvba"
Sub callafunction()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("sub").Activate
    Range("A1") = GetAmount                      'print 55
End Sub

Function GetAmount() As Long
    GetAmount = 55
End Function

Sub useafunction()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("sub").Activate
    Range("A2") = GetValue(24.99)                'print 2499
End Sub

Function GetValue(amount As Currency) As Long
    GetValue = amount * 100
End Function

Sub ByRefByVal()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("sub").Activate
    'ByRef is pass by reference.  Creates a reference. _
    Parameter value changes return to the calling sub or function.  It's default.
    'ByVal is pass by value.  Creates a copy. _
    Parameter value doesn't change return to the calling sub or function.
    Dim x As Long
    x = 1
    Range("A4") = "x before ByRef is " & x       'print x before ByRef is 1
    SubByRef x                                   'don't use parenthesis
    Range("A5") = "x after ByRef is " & x        'print x after ByRef is 99

    x = 1
    Range("A6") = "x before ByVal is " & x       'print x before ByVal is 1
    SubByVal x                                   'don't use parenthesis
    Range("A7") = "x after ByVal is " & x        'print x after ByVal is 1
End Sub

Sub SubByRef(ByRef x As Long)
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("sub").Activate
    x = 99
End Sub

Sub SubByVal(ByVal x As Long)
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("sub").Activate
    x = 99
End Sub

Sub mainparameters()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("sub").Activate
    'RM:  Can't get Daily Report to print on Range("A9")
    Dim x As String
    optionalparameters
    Range("A9") = x                              'print null
    x = "Weekly Report"
    optionalparameters x
    Range("A10") = x                             'print Weekly Report
End Sub

Sub optionalparameters(Optional reportname As String = "Daily Report")
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("sub").Activate
    reportname
End Sub


