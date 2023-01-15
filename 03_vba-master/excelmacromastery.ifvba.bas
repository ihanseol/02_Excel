Attribute VB_Name = "ifvba"
Sub ifstatement()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("if").Activate
    'if [condition] then [do something] end if
    'if [condition] then [do something] else [condition2] end if
    'if [condition] then [do something] elseif [condition2] then [do something2] else [dosomething3] _
    end if
    Dim score As Long
    score = 45
    If score = 100 Then
        Range("A1").Value = "Perfect"
    ElseIf score > 50 Then
        Range("A1").Value = "Passed"
    ElseIf score > 30 Then
        Range("A1").Value = "Try again"
    Else
        Range("A1").Value = "Yikes"
    End If

    Dim i As Long
    For i = 4 To 13
        If Range("C" & i) > 50 Then
            Range("E" & i).Value = "Over 50"
        End If
    Next
    Columns(5).Delete                            'Delete column

    For i = 4 To 13
        If Range("C" & i) >= 85 Then
            Range("E" & i).Value = "High Destinction"
        ElseIf Range("C" & i) >= 75 Then
            Range("E" & i).Value = "Destinction"
        ElseIf Range("C" & i) >= 55 Then
            Range("E" & i).Value = "Credit"
        ElseIf Range("C" & i) >= 50 Then
            Range("E" & i).Value = "Pass"
        Else
            Range("E" & i).Value = "Stay Alive"
        End If
    Next
    Columns(5).Delete                            'Delete column
End Sub

Sub ifstatementwithvariables()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("if").Activate
    Dim firstrow, lastrow, i, marks As Long
    Dim rank As String
    firstrow = 4
    lastrow = Cells(Rows.count, 1).End(xlUp).Row 'get last row number from entire worksheet

    For i = firstrow To lastrow
        marks = Range("C" & i).Value
        If marks >= 85 Then
            rank = "High Destinction"
        ElseIf marks >= 75 Then
            rank = "Destinction"
        ElseIf marks >= 55 Then
            rank = "Credit"
        ElseIf marks >= 50 Then
            rank = "Pass"
        Else
            rank = "Stay Alive"
        End If
        Range("E" & i).Value = rank
    Next
    Columns(5).Delete                            'Delete column

    Dim subject As String
    firstrow = 4
    lastrow = Cells(Rows.count, 1).End(xlUp).Row 'get last row number from entire worksheet
    For i = firstrow To lastrow
        marks = Range("C" & i).Value
        subject = Range("D" & i).Value
        If marks >= 85 And (subject = "French" Or subject = "History") Then
            rank = "High Destinction"
        ElseIf marks >= 75 And (subject = "French" Or subject = "History") Then
            rank = "Destinction"
        ElseIf marks >= 55 And (subject = "French" Or subject = "History") Then
            rank = "Credit"
        ElseIf marks >= 50 And (subject = "French" Or subject = "History") Then
            rank = "Pass"
        Else
            rank = "Excluded"
        End If
        Range("E" & i).Value = rank
    Next
    Columns(5).Delete                            'Delete column
End Sub

Sub ifstatementiff()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("if").Activate
    'VBA has a function similar to Excel =IF() or =IF(condition, action if true, action if false)
    'It's IIf(condition, action if true, action if false)
    Dim i As Long
    Dim result As String
    For i = 4 To 13
        result = IIf(Range("C" & i) >= 50, "Equal or Over 50", "Under 50")
        Range("E" & i).Value = result
    Next
    Columns(5).Delete                            'Delete column
    'Nested IIF
    For i = 4 To 13
        result = IIf(Range("C" & i) >= 50, "Equal or Over 50", IIf(Range("C" & i) >= 30, _
                                                                   "Equal or Over 30", "Under 30"))
        Range("E" & i).Value = result
    Next
    Columns(5).Delete                            'Delete column
End Sub

Sub ifstatementnestedif()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("if").Activate
    Dim i, markscell As Long
    Dim subjectcell As String
    For i = 4 To 13
        markscell = Range("C" & i).Value
        subjectcell = Range("D" & i).Value
        If markscell >= 85 Then
            If subjectcell = "History" Then
                Range("E" & i).Value = subjectcell & " High Distinction"
            ElseIf subjectcell = "French" Then
                Range("E" & i).Value = subjectcell & " High Distinction"
            Else
                Range("E" & i).Value = "High Distinction"
            End If
        ElseIf markscell >= 75 Then
            If subjectcell = "History" Then
                Range("E" & i).Value = subjectcell & " Distinction"
            ElseIf subjectcell = "French" Then
                Range("E" & i).Value = subjectcell & " Distinction"
            Else
                Range("E" & i).Value = "Distinction"
            End If
        ElseIf markscell >= 40 Then
            If subjectcell = "History" Then
                Range("E" & i).Value = subjectcell & " Pass"
            ElseIf subjectcell = "French" Then
                Range("E" & i).Value = subjectcell & " Pass"
            ElseIf subjectcell = "History" Then
                Range("E" & i).Value = subjectcell & " Pass"
            Else
                Range("E" & i).Value = "Pass"
            End If
        Else
            Range("E" & i).Value = "The rest"
        End If
    Next
    Columns(5).Delete                            'Delete column
End Sub

Sub caseselectstatement()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("if").Activate
    'Select Case [variable]
    'Case [condition1]
    'Case [condition2]
    'Case [condition3]
    'Case [conditionn]
    'Case Else
    'End Select
    Dim i, marks As Long
    Dim rating As String

    For i = 4 To 13
        marks = Range("C" & i).Value
        Select Case marks
        Case 85 To 100
            rating = "High Destinction"
        Case 75 To 84
            rating = "Destinction"
        Case 55 To 74
            rating = "Credit"
        Case 40 To 54
            rating = "Pass"
        Case Else
            rating = "Fail"
        End Select
        Range("E" & i).Value = rating
    Next i
    Columns(5).Delete
End Sub

Sub caseisstatement()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("if").Activate
    'Select Case [variable]
    'Case Is [condition1]
    'Case Is [condition2]
    'Case Is [condition3]
    'Case Is [conditionn]
    'Case Else
    'End Select
    Dim i, marks As Long
    Dim rating As String

    For i = 4 To 13
        marks = Range("C" & i).Value
        Select Case marks
        Case Is >= 85
            rating = "High Destinction"
        Case Is >= 75
            rating = "Destinction"
        Case Is >= 55
            rating = "Credit"
        Case Is >= 40
            rating = "Pass"
        Case Else
            rating = "Fail"
        End Select
        Range("E" & i).Value = rating
    Next i
    Columns(5).Delete

    Dim marksquick As Long
    marksquick = 7
    Select Case marksquick
    Case Is = 5, 7, 9
        MsgBox "Yes there is a " & marksquick
    Case Else
        MsgBox "Nothing"
    End Select
End Sub


