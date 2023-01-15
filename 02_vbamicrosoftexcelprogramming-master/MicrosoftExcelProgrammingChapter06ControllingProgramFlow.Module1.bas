Attribute VB_Name = "Module1"
Sub DoWhileLoop()
    Workbooks("MicrosoftExcelProgrammingChapter06ControllingProgramFlow.xlsm").Worksheets("Sheet1").Activate
    Dim Counter As Integer
    Dim RowNum As Integer
    Counter = 1
    RowNum = 1
    Do While Counter < 11
        If Counter < 0 Then
            Exit Do
        End If
        Cells(RowNum + 9, 1).Value = Counter
        Counter = Counter + 1
        RowNum = RowNum + 1
    Loop
End Sub

Sub DoUntilLoop()
    Dim RowNum As Integer
    RowNum = 10
    Do Until IsEmpty(Cells(RowNum, 8))
        Cells(RowNum, 9).Value = Cells(RowNum, 8) * 0.1
        RowNum = RowNum + 1
    Loop
End Sub

Sub ForNextLoop()
    Dim Counter As Integer
    For Counter = 29 To 32 Step 1
        Range("A" & Counter).Value = Counter
    Next Counter
End Sub

Sub ForEachLoop()
    Dim NewArray(29 To 32) As Integer
    For Each eachNewArray In NewArray
        'Range("H" & eachNewArray).Value = eachNewArray
        'Raymond Mar:  skipped For Each Loop
        ActiveCell.Value = eachNewArray
    Next eachNewArray
End Sub

Sub IfThenElseStatement()
    Dim rownumber As Integer
    rownumber = 34
    Do Until IsEmpty(Cells(rownumber, 1))
        If Cells(rownumber, 1).Value Mod 2 = 0 Then
            Cells(rownumber, 2).Value = "Even Number"
        Else
            Cells(rownumber, 2).Value = "Odd Number"
        End If
        rownumber = rownumber + 1
    Loop
End Sub

Sub IfThenElseIfStatement()
    Dim rownumber As Integer
    rownumber = 34
    Do While Not IsEmpty(Range("H" & rownumber))
        If Range("I" & rownumber).Value = "TX" Then
            Range("J" & rownumber).Value = Range("H" & rownumber).Value + 10
        ElseIf Range("I" & rownumber).Value = "CA" Then
            Range("J" & rownumber).Value = Range("H" & rownumber).Value + 20
        ElseIf Range("I" & rownumber).Value = "FL" Then
            Range("J" & rownumber).Value = Range("H" & rownumber).Value + 30
        Else
            Range("J" & rownumber).Value = Range("H" & rownumber).Value + 40
        End If
        rownumber = rownumber + 1
    Loop
End Sub

Sub SelectCase()
    Dim rownumber As Integer
    rownumber = 47
    Do While Not IsEmpty(Range("A" & rownumber))
        Select Case Range("B" & rownumber)
        Case "TX"
            Range("C" & rownumber).Value = Range("A" & rownumber).Value + 10
        Case "CA"
            Range("C" & rownumber).Value = Range("A" & rownumber).Value + 20
        Case "FL"
            Range("C" & rownumber).Value = Range("A" & rownumber).Value + 30
        Case Else
            Range("C" & rownumber).Value = Range("A" & rownumber).Value + 40
        End Select
        rownumber = rownumber + 1
    Loop
End Sub

