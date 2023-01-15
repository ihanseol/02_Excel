Attribute VB_Name = "Chapter06"
Sub dowhile()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("6").Activate
    Dim counter As Integer
    counter = 1
    Do While counter < 11
        Cells(counter, 1).Value = counter
        counter = counter + 1
    Loop
    counter = 1
    Do While Not (IsEmpty(Cells(counter, 1)))
        If Cells(counter, 1).Value Mod 2 = 0 Then
            Cells(counter, 2).Value = "even number"
        End If
        counter = counter + 1
    Loop
End Sub

Sub dountil()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("6").Activate
    Dim counter As Integer
    counter = 1
    Do Until counter = 11
        Cells(counter, 1).Value = counter
        counter = counter + 1
    Loop
    counter = 1
    Do Until IsEmpty(Cells(counter, 1))
        If Cells(counter, 1).Value Mod 2 = 0 Then
            Cells(counter, 3).Value = "even with do until loop"
        End If
        counter = counter + 1
    Loop
End Sub

Sub exitdo()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("6").Activate
    Dim counter As Integer
    counter = 13
    Do While counter < 23
        Cells(counter, 1).Value = counter
        If counter = 17 Then
            Cells(counter, 2).Value = "Let's end the loop"
            Exit Do
        End If
        counter = counter + 1
    Loop
End Sub

Sub fornextloop()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("6").Activate
    Dim month(1 To 3) As String
    month(1) = "Jan"
    month(2) = "Feb"
    month(3) = "Mar"
    For i = 1 To 10 Step 1
        Range("F" & i).Value = "Simple for loop " & i
    Next i
    For i = 2 To 20 Step 2
        Range("G" & i).Value = "Another simple for loop " & i
        Range("G" & i).Select
        ActiveCell.offset(0, -1).Font.Color = RGB(255, 0, 0) 'one cell to left font color red
    Next i
    For i = 1 To 3 Step 1
        Range("J" & i).Value = month(i)          'print Jan, Feb, Mar Cells J1:J3
    Next i
End Sub

Sub foreachinloop()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("6").Activate
    Dim loverange As Range
    Dim count As Integer
    Set loverange = Range("L10:L20")
    For Each cell In Range("K1:K10")
        cell.Value = "for each test"
    Next
    count = 1
    For Each ube In loverange:
        ube.Value = "ube " & count
        count = count + 5
    Next
End Sub

Sub ifthenelse()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("6").Activate
    Dim randomrange As Range
    Dim rownumber As Integer
    Set randomrange = Range("A20:A40")
    rownumber = 20
    For Each rr In randomrange:
        rr.Value = Int(Rnd() * 100 + 1)
    Next
    Do Until IsEmpty(Cells(rownumber, 1))
        If Cells(rownumber, 1).Value Mod 5 = 0 Then
            Cells(rownumber, 2).Value = "Number is divisible by 5"
        ElseIf Cells(rownumber, 1).Value Mod 3 = 0 Then
            Cells(rownumber, 2).Value = "Number is divisible by 3"
        Else
            Cells(rownumber, 2).Value = "A number"
        End If
        rownumber = rownumber + 1
    Loop
    'or
    rownumber = 20
    For Each rrr In randomrange:
        If Range("A" & rownumber).Value Mod 5 = 0 Then
            rrr.offset(0, 3).Value = "Number is divisible by 5"
        ElseIf Range("A" & rownumber).Value Mod 3 = 0 Then
            rrr.offset(0, 3).Value = "Number is divisible by 3"
        Else
            rrr.offset(0, 3).Value = "A number"
        End If
        rownumber = rownumber + 1
    Next
End Sub

Sub selectcase()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("6").Activate
    Dim rownumber As Integer
    rownumber = 2
    Do While Not IsEmpty(Cells(rownumber, 15))
        Select Case Cells(rownumber, 15)
        Case "TX"
            Cells(rownumber, 16).Value = Cells(rownumber, 14).Value * 1.05
        Case "FL"
            Cells(rownumber, 16) = Cells(rownumber, 14) * 1.08
        Case "CA"
            Cells(rownumber, 16) = Cells(rownumber, 14) * 1.1
        Case "UT"
            Cells(rownumber, 16) = Cells(rownumber, 14) * 1.04
        Case Else
            Cells(rownumber, 16) = Cells(rownumber, 14)
        End Select
        rownumber = rownumber + 1
    Loop
    For Each cell In Range("N2:N11")
        Select Case cell
        Case 5 To 10
            cell.offset(0, 3).Value = "between 5 to 10"
        Case 11 To 20
            cell.offset(0, 3).Value = "between 11 to 20"
        End Select
    Next
    For Each cell In Range("N13:N22")
        Select Case cell
        Case Is < 11
            cell.offset(0, 3).Value = "less then 11"
        Case Is < 21
            cell.offset(0, 3).Value = "less than 21"
        Case Else
            cell.offset(0, 3).Value = "big number"
        End Select
    Next
    For Each cell In Range("O13:O22")
        Select Case cell
        Case "TX", "CA"
            cell.offset(0, 1).Value = cell.offset(0, -1) * 1.5
        Case Else
            cell.offset(0, 1).Value = cell.offset(0, -1).Value
        End Select
    Next
End Sub

Sub FivePercent(rownumber, tax)
    Cells(rownumber, 16).Value = Cells(rownumber, 14).Value * tax
End Sub

Sub TenPercent(rownumber)
    Cells(rownumber, 16).Value = Cells(rownumber, 14) * 1.1
End Sub

Sub callaprocedure()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("6").Activate
    Dim rownumber As Integer
    rownumber = 25
    Do While Not (IsEmpty(Cells(rownumber, 15)))
        If Cells(rownumber, 15) = "TX" Then
            Call FivePercent(rownumber, 1.05)
        ElseIf Cells(rownumber, 15) = "FL" Then
            Call TenPercent(rownumber)
        ElseIf Cells(rownumber, 15) = "CA" Then
            Call FivePercent(rownumber, 1.05)
        Else
            Cells(rownumber, 16).Value = Cells(rownumber, 14)
        End If
        rownumber = rownumber + 1
    Loop
    'RM:  if statements are better no need to call; e.g. if cells Cells(rownumber, 15) = "TX" or _
    Cells(rownumber, 15) = "CA"
End Sub


