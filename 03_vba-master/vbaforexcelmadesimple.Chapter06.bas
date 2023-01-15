Attribute VB_Name = "Chapter06"
Sub activecellpractice()
    'practice ActiveCell.Value and ActiveCell.Offset(1,0).Select with For Loop And If Statement
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("6").Activate
    Dim highnumber As Integer
    Dim lownumber As Integer
    highnumber = 10
    lownumber = 1
    Range("B1:B10").ClearContents
    For n = 1 To 10 Step 1
        Range("A" & n).Value = Int((highnumber - lownumber + 1) * Rnd + lownumber)
    Next
    Range("A1").Select
    For n = 1 To 10 Step 1
        If ActiveCell.Value > 5 Then
            Range("B" & n).Value = "greater than 5"
        End If
        ActiveCell.Offset(1, 0).Select
    Next
End Sub

Sub ifelseifelse()
    'practice ActiveCell.Value and ActiveCell.Offset(1,0).Select with For Loop And If Statement
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("6").Activate
    Dim highnumber As Integer
    Dim lownumber As Integer
    Dim biggestcolumnnumber As Integer
    highnumber = 100
    lownumber = 1
    biggestcolumnnumber = 100
    Range("F1:F" & biggestcolumnnumber).ClearContents
    For n = 1 To biggestcolumnnumber Step 1
        Range("E" & n).Value = Int((highnumber - lownumber + 1) * Rnd + lownumber)
    Next
    Range("E1").Select
    For n = 1 To biggestcolumnnumber Step 1
        If ActiveCell.Value > 80 Then
            Range("F" & n).Value = "greater than 80"
        ElseIf ActiveCell.Value > 50 Then
            Range("F" & n).Value = "greater than 50"
        ElseIf ActiveCell.Value > 40 Then
            Range("F" & n).Value = "greater than 40"
        Else
            Range("F" & n).Value = ""
        End If
        ActiveCell.Offset(1, 0).Select
    Next
End Sub

Sub fruitsandor()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("6").Activate
    Range("B13:B17").ClearContents
    columnnumber = 13
    For Each Cell In Range("A13:A17")
        If (Range("A" & columnnumber).Value = "apple") Or (Range("A" & columnnumber).Value = "orange") Then
            Range("B" & columnnumber).Value = "good fruit"
        ElseIf Range("A" & columnnumber).Value = "banana" Or Range("A" & columnnumber).Value = "peach" Then
            Range("B" & columnnumber).Value = "yummy"
        End If
        columnnumber = columnnumber + 1
    Next
    If Range("A19").Value = "yes" And Range("A20").Value = "no" Then
        Range("A21").Value = "yesno"
    End If
End Sub

Sub selectcase()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("6").Activate
    Range("H3").ClearContents
    Dim ticketprice As Integer
    Dim randomnumber As Integer
    ticketprice = 45
    randomnumber = Int((1000 - 1 + 1) * Rnd + 1)
    Range("H1").Value = ticketprice
    Select Case ticketprice
    Case 20
        Range("H2").Value = "cheap"
    Case 30
        Range("H2").Value = "affordable"
    Case 40 To 49.99
        Range("H2").Value = "no student can afford the range"
    Case 50
        Range("H2").Value = "higher class"
    Case 60
        Range("H2").Value = "Sclap for profit"
        Range("H3").Value = "Another statement"
    Case 100
        Range("H2").Value = "Too expensive"
    Case Else
        Range("H2").Value = "No ticket price comment"
    End Select
    Range("H5").Value = randomnumber
    Select Case randomnumber
    Case Is <= 100
        Range("H6").Value = "Too small"
    Case Is <= 250
        Range("H6").Value = "Small"
    Case 251 To 300
        Range("H6").Value = "Okay"
    Case 301 To 499
        Range("H6").Value = "Good"
    Case 500
        Range("H6").Value = "The Middle"
    Case Is <= 750
        Range("H6").Value = "Great"
    Case Is <= 999
        Range("H6").Value = "High"
    Case 1000
        Range("H6").Value = "Jackpot"
    Case Else
        Range("H6").Value = "Unknown"
    End Select
End Sub

