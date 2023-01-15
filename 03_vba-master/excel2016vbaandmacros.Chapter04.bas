Attribute VB_Name = "Chapter04"
Sub fornextloops()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("4").Activate
    'Column A print 1 to 10
    For i = 1 To 10
        Cells(i, 1).Value = i
    Next i
    'Column B check Column A even number
    For i = 1 To 10
        If Range("A" & i).Value Mod 2 = 0 Then        'Remainder after division or Python's % mod
            Range("B" & i).Value = "even"
            Range("A" & i & ":B" & i).Interior.ColorIndex = 4
        End If
    Next i
        
    'Count occupied rows in a column
    Range("A13").Value = Cells(Rows.Count, 1).End(xlUp).Row        'print 10 or 15, count rows occupied column 1
    Range("A14").Value = Cells(Rows.Count, 2).End(xlUp).Row        'print 10, count rows occupied column 2
    Range("A15").Value = Cells(Rows.Count, 3).End(xlUp).Row        'print 1, count rows occupied column 3
    
    'Column H check Column G blank cell
    finalrow = Cells(Rows.Count, 7).End(xlUp).Row
    Range("I1").Value = finalrow        'print 17
    For i = 1 To finalrow Step 1
        If IsEmpty(Range("G" & i)) Then
            Range("H" & i).Value = "Empty"
            Range("H" & i).Interior.Color = rgbRed
        End If
    Next i
    
    'Column K copy every tenth cell Column K
    finalrowk = Cells(Rows.Count, 11).End(xlUp).Row
    stepnumber = 10
    For i = 10 To finalrowk Step stepnumber
        Cells(i, 11).Copy Destination:=Cells(i, 12)        'copy next cell on right column L
    Next i
    
    'Column A delete s54 A24:A31
    For i = 31 To 24 Step -1
        If Cells(i, 1).Value = "s54" Then
            Rows(i).Delete
        End If
    Next i
    
    'Column K print 1 to 100
    For i = 1 To 100
        Cells(i, 11).Value = i
    Next i
    
    'Column C ending a for loop early
    For i = 1 To 10
        If Cells(i, 1).Value > 7 Then
            Exit For
        Else
            Cells(i, 3).Value = Range("A" & i).Value
        End If
    Next i
End Sub

Sub doloops()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("4").Activate
    'RM:  books poor explanation do loop
    Range("A30").Select
    ActiveCell.Offset(0, 2).Range("A1").Select        'moves right two cells from present selection
    ActiveCell.Offset(1, -2).Range("A1").Select        'moves down one cell left two cells from present selection
    ActiveCell.Offset(-1, 3).Range("A1").Select        'moves up one cell and right three cells
    ActiveCell.Offset(1, 0).Range("A1").Select        'moves down one cell from present selection
End Sub

Sub dowhileloops()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("4").Activate
    i = 1
    Do While i < 11
        Range("N" & i).Value = i
        i = i + 1
    Loop
    'Random Integer Range--> Int ((upperbound - lowerbound + 1) * Rnd + lowerbound)
    randomnumberend = Int((10 - 1 + 1) * Rnd + 1)
    Range("N12").Value = randomnumberend
    n = 0
    i = 1
    Do While n < randomnumberend
        n = Int((10 - 1 + 1) * Rnd + 1)
        Range("N" & i).Value = n
        i = i + 1
    Loop
End Sub

Sub dountilloops()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("4").Activate
    i = 1
    Do Until i = 11
        Range("O" & i).Value = i
        i = i + 1
    Loop
End Sub

Sub foreachloops()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("4").Activate
    'Range("K1", Range("K1").End(xlDown)).Select
    For Each cell In Range("K1", Range("K1").End(xlDown)):
        If cell.Value Mod 2 = 0 Then
            cell.Font.Bold = True
        End If
    Next cell
End Sub

Sub ifthenelse()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("4").Activate
    finalrow = Cells(Rows.Count, 7).End(xlUp).Row
    Range("I1").Value = finalrow        'print 17
    For i = 1 To finalrow Step 1
        If IsEmpty(Range("G" & i)) Then
            Range("H" & i).Value = "Empty"
            Range("H" & i).Interior.Color = rgbRed
        Else
            Range("H" & i).Value = "Occupied"
            Range("H" & i).Interior.Color = rgbLightBlue
        End If
    Next i
    For n = 1 To 10 Step 1
        Range("Q" & n).Value = Int((3 - 1 + 1) * Rnd + 1)
    Next n
    For n = 1 To 10 Step 1
        If Range("Q" & n) = 1 Then
            Range("R" & n) = "Fruit"
        ElseIf Range("Q" & n) = 2 Then
            Range("R" & n) = "Vegetable"
        ElseIf Range("Q" & n) = 3 Then
            Range("R" & n) = "Herb"
        Else
            Range("R" & n).Value = "Error"
        End If
    Next n
End Sub

Sub selectcaseend()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("4").Activate
    bottomrow = Cells(Rows.Count, 18).End(xlUp).Row
    For n = 1 To 10 Step 1
        Range("Q" & n).Value = Int((3 - 1 + 1) * Rnd + 1)
    Next n
    For n = 1 To bottomrow Step 1
        Select Case Range("Q" & n).Value
            Case 1
                Range("R" & n) = "Fruit"
            Case 2
                Range("R" & n) = "Vegetable"
            Case 3
                Range("R" & n) = "Herb"
        End Select
    Next n
    For n = 20 To 30 Step 1
        Range("Q" & n).Value = Int((100 - 1 + 1) * Rnd + 1)
    Next n
    For n = 20 To 30 Step 1
        Select Case Range("Q" & n).Value
            Case 1 To 20
                Range("R" & n) = "the 20s"
            Case 21 To 50
                Range("R" & n) = "the 50s"
            Case 51 To 75
                Range("R" & n) = "the 75s"
            Case Else
                Range("R" & n) = "big"
        End Select
    Next n
    For n = 20 To 30 Step 1
        Select Case Range("Q" & n).Value
            Case Is < 21
                Range("S" & n) = "the 20s"
            Case Is < 51
                Range("S" & n) = "the 50s"
            Case Is < 76
                Range("S" & n) = "the 75s"
            Case Else
                Range("S" & n) = "big"
        End Select
    Next n
    
End Sub
