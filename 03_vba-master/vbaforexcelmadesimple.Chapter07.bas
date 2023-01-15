Attribute VB_Name = "Chapter07"
Sub foreachelement()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("7").Activate

End Sub

Sub foreachcell()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("7").Activate
    For Each Cell In Range("A1:A20")
        Cell.Value = Int((100 - 30 + 1) * Rnd + 30)
        Cell.Offset(0, 1).Value = Int((2 - 1 + 1) * Rnd + 1)
        If Cell.Offset(0, 1).Value = 1 Then
            Cell.Offset(0, 1).Value = "m"
        Else
            Cell.Offset(0, 1).Value = "f"
        End If
    Next
    For Each Cell In Range("A1:A20")
        If Cell.Value >= 90 And Cell.Offset(0, 1).Value = "f" Then
            female90count = female90count + 1
        ElseIf Cell.Value >= 70 And Cell.Offset(0, 1).Value = "f" Then
            female70count = female70count + 1
        ElseIf Cell.Value >= 50 And Cell.Offset(0, 1).Value = "f" Then
            female50count = female50count + 1
        ElseIf Cell.Value >= 30 And Cell.Offset(0, 1).Value = "f" Then
            female30count = female30count + 1
        End If
        If Cell.Value >= 90 And Cell.Offset(0, 1).Value = "m" Then
            male90count = male90count + 1
        ElseIf Cell.Value >= 70 And Cell.Offset(0, 1).Value = "m" Then
            male70count = male70count + 1
        ElseIf Cell.Value >= 50 And Cell.Offset(0, 1).Value = "m" Then
            male50count = male50count + 1
        ElseIf Cell.Value >= 30 And Cell.Offset(0, 1).Value = "m" Then
            male30count = male30count + 1
        End If
    Next
    Range("B22").Value = female90count
    Range("B23").Value = female70count
    Range("B24").Value = female50count
    Range("B25").Value = female30count
    Range("B26").Value = male90count
    Range("B27").Value = male70count
    Range("B28").Value = male50count
    Range("B29").Value = male30count
    Range("B30").Value = Application.sum(Range("B22:B29"))
End Sub

Sub fornext()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("7").Activate
    Dim number_n As Integer
    For n = 1 To 10 Step 1
        number_n = Int((10 - 1 + 1) * Rnd + 1)
        Range("D" & n).Value = number_n
        sum_n = sum_n + number_n
    Next
    Range("D12").Value = sum_n
End Sub

Sub fornext2()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("7").Activate
    Const threshold As Integer = 50
    Dim countrows As Integer
    'get count of contiguous rows
    countrows = Range("A1", Range("A1").End(xlDown)).Count
    Range("A1").Select
    For Count = 1 To countrows:
        If ActiveCell.Value > threshold Then
            ActiveCell.Offset(0, 2).Value = "It's in"
        End If
        ActiveCell.Offset(1, 0).Select
    Next
    Range("C1:C" & countrows).ClearContents
    'same as
    For Count = 1 To countrows:
        If Cells(Count, 1).Value > threshold Then
            Cells(Count, 3).Value = "It's in!"
        End If
    Next
    Range("C1:C" & countrows).ClearContents
End Sub

Sub exitforloop()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("7").Activate
    Const fruit As String = "orange"
    For Each Cell In Range("G1:G12")
        If Cell = fruit Then
            Cell.Offset(0, 1).Value = "Found " & fruit
            Exit For
        End If
    Next
End Sub

Sub dountilloop()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("7").Activate
    Dim inputnumber As Long
    Range("D20", Range("D20").End(xlDown)).ClearContents
    Range("D20").Select
    Do
        inputnumber = InputBox("Enter an integer.  Type quit to exit")
        ActiveCell.Value = inputnumber
        ActiveCell.Offset(1, 0).Select
    Loop Until inputnumber = 0
End Sub

Sub dowhileloop()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("7").Activate
    Dim inputnumber As Integer
    Range("D20", Range("D20").End(xlDown)).ClearContents
    Range("D20").Select
    'initiate do while loop inputnumber set to 1
    inputnumber = 1
    Do While inputnumber > 0
        inputnumber = InputBox("Enter an integer greater than zero. Negative number to exit")
        If inputnumber > 0 Then
            ActiveCell.Value = inputnumber
            ActiveCell.Offset(1, 0).Select
        End If
    Loop
End Sub

