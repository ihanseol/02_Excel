Option Explicit

Private Sub CommandButton1_Click()
    Sheets("aggWhpa").Visible = False
    Sheets("Well").Select
End Sub


Function getDirectionFromWell(i) As Integer

    If Sheets(CStr(i)).Range("k12").Font.Bold Then
        getDirectionFromWell = Sheets(CStr(i)).Range("k12").value
    Else
        getDirectionFromWell = Sheets(CStr(i)).Range("l12").value
    End If

End Function


Private Sub CommandButton2_Click()
' Collect Data

    Dim fName As String
    Dim nofwell, i As Integer
    
    Dim Q, daeSoo, T1, S1, direction, gradient As Double
    
    nofwell = sheets_count()
    If ActiveSheet.name <> "aggWhpa" Then Sheets("aggWhpa").Select
    Call EraseCellData("C4:O34")
    
    TurnOffStuff
    
    For i = 1 To nofwell
        Q = Sheets(CStr(i)).Range("c16").value
        daeSoo = Sheets(CStr(i)).Range("c14").value
        
        T1 = Sheets(CStr(i)).Range("e7").value
        S1 = Sheets(CStr(i)).Range("g7").value
        
        direction = getDirectionFromWell(i)
        gradient = Sheets(CStr(i)).Range("k18").value
        
        Call WriteWellData_Single(Q, daeSoo, T1, S1, direction, gradient, i)
    Next i
    
    Sheets("aggWhpa").Select
    
    Call MakeAverageAndMergeCells(nofwell)
    Call DrawOutline
    TurnOnStuff
    
End Sub

Private Sub WriteWellData_Single(Q As Variant, daeSoo As Variant, T1 As Variant, S1 As Variant, direction As Variant, gradient As Variant, ByVal i As Integer)
    
    Call UnmergeAllCells
        
    Cells(3 + i, "c").value = "W-" & CStr(i)
    Cells(3 + i, "e").value = Q
    Cells(3 + i, "f").value = T1
    Cells(3 + i, "i").value = daeSoo
    Cells(3 + i, "k").value = direction
    Cells(3 + i, "m").value = Format(gradient, "###0.0000")
    Cells(4, "d").value = "5년"
    
End Sub


Sub MakeAverageAndMergeCells(ByVal nofwell As Integer)
    Dim t_sum, daesoo_sum, gradient_sum, direction_sum As Double
    Dim i As Integer

    For i = 1 To nofwell
        t_sum = t_sum + Range("F" & (i + 3)).value
        daesoo_sum = daesoo_sum + Range("I" & (i + 3)).value
        direction_sum = direction_sum + Range("K" & (i + 3)).value
        gradient_sum = gradient_sum + Range("M" & (i + 3)).value
    Next i
    
    
    Cells(4, "g").value = Round(t_sum / nofwell, 4)
    Cells(4, "g").NumberFormat = "0.0000"
    
    Cells(4, "j").value = Round(daesoo_sum / nofwell, 1)
    Cells(4, "j").NumberFormat = "0.0"
        
    Cells(4, "l").value = Round(direction_sum / nofwell, 1)
    Cells(4, "l").NumberFormat = "0.0"
        
    Cells(4, "n").value = Round(gradient_sum / nofwell, 4)
    Cells(4, "n").NumberFormat = "0.0000"
       
    Cells(4, "o").value = "무경계조건"
    Cells(4, "h").value = 0.03
    
    Call merge_cells("d", nofwell)
    Call merge_cells("g", nofwell)
    Call merge_cells("j", nofwell)
    Call merge_cells("l", nofwell)
    Call merge_cells("n", nofwell)
    Call merge_cells("o", nofwell)
    Call merge_cells("h", nofwell)

End Sub


Sub merge_cells(cel As String, ByVal nofwell As Integer)

    Range(cel & CStr(4) & ":" & cel & CStr(nofwell + 3)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    
End Sub

Sub FindMergedCellsRange()
    Dim mergedRange As Range
    Set mergedRange = ActiveSheet.Range("A1").MergeArea
    MsgBox "The range of merged cells is " & mergedRange.Address
End Sub



Sub UnmergeAllCells()
    Dim ws As Worksheet
    Dim cell As Range
    
    Set ws = ActiveSheet
    
    For Each cell In ws.UsedRange
        If cell.MergeCells Then
            cell.MergeCells = False
        End If
    Next cell
    
    Call SetBorderLine_Default
End Sub


Sub SetBorderLine_Default()

    Range("C4:O17").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub


Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub


Sub DrawOutline()

    Application.ScreenUpdating = False
    
    Range("C3:O34").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("B31").Select
    
    Application.ScreenUpdating = True
End Sub






