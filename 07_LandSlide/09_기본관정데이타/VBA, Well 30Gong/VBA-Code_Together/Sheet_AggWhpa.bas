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
    
    Dim Q() As Double
    Dim daeSoo() As Double
    
    Dim T1() As Double
    Dim S1() As Double
    
    Dim direction() As Integer
    Dim gradient() As Double
    
    
    nofwell = GetNumberOfWell()
    If ActiveSheet.name <> "aggWhpa" Then Sheets("aggWhpa").Select
    
    ReDim Q(1 To nofwell) As Double
    ReDim daeSoo(1 To nofwell) As Double
    
    ReDim T1(1 To nofwell) As Double
    ReDim S1(1 To nofwell) As Double
    
    ReDim direction(1 To nofwell) As Integer
    ReDim gradient(1 To nofwell) As Double
      

    ' --------------------------------------------------------------------------------------
    
    For i = 1 To nofwell
        
        Sheets(CStr(i)).Select
        
        Q(i) = Sheets(CStr(i)).Range("c16").value
        daeSoo(i) = Sheets(CStr(i)).Range("c14").value
        
        T1(i) = Sheets(CStr(i)).Range("e7").value
        S1(i) = Sheets(CStr(i)).Range("g7").value
        
        direction(i) = getDirectionFromWell(i)
        gradient(i) = Sheets(CStr(i)).Range("k18").value
        
    Next i


    Sheets("aggWhpa").Select
    Call WriteWellData(Q, daeSoo, T1, S1, direction, gradient, nofwell)
    Call DrawOutline
    
    Range("a1").Select
    Application.CutCopyMode = False
    
End Sub

Sub DrawOutline()

    Application.ScreenUpdating = False
    
    Range("C3:O17").Select
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


Private Sub WriteWellData(Q As Variant, daeSoo As Variant, T1 As Variant, S1 As Variant, direction As Variant, gradient As Variant, ByVal nofwell As Integer)
    Dim i As Integer
    Dim t_sum As Double
    Dim daesoo_sum As Double
    Dim gradient_sum As Double
    Dim direction_sum As Double
    
    t_sum = 0
    daesoo_sum = 0
    gradient_sum = 0
    direction_sum = 0
    
    
    Application.ScreenUpdating = False
    Call EraseCellData("C4:O34")
            
    Call UnmergeAllCells
    
    For i = 1 To nofwell
    
        Cells(3 + i, "c").value = "W-" & CStr(i)
        
        Cells(3 + i, "e").value = Q(i)
        Cells(3 + i, "f").value = T1(i)
        t_sum = t_sum + T1(i)
        
        Cells(3 + i, "i").value = daeSoo(i)
        daesoo_sum = daesoo_sum + daeSoo(i)
        
        Cells(3 + i, "k").value = direction(i)
        direction_sum = direction_sum + direction(i)
        
        Cells(3 + i, "m").value = Format(gradient(i), "###0.0000")
        gradient_sum = gradient_sum + gradient(i)
        
        Cells(4, "d").value = "5년"
    
    Next i
    
   
    Cells(4, "g").value = Round(t_sum / nofwell, 4)
    Cells(4, "g").NumberFormat = "0.0000"
    Call merge_cells("d", nofwell)
    Call merge_cells("g", nofwell)
    
    Cells(4, "j").value = Round(daesoo_sum / nofwell, 1)
    Cells(4, "j").NumberFormat = "0.0"
    Call merge_cells("j", nofwell)
    
    Cells(4, "l").value = Round(direction_sum / nofwell, 1)
    Cells(4, "l").NumberFormat = "0.0"
    Call merge_cells("l", nofwell)
    
    Cells(4, "n").value = Round(gradient_sum / nofwell, 4)
    Cells(4, "n").NumberFormat = "0.0000"
    Call merge_cells("n", nofwell)
    
    Cells(4, "o").value = "무경계조건"
    Call merge_cells("o", nofwell)
    
    Cells(4, "h").value = 0.03
    Call merge_cells("h", nofwell)
    
    Application.ScreenUpdating = True
    
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






