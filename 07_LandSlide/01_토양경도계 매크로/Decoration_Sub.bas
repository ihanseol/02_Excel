Attribute VB_Name = "Decoration_Sub"
Option Explicit

Function GetInsertionPoint() As Integer

    Range("C9999").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(-2, 0).Select
     
    GetInsertionPoint = ActiveCell.Row
    
End Function


Sub fillRestArea(ByVal i As Integer, ByVal j As Integer)
    
    Range("B" & i).Select
    ActiveCell.FormulaR1C1 = "»óºÎ"
    Range("B" & (i + 1)).Select
    ActiveCell.FormulaR1C1 = "ÁßºÎ"
    Range("B" & j).Select
    ActiveCell.FormulaR1C1 = "ÇÏºÎ"
    

End Sub


Sub CellDecoration()
Attribute CellDecoration.VB_ProcData.VB_Invoke_Func = " \n14"
    
    ' i : row start
    ' j : row end
    
    Dim i, j As Integer
    
    i = GetInsertionPoint()
    j = i + 2
    
    Call fillRestArea(i, j)
    
    Range("A" & i & ":N" & j).Select
    Application.CutCopyMode = False
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    
    ' Range("A2:A4").Select
    Range("A" & i & ":A" & j).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Selection.Merge
    
    'Range("M2:M4").Select
    Range("M" & i & ":M" & j).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Selection.Merge
    
    'Range("N2:N4").Select
    Range("N" & i & ":N" & j).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    Selection.Merge
    
    'Range("A2:N4").Select
    Range("A" & i & ":N" & j).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Range("o2").Select
    
End Sub

Sub FinalDecoration()
Attribute FinalDecoration.VB_ProcData.VB_Invoke_Func = " \n14"

    Cells.Select
    
    With Selection.Font
        .Name = "¸¼Àº °íµñ"
        .Size = 10
    End With
    With Selection
        .VerticalAlignment = xlCenter
    End With
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Range("T23").Select
End Sub
