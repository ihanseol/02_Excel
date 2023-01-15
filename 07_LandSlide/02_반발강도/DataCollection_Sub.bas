Attribute VB_Name = "DataCollection_Sub"
Option Explicit

'
'SECTION(1) = "B7"
'SECTION(2) = "E7"
'SECTION(3) = "H7"
'SECTION(4) = "K7"

Function GetInsertionPoint() As Integer

    Range("C300").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(-19, 0).Select
     
    GetInsertionPoint = ActiveCell.Row
    
End Function


Sub WriteSectionName(ByVal sheet As Integer, ByVal sector As Integer)
    Dim SECTION(1 To 4) As String
    Dim i, j As Integer
    
    ' 1 Sector --> 2
    ' 2 Sector --> 7
    ' 3 Sector --> 12
    ' 4 Sector --> 17
    
    SECTION(1) = "B7"
    SECTION(2) = "E7"
    SECTION(3) = "H7"
    SECTION(4) = "K7"
        
    i = GetInsertionPoint() + (sector - 1) * 5
    j = i + 4
    Range("B" & i).Value = Sheets(sheet).Range(SECTION(sector)).Value
    
End Sub



' it : block 1 sector
' it = 1, 2, 3, 4

Sub CellDecoration(ByVal sector As Integer)
    ' i : row start
    ' j : row end
    
    Dim i, j As Integer
    
    ' 1 Sector --> 2
    ' 2 Sector --> 7
    ' 3 Sector --> 12
    ' 4 Sector --> 17
    
    
    i = GetInsertionPoint() + (sector - 1) * 5
    j = i + 4
   
   
    'Range("B2:K6").Select
    Range("B" & i & ":K" & j).Select
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
        
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
    
    'Range("B2:B6").Select
    Range("B" & i & ":B" & j).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = False
    End With
    Selection.Merge
    
    'Range("G2:G6").Select
    Range("G" & i & ":G" & j).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = False
    End With
    Selection.Merge
    
    'Range("H2:H6").Select
    Range("H" & i & ":H" & j).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = False
    End With
    Selection.Merge
    
    'Range("I2:I6").Select
    Range("I" & i & ":I" & j).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = False
    End With
    Selection.Merge
    
    'Range("J2:J6").Select
    Range("J" & i & ":J" & j).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = False
    End With
    Selection.Merge
    
    'Range("K2:K6").Select
    Range("K" & i & ":K" & j).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = False
    End With
    Selection.Merge
    
    'Range("B2:K6").Select
    Range("B" & i & ":K" & j).Select
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
    
    
End Sub

Sub FinalDecoration()

    Cells.Select
    
    With Selection.Font
        .Name = "¸¼Àº°íµñ"
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
