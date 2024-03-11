Attribute VB_Name = "mod_W1_TESTVALUE"
Option Explicit

Public Sub rows_and_column()
    Debug.Print Cells(20, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Debug.Print Range("a20").Row & " , " & Range("a20").Column
    
    Range("B2:Z44").Rows(3).Select
End Sub

Public Sub ShowNumberOfRowsInSheet1Selection()
    Dim area        As Range
    
    Dim selectedRange As Excel.Range
    Set selectedRange = Selection
    
    Dim areaCount   As Long
    areaCount = Selection.Areas.Count
    
    If areaCount <= 1 Then
        MsgBox "The selection contains " & _
               Selection.Rows.Count & " rows."
    Else
        Dim areaIndex As Long
        areaIndex = 1
        For Each area In Selection.Areas
            MsgBox "Area " & areaIndex & " of the selection contains " & _
                   area.Rows.Count & " rows." & " Selection 2 " & Selection.Areas(2).Rows.Count & " rows."
            areaIndex = areaIndex + 1
        Next
    End If
End Sub



' Refactor 2023/10/20
Public Function myRandBetween(i As Double, j As Double, Optional div As Double = 100) As Double
    
    Dim SIGN        As Integer
    
    ' Random Generation from i to j
    ' div - devide  factor ...
    
    
    SIGN = IIf(WorksheetFunction.RandBetween(0, 1), 1, -1)
    myRandBetween = (Application.WorksheetFunction.RandBetween(i, j) / div) * SIGN

End Function


Public Function myRandBetween2(i As Double, j As Double, Optional div As Double = 100) As Double
    Dim SIGN        As Integer
    
    myRandBetween2 = (Application.WorksheetFunction.RandBetween(i, j) / div)
End Function


' Refactor 2023/10/20

Public Sub rnd_between()
    Dim i As Integer
    Dim SIGN As Integer
    
    For i = 14 To 24
        SIGN = IIf(WorksheetFunction.RandBetween(0, 1), 1, -1)
        Cells(i, 14).Value = (WorksheetFunction.RandBetween(7, 12) / 100) * SIGN
        
        With Cells(i, 14)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .NumberFormatLocal = "0.00"
        End With
    Next i
End Sub



