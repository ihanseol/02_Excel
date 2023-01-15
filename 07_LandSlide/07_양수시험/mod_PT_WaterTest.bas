Attribute VB_Name = "mod_PT_WaterTest"
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

Public Function myRandBetween(i As Double, j As Double, Optional div As Double = 100) As Double
    Dim SIGN        As Integer
    
    If Application.WorksheetFunction.RandBetween(0, 1) Then
        SIGN = 1
    Else
        SIGN = -1
    End If
    
    myRandBetween = (Application.WorksheetFunction.RandBetween(i, j) / div) * SIGN
End Function

Public Function myRandBetween2(i As Double, j As Double, Optional div As Double = 100) As Double
    Dim SIGN        As Integer
    
    myRandBetween2 = (Application.WorksheetFunction.RandBetween(i, j) / div)
    
End Function

Public Sub rnd_between()
    Dim i, SIGN     As Integer
    
    For i = 14 To 24
        If Application.WorksheetFunction.RandBetween(0, 1) Then
            SIGN = 1
        Else
            SIGN = -1
        End If
        
        Cells(i, 14).Value = (Application.WorksheetFunction.RandBetween(7, 12) / 100) * SIGN
        Cells(i, 14).HorizontalAlignment = xlCenter
        Cells(i, 14).VerticalAlignment = xlCenter
        Cells(i, 14).NumberFormatLocal = "0.00"
    Next i
End Sub

Sub make_adjust_value()
    Dim i           As Integer
    
    For i = 14 To 23
        Cells(i, "h").Value = Round(myRandBetween(1, 3, 10), 1)
        Cells(i, "i").Value = myRandBetween(1, 3, 1)
        Cells(i, "j").Value = Round(myRandBetween(7, 13, 100), 2)
    Next i
End Sub

