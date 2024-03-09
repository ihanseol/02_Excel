Option Explicit

Public Sub rows_and_column()
    Debug.Print Cells(20, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Debug.Print Range("a20").Row & " , " & Range("a20").Column
    
    Range("B2:Z44").Rows(3).Select
End Sub

Public Sub ShowNumberOfRowsInSheet1Selection()
    Dim area        As Range
    
    ' Worksheets("Sheet1").Activate
    Dim selectedRange As Excel.Range
    Set selectedRange = Selection
    
    Dim areaCount   As Long
    areaCount = Selection.Areas.count
    
    If areaCount <= 1 Then
        MsgBox "The selection contains " & _
               Selection.Rows.count & " rows."
    Else
        Dim areaIndex As Long
        areaIndex = 1
        For Each area In Selection.Areas
            MsgBox "Area " & areaIndex & " of the selection contains " & _
                   area.Rows.count & " rows." & " Selection 2 " & Selection.Areas(2).Rows.count & " rows."
            areaIndex = areaIndex + 1
        Next
    End If
End Sub


' Refactor 2023/10/20
Function myRandBetween(i As Integer, j As Integer, Optional div As Integer = 100) As Single
    Dim SIGN        As Integer
    
    SIGN = IIf(WorksheetFunction.RandBetween(0, 1), 1, -1)
    
    myRandBetween = (WorksheetFunction.RandBetween(i, j) / div) * SIGN
End Function

Function myRandBetween2(i As Integer, j As Integer, Optional div As Integer = 100) As Single
    Dim SIGN        As Integer
    
    myRandBetween = (WorksheetFunction.RandBetween(i, j) / div)
End Function


' Refactor 2023/10/20
Public Sub rnd_between()
    Dim i As Integer
    
    For i = 14 To 24
        Dim SIGN As Integer
        SIGN = IIf(WorksheetFunction.RandBetween(0, 1), 1, -1)
        
        Cells(i, 14).value = (WorksheetFunction.RandBetween(7, 12) / 100) * SIGN
        
        With Cells(i, 14)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .NumberFormatLocal = "0.00"
        End With
    Next i
End Sub


