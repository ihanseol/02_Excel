Attribute VB_Name = "mod_compare"
Option Explicit



Private Function lastRowByKey(cell As String) As Long

    lastRowByKey = Range(cell).End(xlDown).Row

End Function

Private Sub setcolor(ByVal nrow As Integer)

    Rows(CStr(nrow) & ":" & CStr(nrow)).Select
    
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
        .PatternTintAndShade = 0
    End With
End Sub



'if compare is false then return true

Private Function line_compare(ByVal nrow As Integer) As Boolean

    line_compare = False
    
    If (Cells(nrow, "a").Value <> Cells(nrow, "g")) Then
        line_compare = True
        Exit Function
    End If
    
    If (Cells(nrow, "b").Value <> Cells(nrow, "h")) Then
        line_compare = True
        Exit Function
    End If
    
    If (Cells(nrow, "c").Value <> Cells(nrow, "i")) Then
        line_compare = True
        Exit Function
    End If
    
    If (Cells(nrow, "d").Value <> Cells(nrow, "j")) Then
        line_compare = True
        Exit Function
    End If
    
    
End Function

Sub main()

Dim nlast As Long
Dim i As Integer


nlast = lastRowByKey("a5")

For i = 5 To nlast

    If line_compare(i) Then
        Call setcolor(i)
    End If

Next i



End Sub
