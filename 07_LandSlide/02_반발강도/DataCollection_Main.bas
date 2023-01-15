Attribute VB_Name = "DataCollection_Main"
Option Explicit


Sub CopyDataFromSheet(i As Integer)
        
    Sheets(i).Activate
        
    Range("B58:J77").Select
    Selection.Copy
    
    Sheets("NewSheet").Activate
    Call MoveInsertionPoint
        
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
End Sub

Sub MoveInsertionPoint()

    Range("C300").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    
End Sub


Function GetWorksheet(shtName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = Worksheets(shtName)
End Function


Sub DoDataCollection()

    Dim cnt As Integer, i As Integer, j As Integer
    
    cnt = Worksheets.Count
        
    If Not GetWorksheet("NewSheet") Is Nothing Then
        Application.DisplayAlerts = False
        Worksheets("NewSheet").Delete
        Application.DisplayAlerts = True
        Sheets.Add(After:=Sheets(cnt - 1)).Name = "NewSheet"
        cnt = cnt - 1
    Else
        Sheets.Add(After:=Sheets(cnt)).Name = "NewSheet"
    End If
    
       
    For i = 1 To cnt
        Call CopyDataFromSheet(i)
        For j = 1 To 4
            Call CellDecoration(j)
            Call WriteSectionName(i, j) '(i :sheet number, j : sector number)
        Next
    Next
    
    Call FinalDecoration

End Sub
