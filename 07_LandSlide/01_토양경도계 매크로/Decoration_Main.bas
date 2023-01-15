Attribute VB_Name = "Decoration_Main"
Option Explicit

Private Sub dataCollection()
Attribute dataCollection.VB_ProcData.VB_Invoke_Func = " \n14"

    Range("B9:B18").Select
    Selection.Copy
    Range("B21").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Range("E9:E18").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B22").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Range("H9:H18").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("B23").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
        
End Sub

Private Sub EraseDataCollection()

    Range("B21:K26").Select
    Selection.ClearContents
    Range("G21").Select
    Selection.Copy
    Range("B21:B23").Select
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
End Sub


Private Sub CopyDataFromSheet(i As Integer)
        
    Sheets(i).Activate
    Call dataCollection
    
    Range("B21:K23").Select
    Selection.Copy
      
    Sheets("NewSheet").Select
    Call MoveInsertionPoint
    ActiveSheet.Paste
    Application.CutCopyMode = False
    
End Sub

Private Sub MoveInsertionPoint()

    Range("C9999").Select
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    
End Sub


Function GetWorksheet(shtName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = Worksheets(shtName)
End Function


Private Sub CopyRestData(sheet As Integer)
    Dim i As Integer
    
    i = GetInsertionPoint()
    Sheets("NewSheet").Range("M" & i).Value = Round(Sheets(sheet).Range("M34").Value, 1)
    Sheets("NewSheet").Range("N" & i).Value = Sheets(sheet).Range("J34").Value
    
    Range("A" & i).Value = Sheets(sheet).Name

End Sub


Sub EraseAllTempData()

    Dim cnt As Integer, i As Integer
    
    cnt = Worksheets.Count
    
    For i = 1 To cnt - 1
        Sheets(i).Activate
        Call EraseDataCollection
    Next i

End Sub

Sub MainProcedure()

    Dim cnt As Integer, i As Integer
    
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
        Call CopyRestData(i)
        Call CellDecoration
    Next
    
    Call FinalDecoration
    Call EraseAllTempData

End Sub



