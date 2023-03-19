Attribute VB_Name = "water_GenerateCopy"
Option Explicit

Private Function lastRowByKey(cell As String) As Long

    lastRowByKey = Range(cell).End(xlDown).Row

End Function


Private Function lastRowByRowsCount(cell As String) As Long

    lastRowByRowsCount = Cells(Rows.Count, cell).End(xlUp).Row

End Function

Public Sub clearRowA()

    Columns("A:A").Select
    Selection.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
    Range("m2").Select

End Sub

Private Function lastRowByFind() As Long
    Dim lastRow As Long
    
    lastRow = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    
    lastRowByFind = lastRow
End Function

Private Sub DoCopy(lastRow As Long)
Attribute DoCopy.VB_ProcData.VB_Invoke_Func = " \n14"

    Range("F2:H" & lastRow).Select
    Selection.Copy
    
    Range("n2").Select
    ActiveSheet.Paste
    
    
    ' ¹°·®
    Range("L2:L" & lastRow).Select
    Selection.Copy
    
    Range("q2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("k2:k" & lastRow).Select
    Selection.Copy
    
    Range("r2").Select
    ActiveSheet.Paste
    
    Range("N14").Select
    Application.CutCopyMode = False
    
End Sub


Private Sub CleanSection(lastRow As Long)

    Range("n2:r" & lastRow).Select
    Selection.ClearContents
    Range("P14").Select
    
End Sub

Sub MainMoudleGenerateCopy()

    Dim lastRow As Long
        
    lastRow = lastRowByKey("I1")
    Call DoCopy(lastRow)


End Sub


Sub SubModuleCleanCopySection()

    Dim lastRow As Long
        
    lastRow = lastRowByKey("I1")
    Call CleanSection(lastRow)
    
End Sub

Sub insertRow()

    Dim lastRow As Long, i As Long, j As Long
    Dim selection_origin, selection_target As String
    
    'lastRow = lastRowByKey("A1")

    lastRow = lastRowByRowsCount("A")
    
    Rows(CStr(lastRow + 1) & ":" & CStr(lastRow + 2)).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    
    i = lastRowByKey("A1"): j = i + 2
    selection_origin = "A" & i & ":D" & i
    selection_target = "A" & i & ":D" & j
    
    Range(selection_origin).Select
    Selection.AutoFill Destination:=Range(selection_target), Type:=xlFillDefault
 
    selection_origin = "J" & i & ":L" & i
    selection_target = "J" & i & ":L" & j

    Range(selection_origin).Select
    Selection.AutoFill Destination:=Range(selection_target), Type:=xlFillDefault
    
    Range("R" & i).Select
    Selection.AutoFill Destination:=Range("R" & i & ":R" & j), Type:=xlFillDefault
    
    Application.CutCopyMode = False
    
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.LargeScroll Down:=-1


End Sub



