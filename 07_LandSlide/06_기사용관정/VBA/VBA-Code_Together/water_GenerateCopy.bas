
' ***************************************************************
' water_GenerationCopy
'
' ***************************************************************




Option Explicit

Private Function lastRowByKey(cell As String) As Long
    lastRowByKey = Range(cell).End(xlDown).row
End Function


Private Function lastRowByRowsCount(cell As String) As Long
    lastRowByRowsCount = Cells(Rows.Count, cell).End(xlUp).row
End Function

Public Sub clearRowA()
    
End Sub

Private Function lastRowByFind() As Long
    Dim lastRow As Long
    
    lastRow = Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).row
    
    lastRowByFind = lastRow
End Function

Private Sub DoCopy(lastRow As Long)
    Range("F2:H" & lastRow).Select
    Selection.Copy
    
    Range("n2").Select
    ActiveSheet.Paste
    
    
    ' 물량
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



' return letter of range ...
Function Alpha_Column(Cell_Add As Range) As String
    Dim No_of_Rows As Integer
    Dim No_of_Cols As Integer
    Dim Num_Column As Integer
    
    No_of_Rows = Cell_Add.Rows.Count
    No_of_Cols = Cell_Add.Columns.Count
    
    If ((No_of_Rows <> 1) Or (No_of_Cols <> 1)) Then
        Alpha_Column = ""
        Exit Function
    End If
    
    Num_Column = Cell_Add.Column
    If Num_Column < 26 Then
        Alpha_Column = Chr(64 + Num_Column)
    Else
        Alpha_Column = Chr(Int(Num_Column / 26) + 64) & Chr((Num_Column Mod 26) + 64)
    End If
End Function


' Ctrl+D , Toggle OX, Toggle SINGO, HEOGA
Sub ToggleOX()
    Dim activeCellColumn, activeCellRow As String
    Dim row As Long
    Dim col As Long
    Dim lastRow As Long
    Dim cp, fillRange As String
    

    activeCellColumn = Split(ActiveCell.address, "$")(1)
    activeCellRow = Split(ActiveCell.address, "$")(2)
  
    row = ActiveCell.row
    col = ActiveCell.Column
    
    Debug.Print Alpha_Column(ActiveCell)
    
    If activeCellColumn = "S" Then
        If ActiveCell.Value = "O" Then
            ActiveCell.Value = "X"
        Else
            ActiveCell.Value = "O"
        End If
    End If
    

    If activeCellColumn = "B" Then
        If ActiveCell.Value = "신고공" Then
            ActiveCell.Value = "허가공"
            With Selection.Font
                .Color = -16776961
                .TintAndShade = 0
            End With
            Selection.Font.Bold = True
        Else
            ActiveCell.Value = "신고공"
             With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            Selection.Font.Bold = False
        End If
    End If
    
    If activeCellColumn = "D" Then
        cp = Replace(ActiveCell.address, "$", "")
        lastRow = lastRowByKey(ActiveCell.address)
        
        fillRange = "D" & Range(cp).row & ":D" & lastRow
        
        Range(cp).Select
        Selection.AutoFill Destination:=Range(fillRange)
        
        Range(cp).Select
    End If
    
    If activeCellColumn = "C" Then
        cp = Replace(ActiveCell.address, "$", "")
        lastRow = lastRowByKey(ActiveCell.address)
        
        fillRange = "C" & Range(cp).row & ":C" & lastRow
        
        Range(cp).Select
        Selection.AutoFill Destination:=Range(fillRange)
        
        Range(cp).Select
    End If
    
    
    If ActiveSheet.Name = "ss" And activeCellColumn = "K" Then
        UserForm_SS.Show
    End If
    
    If ActiveSheet.Name = "aa" And activeCellColumn = "K" Then
        UserForm_AA.Show
    End If
    
    If ActiveSheet.Name = "ii" And activeCellColumn = "K" Then
        UserForm_II.Show
    End If
End Sub


Sub MainMoudleGenerateCopy()
    Dim lastRow As Long
        
    lastRow = lastRowByKey("A1")
    Call DoCopy(lastRow)
End Sub


Sub SubModuleInitialClear()
    Dim lastRow As Long
    Dim userChoice As VbMsgBoxResult
    
    lastRow = lastRowByKey("A1")
  
    userChoice = MsgBox("Do you want to continue?", vbOKCancel, "Confirmation")

    If userChoice <> vbOK Then
        Exit Sub
    End If
    
    Range("e2:j" & lastRow).Select
    Selection.ClearContents
    Range("n2:r" & lastRow).Select
    Selection.ClearContents
    
    If lastRow >= 23 Then
        Rows("23:" & lastRow).Select
        Selection.Delete Shift:=xlUp
    End If
    
    
    If (ActiveSheet.Name = "ii") Then
        Range("l2").Value = 0
    End If
    
    Range("m2").Select
End Sub

Sub SubModuleCleanCopySection()
    Dim lastRow As Long
        
    lastRow = lastRowByKey("A1")
    Range("n2:r" & lastRow).Select
    Selection.ClearContents
    Range("P14").Select
End Sub


' 2023/4/19 - copy modify

Sub insertRow()
    Dim lastRow As Long, i As Long, j As Long
    Dim selection_origin, selection_target As String
    Dim AddingRowCount As Long
    
    'lastRow = lastRowByKey("A1")

    AddingRowCount = 10

    lastRow = lastRowByRowsCount("A")
    
    Rows(CStr(lastRow + 1) & ":" & CStr(lastRow + AddingRowCount)).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    
    i = lastRowByKey("A1")
    j = i + AddingRowCount
    
    selection_origin = "A" & i & ":D" & i
    selection_target = "A" & i & ":D" & j
    
    Range(selection_origin).Select
    Selection.AutoFill Destination:=Range(selection_target), Type:=xlFillDefault
 
    selection_origin = "K" & i & ":M" & i
    selection_target = "K" & i & ":M" & j

    Range(selection_origin).Select
    Selection.AutoFill Destination:=Range(selection_target), Type:=xlFillDefault
    
    Range("S" & i).Select
    Selection.AutoFill Destination:=Range("S" & i & ":S" & j), Type:=xlFillDefault
    
    Application.CutCopyMode = False
    
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.LargeScroll Down:=-1
    ActiveWindow.LargeScroll Down:=-1
End Sub





