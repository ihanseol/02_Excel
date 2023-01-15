Attribute VB_Name = "mod_CopyLongTerm"

Sub make_step_document()
    ' StepTest º¹»ç
    ' select last sheet -- Sheets(Sheets.Count).Select
    
    Application.ScreenUpdating = False
    
    Sheets("StepTest").Select
    Sheets("StepTest").Copy Before:=Sheet15
    
    Application.GoTo Reference:="Print_Area"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    
    Columns("J:AO").Select
    Selection.Delete Shift:=xlToLeft
    
    ActiveSheet.Shapes.Range(Array("CommandButton1")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton2")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("ComboBox1")).Select
    Selection.Delete
    
    Application.GoTo Reference:="Print_Area"
    With Selection.Font
        .name = "¸¼Àº °íµñ"
    End With
    
    Range("J19").Select
    
    ActiveWindow.View = xlPageBreakPreview
    ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
    
    If (Not Contains(Sheets, "Step")) Then
        Sheets("StepTest (2)").name = "Step"
    Else
        Sheets("Step").Delete
        Sheets("StepTest (2)").name = "Step"
    End If
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub

'2019/11/24

Sub modify_cell_value()
    Dim i           As Integer, j As Integer
    
    For i = 10 To 101
        Cells(i, "F").Value = Round(Cells(i, "F").Value, 2)
        Cells(i, "G").Value = Round(Cells(i, "G").Value, 2)
    Next i
End Sub

