Attribute VB_Name = "mod_MakeREPORT"



Sub DuplicateQ2Page(ByVal n As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q2")
    
    For i = 1 To n
        ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        ActiveSheet.name = "p" & i
        
        With ActiveSheet.Tab
            .ThemeColor = xlThemeColorAccent3
            .TintAndShade = 0
        End With
        
        Call SetWellPropertyQ2(i)
    Next i
End Sub


Sub make_step_document()
    ' StepTest ∫πªÁ
    ' select last sheet -- Sheets(Sheets.Count).Select
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("StepTest")
    
    Application.ScreenUpdating = False
    ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    
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
        .name = "∏º¿∫ ∞ÌµÒ"
    End With
    
    Range("J19").Select
    
    ActiveWindow.View = xlPageBreakPreview
    
    ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
    Set ActiveSheet.HPageBreaks(1).Location = Range("A31")
    
    
    If (Not Contains(Sheets, "Step")) Then
        Sheets("StepTest (2)").name = "Step"
    Else
        Sheets("Step").Delete
        Sheets("StepTest (2)").name = "Step"
    End If
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub



Sub Make2880Document()
    Dim lang_code   As Long
    Dim randomNumber As Integer
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("LongTest")
    
    lang_code = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    
    If (Not Contains(Sheets, "out")) Then
        Sheets("LongTest (2)").name = "out"
    Else
        Sheets("out").Delete
        Sheets("LongTest (2)").name = "out"
    End If
    
'    If IsSheetsHasA(ActiveSheet.name) Then
'        randomNumber = Int((100 * Rnd) + 1)
'        ActiveSheet.name = "2880_" & Format(CStr(randomNumber), "00")
'    Else
'        ActiveSheet.name = 2880
'    End If
    
    
    '---------------------------------------------------------------------------------
    Application.GoTo Reference:="Print_Area"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    Application.CutCopyMode = False
    
    With Selection.Font
        .name = "∏º¿∫ ∞ÌµÒ"
    End With
    
    Columns("K:AT").Select
    Selection.Delete Shift:=xlToLeft
    
    Range("N12").Select
    ActiveSheet.Shapes.Range(Array("CommandButton6")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton5")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton7")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("ComboBox1")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton2")).Select
    Selection.Delete
    
    Rows("102:336").Select
    Selection.Delete Shift:=xlUp
    
    Range("F109").Select
    ActiveWindow.SmallScroll Down:=-105
    
    Application.GoTo Reference:="Print_Area"
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    Call mod_InsertENVDATA.Insert_DongHo_Data
    Call mod_InsertENVDATA.delete_dangye_column
    
    Columns("G:I").Select
    
    ' 1042 - korean
    ' 1033 - english
    
    If lang_code = 1042 Then
        Selection.NumberFormatLocal = "G/«•¡ÿ"
    Else
        Selection.NumberFormatLocal = "G/General"
    End If
    
    Range("K13").Select
    Call AfterWork
End Sub



'2019/11/24

Sub modify_cell_value()
    Dim i           As Integer, j As Integer
    
    For i = 10 To 101
        Cells(i, "F").Value = Round(Cells(i, "F").Value, 2)
        Cells(i, "G").Value = Round(Cells(i, "G").Value, 2)
    Next i
End Sub

