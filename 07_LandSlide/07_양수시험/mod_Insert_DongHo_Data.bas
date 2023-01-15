Attribute VB_Name = "mod_Insert_DongHo_Data"

Sub Make2880Document()
    Dim lang_code   As Long
    lang_code = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    
    Sheets("LongTest").Select
    Sheets("LongTest").Copy Before:=Sheet15
    
    If (Not Contains(Sheets, "out")) Then
        Sheets("LongTest (2)").name = "out"
    Else
        Sheets("out").Delete
        Sheets("LongTest (2)").name = "out"
    End If
    
    Application.GoTo Reference:="Print_Area"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    Application.CutCopyMode = False
    
    With Selection.Font
        .name = "¸¼Àº °íµñ"
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
    
    Call Insert_DongHo_Data
    Call delete_dangye_column
    
    Columns("G:I").Select
    
    ' 1042 - korean
    ' 1033 - english
    
    If lang_code = 1042 Then
        Selection.NumberFormatLocal = "G/Ç¥ÁØ"
    Else
        Selection.NumberFormatLocal = "G/General"
    End If
    
    Range("K13").Select
    
    Call AfterWork
End Sub

Sub AfterWork()
    ActiveWindow.View = xlPageBreakPreview
    Set ActiveSheet.HPageBreaks(1).Location = Range("A33")
    Set ActiveSheet.HPageBreaks(2).Location = Range("A56")
    Set ActiveSheet.HPageBreaks(3).Location = Range("A78")
    
    Range("A15").Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

Sub make1440sheet()
    Call delete_1440to2880
    Call make1440Timetable
End Sub

Private Sub make1440Timetable()
    'Range(Source & i).Formula = "=rounddown(" & Target & i & "*$P$6,0)"
    time_injection (54)
    time_injection (69)
    time_injection (73)
    time_injection (75)
    time_injection (77)
End Sub

Private Sub time_injection(ByVal ntime As Integer)
    Range("b" & CStr(ntime)).Formula = "=$B$10+(1440+C" & CStr(ntime) & ")/1440"
End Sub

Private Sub delete_dangye_column()
    Range("A1:A8").Select
    Selection.Cut
    Range("M1").Select
    ActiveSheet.Paste
    
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    
    Range("L1:L8").Select
    Selection.Cut
    Range("A1").Select
    ActiveSheet.Paste
End Sub

Private Sub delete_1440to2880()
    Rows("54:77").Select
    Selection.Delete Shift:=xlUp
    Range("L65").Select
    ActiveWindow.SmallScroll Down:=-12
End Sub

'before delete dangye data
Private Sub Insert_DongHo_Data()
    Dim w()         As Variant
    Dim i           As Integer
    Dim index       As Variant
    
    index = Array(14, 19, 25, 29, 33, 37, 53, 57, 61, 77)
    
    w = Sheet15.Range("d14:f23")
    
    Range("H9").Value = "¿Âµµ( ¡É )"
    Range("I9").Value = "EC (¥ìs/§¯)"
    Range("J9").Value = "pH"
    
    Range("H9:J9").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    For i = 1 To UBound(index) + 1
        Cells(index(i - 1), "h") = w(i, 1)
        Cells(index(i - 1), "i") = w(i, 2)
        Cells(index(i - 1), "j") = w(i, 3)
    Next i
    
    Columns("H:J").Select
    Selection.NumberFormatLocal = "G/Ç¥ÁØ"
End Sub

