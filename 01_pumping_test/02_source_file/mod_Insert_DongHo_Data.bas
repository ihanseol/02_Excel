Attribute VB_Name = "mod_Insert_DongHo_Data"

Sub Macro_MakeLongTermTest()
'
' Macro5 Macro
'
    Sheets("장기양수시험").Select
    Sheets("장기양수시험").Copy Before:=Sheets(13)
        
    Sheets("장기양수시험 (2)").name = "out"
    
    Application.Goto Reference:="Print_Area"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    With Selection.Font
        .name = "맑은 고딕"
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Columns("K:AT").Select
    Selection.Delete Shift:=xlToLeft
    Range("N12").Select
    ActiveSheet.Shapes.Range(Array("CommandButton6")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton5")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton7")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("Frame1")).Select
    Selection.Delete
    ActiveWindow.SmallScroll Down:=18
    ActiveSheet.Shapes.Range(Array("CommandButton2")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton1")).Select
    Selection.Delete
    ActiveWindow.SmallScroll Down:=69
    Rows("102:336").Select
    Selection.Delete Shift:=xlUp
    Range("J105").Select
    ActiveSheet.Previous.Select
    ActiveSheet.Next.Select
    Range("F109").Select
    ActiveWindow.SmallScroll Down:=-105
    
    Call Insert_DongHo_Data
    Call delete_dangye_column
    
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



Private Sub Insert_DongHo_Data()

Range("H9").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[4]C"
    Range("I9").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[4]C"
    Range("J9").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[4]C"
    Range("H9:J9").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("H14").Select
    ActiveCell.FormulaR1C1 = "='w1'!RC[-4]"
    Range("I14").Select
    ActiveCell.FormulaR1C1 = "='w1'!RC[-4]"
    Range("J14").Select
    ActiveCell.FormulaR1C1 = "='w1'!RC[-4]"
    Range("J15").Select

    Range("H19").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-4]C[-4]"
    Range("I19").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-4]C[-4]"
    Range("J19").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-4]C[-4]"
    Range("J20").Select
    ActiveWindow.SmallScroll Down:=6
    Range("H25").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-9]C[-4]"
    Range("I25").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-9]C[-4]"
    Range("J25").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-9]C[-4]"
    Range("J26").Select
    ActiveWindow.SmallScroll Down:=6
    Range("H29").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-12]C[-4]"
    Range("I29").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-12]C[-4]"
    Range("J29").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-12]C[-4]"
    Range("J30").Select
    ActiveWindow.SmallScroll Down:=3
    Range("H33").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-15]C[-4]"
    Range("I33").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-15]C[-4]"
    Range("J33").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-15]C[-4]"
    Range("H37").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-18]C[-4]"
    Range("I37").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-18]C[-4]"
    Range("J37").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-18]C[-4]"
    Range("J38").Select
    ActiveWindow.SmallScroll Down:=18
    Range("H53").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-33]C[-4]"
    Range("I53").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-33]C[-4]"
    Range("J53").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-33]C[-4]"
    Range("J54").Select
    
    ActiveWindow.SmallScroll Down:=6
    Range("H57").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-36]C[-4]"
    Range("I57").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-36]C[-4]"
    Range("J57").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-36]C[-4]"
    Range("J58").Select
    ActiveWindow.SmallScroll Down:=9

    Range("H61").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-39]C[-4]"
    Range("I61").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-39]C[-4]"
    Range("J61").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-39]C[-4]"
    Range("J62").Select
    ActiveWindow.SmallScroll Down:=6
    Range("H77").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-54]C[-4]"
    Range("I77").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-54]C[-4]"
    Range("J77").Select
    ActiveCell.FormulaR1C1 = "='w1'!R[-54]C[-4]"
    Range("H78:J78").Select
    ActiveWindow.SmallScroll Down:=-54
    Columns("H:J").Select
    Selection.NumberFormatLocal = "G/표준"
    Range("I12").Select

End Sub

