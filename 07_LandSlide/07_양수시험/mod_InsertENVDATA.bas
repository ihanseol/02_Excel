Attribute VB_Name = "mod_InsertENVDATA"


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

Sub delete_dangye_column()
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
Sub Insert_DongHo_Data()
    Dim w()         As Variant
    Dim i           As Integer
    Dim index       As Variant
    
    index = Array(14, 19, 25, 29, 33, 37, 53, 57, 61, 77)
    
    w = Sheet15.Range("d14:f23").Value
    
    Range("H9").Value = "온도( ℃ )"
    Range("I9").Value = "EC (μs/㎝)"
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
    Selection.NumberFormatLocal = "G/표준"
End Sub

