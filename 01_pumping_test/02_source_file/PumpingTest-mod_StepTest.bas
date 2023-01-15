Attribute VB_Name = "mod_StepTest"
Option Explicit


Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

'»ç¶ûÇØ'
'this is a greate computer font i hava ever seen
'this is insanely greate

Sub step_pumping_test()
    Dim i As Integer

    Application.ScreenUpdating = False

    ActiveWindow.SmallScroll Down:=-33
    
    Range("D3:D7").Select
    Selection.Copy
    
    ActiveWindow.SmallScroll Down:=15
    Range("Q44").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    Selection.NumberFormatLocal = "0"
    Application.CutCopyMode = False
    
    
    Range("Q44:Q48").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    For i = 1 To 5
        Cells(43 + i, "Q").Value = Round(Cells(43 + i, "Q").Value, 0)
    Next i
    
    
    
    Range("A3:A7").Select
    Selection.Copy
    ActiveWindow.SmallScroll Down:=18
    Range("R44").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    Selection.NumberFormatLocal = "0.00"
    
    
    Range("B3:B7").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.SmallScroll Down:=24
    Range("S44").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    Selection.NumberFormatLocal = "0.00"
    
    Range("G3:G7").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.SmallScroll Down:=18
    Range("T44").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    Selection.NumberFormatLocal = "0.000"
    
    
    Range("F3:F7").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.SmallScroll Down:=15
    Range("U44").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    Application.CutCopyMode = False
   
    Range("T44:T48").Select
    Selection.NumberFormatLocal = "0.000"
    Range("S54").Select
    
    Application.ScreenUpdating = True
End Sub

Sub vertical_copy()


    Application.ScreenUpdating = False
    
    ActiveWindow.SmallScroll Down:=-6
    Range("Q44:Q48").Select
    Selection.Copy
    Range("Q51").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    
    Range("R44:R48").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("Q57").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    
    Range("S44:S48").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("Q63").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    
    Range("T44:T48").Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("Q69").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
        
    Application.CutCopyMode = False
    Selection.NumberFormatLocal = "0.000"
    Range("U44:U48").Select
    Selection.Copy
    
    Range("Q75").Select
    ActiveSheet.Paste
    ActiveWindow.SmallScroll Down:=0
    Range("R70").Select
    Application.CutCopyMode = False
    
    ActiveWindow.LargeScroll Down:=-2
    Range("P35").Select
    
    Application.ScreenUpdating = True

End Sub

Function get_chart_equation(ByVal chartname) As String
    
    Dim objTrendline As Trendline
    Dim strEquation As String
    
    With ActiveSheet.ChartObjects(chartname).Chart
        Set objTrendline = .SeriesCollection(1).Trendlines(1)
        With objTrendline
            .DisplayRSquared = False
            .DisplayEquation = True
            strEquation = .DataLabel.Text
        End With
    End With

    get_chart_equation = strEquation
    
End Function

Function split_string(ByVal name As String) As String()
    Dim myarray() As String
    
    myarray = Split(name)
    split_string = myarray
End Function

Sub get_chart7(ByRef c1 As Double, ByRef d1 As Double)
    
    Dim eq As String
    Dim t_array() As String
    Dim c, d As Double
    
    eq = get_chart_equation("Chart 7")
    t_array = split_string(eq)
    
    c = CDbl(t_array(2))
    d = CDbl(t_array(5))
    
    Range("p37").Value = c
    Range("p38").Value = d
       
    c1 = c
    d1 = d

End Sub

Sub get_chart8(ByRef c1 As Double, ByRef d1 As Double)
    
    Dim eq As String
    Dim t_array() As String
    Dim c, d As Double
    
    eq = get_chart_equation("Chart 8")
    t_array = split_string(eq)
    
    c = Abs(Round(CDbl(t_array(2)), 3))
    d = Round(CDbl(t_array(5)), 3)
    
    Range("q37").Value = c
    Range("q38").Value = d

    c1 = c
    d1 = d

End Sub

Sub ChangeCharts()
    ' Application.ScreenUpdating = False
    
    Dim myChart As ChartObject
    
    For Each myChart In ActiveSheet.ChartObjects
        myChart.Chart.Refresh
    Next myChart
    
    ' Application.ScreenUpdating = True
End Sub

Sub set_CB1()

    Dim c As Double
    Dim d As Double

    Call get_chart7(c, d)
    
    Range("a31").Value = c
    Range("b31").Value = d

End Sub

Sub set_CB2()

    Dim c As Double
    Dim d As Double

    Call get_chart8(c, d)
    
    Range("b38").Value = c
    Range("c38").Value = d
    Range("a38").Value = Range("d39").Value

End Sub


