Attribute VB_Name = "mod_W1StepTEST"
Option Explicit
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

Sub step_pumping_test()
    Dim i           As Integer
    
    Application.ScreenUpdating = False
    
    ' ----------------------------------------------------------------
    
    Range("D3:D7").Select
    Selection.Copy
    
    Range("Q44").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    Selection.NumberFormatLocal = "0"
    
    Range("Q44:Q48").Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
    End With
    
    
    Call CutDownNumber("Q", 0)
    
    ' ----------------------------------------------------------------
    
    Range("A3:A7").Select
    Selection.Copy
    Range("R44").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    Selection.NumberFormatLocal = "0.00"
    
    
    ' ----------------------------------------------------------------
    
    Range("B3:B7").Select
    Selection.Copy
      
    Range("S44").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    Selection.NumberFormatLocal = "0.00"
    
    ' ----------------------------------------------------------------

    Range("G3:G7").Select
    Selection.Copy
    Range("T44").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    ' Selection.NumberFormatLocal = "0.000"
    
    Call CutDownNumber("T", 3)
     Application.CutCopyMode = False
    ' ----------------------------------------------------------------
    
    Range("F3:F7").Select
    Selection.Copy
    Range("U44").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                           :=False, Transpose:=False
    
    Range("T44:T48").Select
    Selection.NumberFormatLocal = "0.000"
    
    ' ----------------------------------------------------------------
    
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
End Sub

Sub CutDownNumber(po As String, cutdown As Integer)
    Dim i, chrcode As Integer
    For i = 1 To 5
        Cells(i + 43, po).Value = Format(Round(Cells(i + 43, po).Value, cutdown), "###0.000")
    Next i
End Sub

Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .Value = ""
    End With
End Sub

Sub vertical_copy()
    Dim strValue(1 To 5) As String
    Dim result As String
    Dim i As Integer
    
    EraseCellData ("q64:x64")
    
    strValue(1) = "Q44:Q48"
    strValue(2) = "R44:R48"
    strValue(3) = "S44:S48"
    strValue(4) = "T44:T48"
    strValue(5) = "U44:U48"

    For i = 1 To 5
        result = ConcatenateCells(strValue(i))
        Cells(64, Chr(81 + i - 1)).Value = result
    Next i
    
    Range("q63").Select
    Call WriteStringEtc
End Sub


Sub WriteStringEtc()
    Dim i As Integer
    Dim cv1, cv2, cv3, append As String
    Dim arr() As Variant
    arr = Array(0, 120, 240, 360, 480)
    
    For i = 1 To 5
        If i = 5 Then
            append = ""
        Else
            append = vbLf
        End If
        
        cv1 = cv1 & CStr(i) & append
        cv2 = cv2 & CStr(arr(i - 1)) & append
        cv3 = cv3 & CStr(120) & append
    Next i
    
    Cells(64, "v").Value = cv1
    Cells(64, "w").Value = cv2
    Cells(64, "x").Value = cv3
End Sub

Function ConcatenateCells(inRange As String) As String
    Dim cell As Range
    Dim concatenatedValue As String
    Dim sFormat(1 To 5) As String
    Dim i As Integer
    
    
    sFormat(1) = "###0"
    sFormat(2) = "###0.00"
    sFormat(3) = "###0.00"
    sFormat(4) = "###0.000"
    sFormat(5) = "###0.000"
    
    i = Asc(Left(inRange, 1)) - Asc("P")
        
    For Each cell In Range(inRange)
        concatenatedValue = concatenatedValue & Format(cell.Value, sFormat(i)) & vbLf
    Next cell
    
     ConcatenateCells = Left(concatenatedValue, Len(concatenatedValue) - 1)
End Function


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
    Dim myarray()   As String
    
    myarray = Split(name)
    split_string = myarray
End Function

Sub get_chart7(ByRef c1 As Double, ByRef d1 As Double)
    Dim eq          As String
    Dim t_array()   As String
    Dim c, d        As Double
    
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
    Dim eq          As String
    Dim t_array()   As String
    Dim c, d        As Double
    
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
    Dim myChart     As ChartObject
    
    For Each myChart In ActiveSheet.ChartObjects
        myChart.Chart.Refresh
    Next myChart
End Sub

Sub set_CB1()
    Dim c           As Double
    Dim d           As Double
    
    Call get_chart7(c, d)
    
    Range("a31").Value = c
    Range("b31").Value = d
End Sub

Sub set_CB2()
    Dim c           As Double
    Dim d           As Double
    
    Call get_chart8(c, d)
    
    Range("b38").Value = c
    Range("c38").Value = d
    Range("a38").Value = Range("d39").Value
End Sub

