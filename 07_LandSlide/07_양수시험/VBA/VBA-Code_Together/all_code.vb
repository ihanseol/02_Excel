
Option Explicit


' pumping test
Private Sub CommandButton1_Click()
    Call step_pumping_test
    Call vertical_copy
End Sub

Private Sub CommandButton2_Click()
    Call set_CB1
End Sub

Private Sub CommandButton3_Click()
    Call set_CB2
End Sub

Private Sub CommandButton4_Click()
    Call make_step_document
End Sub

Private Sub CommandButton5_Click()
    'Call make_long_document
    Call Make2880Document
End Sub

Private Sub CommandButton6_Click()
    Dim gong As Integer
    Dim KeyCell As Range
    
    Call adjustChartGraph
    
    Set KeyCell = Range("J48")
    
    gong = Val(CleanString(KeyCell.Value))
    Call SetChartTitleText(gong)
End Sub

Private Sub CommandButton7_Click()
    Call Make2880Document
    Call make1440sheet
End Sub


Private Sub CommandButton8_Click()
    Call set_CB_ALL
End Sub

Private Sub Worksheet_Activate()
  
'  Dim gong As Integer
'  Dim KeyCell As Range
'
'  Set KeyCell = Range("J48")
'
'  gong = Val(CleanString(KeyCell.Value))
'  Call SetChartTitleText(gong)

End Sub



Private Sub SetChartTitleText(ByVal i As Integer)

    Call SetGONGBEON

    ActiveSheet.ChartObjects("Chart 7").Activate
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "양수량(㎥/day)(W-" & CStr(i) & ")"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "양수량(㎥/day)(W-" & CStr(i) & ")"
    
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlValue).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "비수위강하량(day/㎡)"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "비수위강하량(day/㎡)"
    

    ActiveSheet.ChartObjects("Chart 5").Activate
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "양수량(㎥/day)(W-" & CStr(i) & ")"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "양수량(㎥/day)(W-" & CStr(i) & ")"
    
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlValue).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "비수위강하량(day/㎡)"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "비수위강하량(day/㎡)"
    
    ActiveSheet.ChartObjects("Chart 9").Activate
    ActiveChart.Axes(xlCategory).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "양수량(Q)(W-" & CStr(i) & ")"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "양수량(Q)(W-" & CStr(i) & ")"
    
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlValue).AxisTitle.Select
    ActiveChart.Axes(xlValue, xlPrimary).AxisTitle.Text = "수위강하량(Sw)"
    Selection.Format.TextFrame2.TextRange.Characters.Text = "수위강하량(Sw)"
    
End Sub








Private Sub CommandButton1_Click()
    Call show_gachae
End Sub


Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
Private Sub CommandButton1_Click()
    Call hide_gachae
End Sub

Private Sub Worksheet_Activate()

    If (Range("B14").Value < Range("B15").Value) Then
        Call cellRED
    Else
        Call cellGREEN
    End If
    
    Range("D15").Select

End Sub


Sub cellRED()
    Range("A15:B15").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 192
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub cellGREEN()

    Range("A15:B15").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With

End Sub


Private Sub CommandButton1_Click()
    Call findAnswer_StepTest
End Sub

Private Sub CommandButton2_Click()
    Call check_StepTest
End Sub


' Time Difference
Private Sub CommandButton3_Click()
    Dim mm As Integer
    Dim dtLongTerm, dtStempTime As Date
    
    mm = ComboBox1.Value
    
    ' 장기양수시험 시작시간
    dtLongTerm = Sheets("LongTest").Range("c10").Value
    dtStepTime = dtLongTerm - mm / 1440
    
    Range("c12").Value = dtStepTime
End Sub

Private Sub CommandButton4_Click()
    Dim dtToday, ntime, nDate As Date
    
'    dtToday = Date
'    ntime = TimeSerial(10, 0, 0)
'    nDate = dtToday + ntime
'
'    Range("c12").Value = nDate
    
    UserFormTS2.Show
End Sub

Private Sub Worksheet_Activate()
    Dim arr() As Variant
    Dim i As Integer
    
    ' arr = Array(250, 260, 270, 300, 360, 370, 380, 390, 420, 480, 490, 500, 510, 540, 600, 640, 700)
    arr = Array(600, 640, 700, 730, 800, 830, 900, 930, 1000, 1030, 1100, 1130, 1200, 1440)
        
    If (ActiveSheet.name <> "StepTest") Then Exit Sub
    
    
    If ComboBox1.Value <> arr(UBound(arr)) Then
        ComboBox1.Clear
        For i = LBound(arr) To UBound(arr)
            ComboBox1.AddItem (arr(i))
        Next i
        ComboBox1.Value = arr(UBound(arr))
    End If
    
End Sub




Private Sub CommandButton2_Click()
    UserFormSTime.Show
    ' Call frame_time_setting
    Call TimeSetting
    
    ActiveWindow.SmallScroll Down:=-66
    Range("O10").Select
End Sub

Private Sub frame_time_setting()
    Dim i As Integer
    Dim dStableTime As Integer
    
    Call initDictionary
    
    dStableTime = CInt(shW_LongTEST.ComboBox1.Value)
    MY_TIME = gDicStableTime(dStableTime)

End Sub

Private Sub CommandButton3_Click()
    Call set_daydifference
End Sub

Private Sub CommandButton4_Click()
    Call findAnswer_LongTest
End Sub

Private Sub CommandButton5_Click()
    Call resetValue
End Sub

Private Sub CommandButton6_Click()
    UserFormTS.Show
End Sub

Private Sub CommandButton7_Click()
    Call check_LongTest
End Sub



Private Sub Worksheet_Activate()
    Dim gong1, gong2 As String
    Dim gong, occur As Long
    
    
    Debug.Print ActiveSheet.name
    Call initDictionary
    
    If ActiveSheet.name <> "LongTest" Then
        Exit Sub
    End If
       
    If MY_TIME = 0 Then
        MY_TIME = initialize_myTime
        shW_LongTEST.ComboBox1.Value = gDicMyTime(MY_TIME)
    End If

'   gong = Val(CleanString(shInput.Range("J48").Value))
'   gong1 = "W-" & CStr(gong)
'   gong2 = shInput.Range("i54").Value
'
'   If gong1 <> gong2 Then
'        shInput.Range("i54").Value = gong1
'   End If
    
End Sub


Private Sub CommandButton1_Click()
    Call recover_01
End Sub

Private Sub CommandButton2_Click()
    Call save_original
End Sub

Private Sub Worksheet_Activate()
    Dim gong1, gong2 As String
    Dim gong As Long
    Dim er As Integer
    Dim cellformula As String
    

'    gong = Val(CleanString(shInput.Range("J48").Value))
'
'    gong1 = "W-" & CStr(gong)
'    gong2 = shInput.Range("i54").Value
'
'    If gong1 <> gong2 Then
'        'MsgBox "different : " & g1 & " g2 : " & g2
'        shInput.Range("i54").Value = gong1
'    End If
    

    er = GetEffectiveRadius
        
     Select Case er
        Case erRE1
            cellformula = "=SkinFactor!K8"
        
        Case erRE2
            cellformula = "=SkinFactor!K9"
            
        Case erRE3
            cellformula = "=SkinFactor!K10"
        
        Case Else
            cellformula = "=SkinFactor!C8"
    End Select
    
    Range("A28").Formula = cellformula
    
End Sub



Private Sub CommandButton1_Click()
    Call janggi_01
End Sub

Private Sub CommandButton2_Click()
    Call janggi_02
End Sub

Private Sub CommandButton3_Click()
    Call save_original
End Sub

Private Sub CommandButton4_Click()
    
    Call ToggleWellRadius
End Sub



'0 : skin factor, cell, C8
'1 : Re1,         cell, E8
'2 : Re2,         cell, H8
'3 : Re3,         cell, G10


Private Sub ToggleWellRadius()
    Dim er As Integer
    Dim cellformula As String
    
    er = GetEffectiveRadius
        
     Select Case er
        Case erRE1
            cellformula = "=SkinFactor!K8"
        
        Case erRE2
            cellformula = "=SkinFactor!K9"
            
        Case erRE3
            cellformula = "=SkinFactor!K10"
        
        Case Else
            cellformula = "=SkinFactor!C8"
    End Select

    If (Range("A27").Formula = cellformula) Then
        Range("A27").Formula = 0
    Else
        Range("A27").Formula = cellformula
    End If

End Sub

Private Sub SetEffectiveRadius()
     er = GetEffectiveRadius
            
     Select Case er
        Case erRE1
            cellformula = "=SkinFactor!K8"
        
        Case erRE2
            cellformula = "=SkinFactor!K9"
            
        Case erRE3
            cellformula = "=SkinFactor!K10"
        
        Case Else
            cellformula = "=SkinFactor!C8"
    End Select

    Range("A27").Formula = cellformula
End Sub


Private Sub Worksheet_Activate()
'    Dim gong1, gong2 As String
'    Dim gong As Long
'
'    gong = Val(CleanString(shInput.Range("J48").Value))
'    gong1 = "W-" & CStr(gong)
'    gong2 = shInput.Range("i54").Value
'
'    If gong1 <> gong2 Then
'        'MsgBox "different : " & g1 & " g2 : " & g2
'        shInput.Range("i54").Value = gong1
'    End If
    
    Call SetEffectiveRadius
End Sub



Private Sub CommandButton1_Click()
    Call step_01
End Sub

Private Sub CommandButton2_Click()
    Call save_original
End Sub

Private Sub Worksheet_Activate()
'    Dim gong1, gong2 As String
'    Dim gong As Long
'
'    gong = Val(CleanString(shInput.Range("J48").Value))
'
'    gong1 = "W-" & CStr(gong)
'    gong2 = shInput.Range("i54").Value
'
'    If gong1 <> gong2 Then
'        shInput.Range("i54").Value = gong1
'    End If
'
End Sub


Private Sub CommandButton1_Click()
    UserFormTS1.Show
End Sub


Private Sub CommandButton2_Click()
    Dim i As Integer
    
    For i = 14 To 23
        ' Temp
        Cells(i, "h").Value = Round(myRandBetween(1, 3, 10), 1)
        
        ' EC
        Cells(i, "i").Value = myRandBetween(1, 3, 1)
        
        ' PH
        Cells(i, "j").Value = Round(myRandBetween(7, 13, 100), 2)
    Next i

End Sub

Private Sub CommandButton3_Click()

    Range("L14:N23").Select
    Selection.Copy
    Range("H14").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("K9").Select
    Application.CutCopyMode = False

End Sub


Private Sub SetWellTitle(ByVal gong As Integer)

    Dim strText As String
    
    strText = "W-" & CStr(gong)
    
    Range("b4").Value = "수질 " & CStr(gong) & "번"
    Range("c4").Value = strText
    Range("d12").Value = strText
    Range("h12").Value = strText
    Range("l12").Value = strText
    
End Sub

'
' Random Generator
'

Private Sub CommandButton4_Click()
' 2024,03,11
' Random Generation by Button ...


    Dim i As Integer
    
    For i = 14 To 23
        'Temperature
        Cells(i, "L").Value = myRandBetween(1, 3, 10)
        
        'EC
        Cells(i, "M").Value = myRandBetween(1, 20, 1)
        
        'PH
        Cells(i, "N").Value = myRandBetween(8, 12, 100)
    Next i
    
End Sub




Private Sub Workbook_Open()
      
    'Sheet6.Activate
    sh01_StepSelect.name = "Step.Select"
    
    'Sheet7.Activate
    sh02_JanggiSelect.name = "Janggi.Select"
    
    'Sheet71.Activate
    sh03_RecoverSelect.name = "Recover.Select"
       
       
    With shW_LongTEST.ComboBox1
        .AddItem "60"
        .AddItem "75"
        .AddItem "90"
        .AddItem "105"
        .AddItem "120"
        .AddItem "140"
        .AddItem "160"
        .AddItem "180"
        .AddItem "240"
        .AddItem "300"
        .AddItem "360"
        .AddItem "420"
        .AddItem "480"
        .AddItem "540"
        .AddItem "600"
        .AddItem "660"
        .AddItem "720"
        .AddItem "780"
        .AddItem "840"
        .AddItem "900"
        .AddItem "960"
        .AddItem "1020"
        .AddItem "1080"
        .AddItem "1140"
        .AddItem "1200"
        .AddItem "1260"
        .AddItem "1320"
        .AddItem "1380"
        .AddItem "1440"
        .AddItem "1500"
    End With
   
    Call initDictionary
    ' Call GotoTopPosition
    
End Sub


'Private Sub GotoTopPosition()
'
'    Dim sht As Worksheet
'
'    For Each sht In Application.Worksheets
'        sht.Activate
'        Application.GoTo Reference:=Range("a1"), Scroll:=True
'    Next sht
'
'    shInput.Activate
'
'End Sub
'
Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
Option Explicit

' 2019/11/27 --- adjustment of chart's graph position and x, y scale

Sub adjustChartGraph()
    Dim Q0, Q1, E0, E1, SwQ0, SwQ1, IQ As Double
    
    ' IQ -- Initial Q
    ' SafeYield이다. 양수량
    
    Q0 = Range("D3").Value
    Q1 = Range("D7").Value
    
    E0 = Range("F35").Value
    E1 = Range("F32").Value
    
    SwQ0 = Range("F3").Value
    SwQ1 = Range("F7").Value
    
    IQ = Range("M51").Value
    
    Call setAxisScale("Chart 5", Q0, Q1, SwQ0, SwQ1)
    Call setAxisScale("Chart 7", Q0, Q1, SwQ0, SwQ1)
    
    Call setAxisScale_Efficiency("Chart 8", Q0, Q1, E0, E1)
    
    Call SetGONGBEON
End Sub


Public Sub SetGONGBEON()
    Dim gong As Integer
                  
    If ActiveSheet.name = "Input" Then
        gong = Val(CleanString(Range("J48").Value))
        Range("i54").Value = "W-" & gong
    End If
End Sub



Function determinX(ByVal x0 As Double, ByVal x1 As Double) As Double
    determinX = (x1 - x0) / 3
End Function

Function determinY(ByVal y0 As Double, ByVal y1 As Double) As Double
    'determiney 수정 - 2020-6-21
    'y0 = Round(y0 / 10, 0) * 10
    'y1 = Round(y1 / 10, 0) * 10
    
    determinY = (y1 - y0) / 3
End Function

Sub setAxisScale_Efficiency(strName As String, ByVal x0 As Double, ByVal x1 As Double, ByVal y0 As Double, ByVal y1 As Double)
    Dim dresx, dresy As Double
    Dim xMax, xMin, yMax, yMin As Double
    
    dresx = determinX(x0, x1)
    
    xMin = (x0 - dresx)
    xMax = (x1 + dresx)
    
    yMin = WorksheetFunction.RoundDown(y0, -1) - 20
    yMax = WorksheetFunction.RoundUp(y1, -1) + 10
    
    ActiveSheet.ChartObjects(strName).Activate
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MinimumScale = xMin
    ActiveChart.Axes(xlCategory).MaximumScale = xMax
    
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = yMin
    ActiveChart.Axes(xlValue).MaximumScale = yMax
    
    Call setAxisUnit(strName, xMin, xMax, yMin, yMax)
End Sub

Sub setAxisScale(strName As String, ByVal x0 As Double, ByVal x1 As Double, ByVal y0 As Double, ByVal y1 As Double)
    Dim dresx, dresy As Double
    Dim xMax, xMin, yMax, yMin As Double
    
    dresx = determinX(x0, x1)
    dresy = determinY(y0 * 1000, y1 * 1000)
    
    xMin = (x0 - dresx)
    xMax = (x1 + dresx)
    
    yMin = (y0 * 1000 - dresy) / 1000
    yMax = (y1 * 1000 + dresy) / 1000
    
    ActiveSheet.ChartObjects(strName).Activate
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MinimumScale = xMin
    ActiveChart.Axes(xlCategory).MaximumScale = xMax
    
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = yMin
    ActiveChart.Axes(xlValue).MaximumScale = yMax
    
    Call setAxisUnit(strName, xMin, xMax, yMin, yMax)
End Sub

Sub setAxisUnit(strName As String, ByVal x0 As Double, ByVal x1 As Double, ByVal y0 As Double, ByVal y1 As Double)
    Dim dresx, dresy As Double
    Dim xMax, xMin, yMax, yMin As Double
    
    ActiveSheet.ChartObjects(strName).Activate
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MajorUnit = (x1 - x0) / 10
    
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MajorUnit = (y1 - y0) / 4
End Sub



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
    ' StepTest 복사
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
        .name = "맑은 고딕"
    End With
    
    Range("J19").Select
    
    ActiveWindow.View = xlPageBreakPreview
    
    Set ActiveSheet.HPageBreaks(1).Location = Range("A31")
    
    
    
    If (Not Contains(Sheets, "Step")) Then
        Sheets("StepTest (2)").name = "Step"
    Else
        Sheets("Step").Delete
        Sheets("StepTest (2)").name = "Step"
    End If
    
     ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
     
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
        .name = "맑은 고딕"
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
        Selection.NumberFormatLocal = "G/표준"
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

Option Explicit

Public Sub rows_and_column()
    Debug.Print Cells(20, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Debug.Print Range("a20").Row & " , " & Range("a20").Column
    
    Range("B2:Z44").Rows(3).Select
End Sub

Public Sub ShowNumberOfRowsInSheet1Selection()
    Dim area        As Range
    
    Dim selectedRange As Excel.Range
    Set selectedRange = Selection
    
    Dim areaCount   As Long
    areaCount = Selection.Areas.Count
    
    If areaCount <= 1 Then
        MsgBox "The selection contains " & _
               Selection.Rows.Count & " rows."
    Else
        Dim areaIndex As Long
        areaIndex = 1
        For Each area In Selection.Areas
            MsgBox "Area " & areaIndex & " of the selection contains " & _
                   area.Rows.Count & " rows." & " Selection 2 " & Selection.Areas(2).Rows.Count & " rows."
            areaIndex = areaIndex + 1
        Next
    End If
End Sub



' Refactor 2023/10/20
Public Function myRandBetween(i As Double, j As Double, Optional div As Double = 100) As Double
    
    Dim SIGN        As Integer
    
    ' Random Generation from i to j
    ' div - devide  factor ...
    
    
    SIGN = IIf(WorksheetFunction.RandBetween(0, 1), 1, -1)
    myRandBetween = (Application.WorksheetFunction.RandBetween(i, j) / div) * SIGN

End Function


Public Function myRandBetween2(i As Double, j As Double, Optional div As Double = 100) As Double
    Dim SIGN        As Integer
    
    myRandBetween2 = (Application.WorksheetFunction.RandBetween(i, j) / div)
End Function


' Refactor 2023/10/20

Public Sub rnd_between()
    Dim i As Integer
    Dim SIGN As Integer
    
    For i = 14 To 24
        SIGN = IIf(WorksheetFunction.RandBetween(0, 1), 1, -1)
        Cells(i, 14).Value = (WorksheetFunction.RandBetween(7, 12) / 100) * SIGN
        
        With Cells(i, 14)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .NumberFormatLocal = "0.00"
        End With
    Next i
End Sub



Global WB_NAME      As String

Public Function MyDocsPath() As String
    MyDocsPath = Environ$("USERPROFILE") & "\" & "Documents"
    Debug.Print MyDocsPath
End Function

Public Function WB_HEAD() As String
    Dim num As Integer
    
    num = GetNumbers(Worksheets("Input").Range("I54").Value)
    
    If num >= 10 Then
        WB_HEAD = MyDocsPath + "\" + Left(ThisWorkbook.name, 6)
    Else
        WB_HEAD = MyDocsPath + "\" + Left(ThisWorkbook.name, 5)
    End If
    
    Debug.Print WB_HEAD
End Function

Sub janggi_01()
    ActiveWorkbook.SaveAs fileName:= _
                          WB_HEAD + "_janggi_01.dat", FileFormat _
                          :=xlTextPrinter, CreateBackup:=False
End Sub

Sub janggi_02()
    ActiveWorkbook.SaveAs fileName:= _
                          WB_HEAD + "_janggi_02.dat", FileFormat _
                          :=xlTextPrinter, CreateBackup:=False
End Sub

Sub recover_01()
    Debug.Print WB_HEAD
    ActiveWorkbook.SaveAs fileName:= _
                          WB_HEAD + "_recover_01.dat", FileFormat:= _
                          xlTextPrinter, CreateBackup:=False
End Sub

Sub step_01()
    Range("a1").Select
    
    ActiveWorkbook.SaveAs fileName:= _
                          WB_HEAD + "_step_01.dat", FileFormat:= _
                          xlTextPrinter, CreateBackup:=False
End Sub

Sub save_original()
    ActiveWorkbook.SaveAs fileName:=WB_HEAD + "_OriginalSaveFile", FileFormat:= _
                          xlOpenXMLWorkbookMacroEnabled, CreateBackup:=False
End Sub



Public MY_TIME      As Integer

Public gDicStableTime As Scripting.Dictionary
Public gDicMyTime   As Scripting.Dictionary

Sub initDictionary()
    Set gDicStableTime = New Scripting.Dictionary
    Set gDicMyTime = New Scripting.Dictionary
    
    gDicStableTime.Add Key:=60, item:=17
    gDicStableTime.Add Key:=75, item:=18
    gDicStableTime.Add Key:=90, item:=19
    gDicStableTime.Add Key:=105, item:=20
    gDicStableTime.Add Key:=120, item:=21
    gDicStableTime.Add Key:=140, item:=22
    gDicStableTime.Add Key:=160, item:=23
    gDicStableTime.Add Key:=180, item:=24
    gDicStableTime.Add Key:=240, item:=25
    gDicStableTime.Add Key:=300, item:=26
    gDicStableTime.Add Key:=360, item:=27
    gDicStableTime.Add Key:=420, item:=28
    gDicStableTime.Add Key:=480, item:=29
    gDicStableTime.Add Key:=540, item:=30
    gDicStableTime.Add Key:=600, item:=31
    gDicStableTime.Add Key:=660, item:=32
    gDicStableTime.Add Key:=720, item:=33
    gDicStableTime.Add Key:=780, item:=34
    gDicStableTime.Add Key:=840, item:=35
    gDicStableTime.Add Key:=900, item:=36
    gDicStableTime.Add Key:=960, item:=37
    gDicStableTime.Add Key:=1020, item:=38
    gDicStableTime.Add Key:=1080, item:=39
    gDicStableTime.Add Key:=1140, item:=40
    gDicStableTime.Add Key:=1200, item:=41
    gDicStableTime.Add Key:=1260, item:=42
    gDicStableTime.Add Key:=1320, item:=43
    gDicStableTime.Add Key:=1380, item:=44
    gDicStableTime.Add Key:=1440, item:=45
    gDicStableTime.Add Key:=1500, item:=46
    
    gDicMyTime.Add Key:=17, item:=60
    gDicMyTime.Add Key:=18, item:=75
    gDicMyTime.Add Key:=19, item:=90
    gDicMyTime.Add Key:=20, item:=105
    gDicMyTime.Add Key:=21, item:=120
    gDicMyTime.Add Key:=22, item:=140
    gDicMyTime.Add Key:=23, item:=160
    gDicMyTime.Add Key:=24, item:=180
    gDicMyTime.Add Key:=25, item:=240
    gDicMyTime.Add Key:=26, item:=300
    gDicMyTime.Add Key:=27, item:=360
    gDicMyTime.Add Key:=28, item:=420
    gDicMyTime.Add Key:=29, item:=480
    gDicMyTime.Add Key:=30, item:=540
    gDicMyTime.Add Key:=31, item:=600
    gDicMyTime.Add Key:=32, item:=660
    gDicMyTime.Add Key:=33, item:=720
    gDicMyTime.Add Key:=34, item:=780
    gDicMyTime.Add Key:=35, item:=840
    gDicMyTime.Add Key:=36, item:=900
    gDicMyTime.Add Key:=37, item:=960
    gDicMyTime.Add Key:=38, item:=1020
    gDicMyTime.Add Key:=39, item:=1080
    gDicMyTime.Add Key:=40, item:=1140
    gDicMyTime.Add Key:=41, item:=1200
    gDicMyTime.Add Key:=42, item:=1260
    gDicMyTime.Add Key:=43, item:=1320
    gDicMyTime.Add Key:=44, item:=1380
    gDicMyTime.Add Key:=45, item:=1440
    gDicMyTime.Add Key:=46, item:=1500
End Sub

'10-77 : 2880 (68) - longterm pumping test
'78-101: recover (24) - recover test

Sub set_daydifference()
    Dim n_passed_time() As Integer
    Dim i           As Integer
    Dim day1, day2  As Integer
    
    ReDim n_passed_time(1 To 92)
    
    For i = 1 To 92
        n_passed_time(i) = Cells(i + 9, "D").Value
        If (i > 68) Then
            n_passed_time(i) = Cells(i + 9, "D").Value + 2880
        End If
    Next i
    
    For i = 1 To 92
        Cells(i + 9, "h").Value = Range("c10").Value + n_passed_time(i) / 1440
    Next i
    
    Range("H10:H101").Select
    Selection.NumberFormatLocal = "yyyy""년"" m""월"" d""일"";@"
    Range("A1").Select
    Application.CutCopyMode = False
    
    Application.ScreenUpdating = False
    day1 = Day(Cells(10, "h").Value)
    
    For i = 2 To 92
        day2 = Day(Cells(i + 9, "h").Value)
        If (day2 = day1) Then
            Cells(i + 9, "h").Value = ""
        End If
        day1 = day2
    Next i
    
    Range("h77").Value = "양수종료"
    Range("h78").Value = "회복수위측정"
    Application.ScreenUpdating = True
End Sub

Function find_stable_time() As Integer
    Dim i           As Integer
    
    For i = 10 To 50
        If Range("AC" & CStr(i)).Value = Range("AC" & CStr(i + 1)) Then
            'MsgBox "found " & "AB" & CStr(i) & " time : " & Range("Z" & CStr(i)).Value
            find_stable_time = i
            Exit For
        End If
    Next i
End Function

Function initialize_myTime() As Integer
    initialize_myTime = gDicStableTime(shW_aSkinFactor.Range("g16").Value)
End Function

Sub TimeSetting()
    Dim stable_time, h1, h2, my_random_time As Integer
    Dim myRange     As String
    
    stable_time = find_stable_time()
    
    If MY_TIME = 0 Then
        MY_TIME = initialize_myTime
        my_random_time = MY_TIME
    Else
        my_random_time = MY_TIME
    End If
    
    If stable_time < my_random_time Then
        h1 = stable_time
        h2 = my_random_time
        Range("ac" & CStr(h1)).Select
        myRange = "AC" & CStr(h1) & ":AC" & CStr(h2)
        
    ElseIf stable_time > my_random_time Then
        h1 = my_random_time
        h2 = stable_time
        Range("ac" & CStr(h2 + 1)).Select
        myRange = "AC" & CStr(h1 + 1) & ":AC" & CStr(h2 + 1)
    Else
        Exit Sub
    End If
    
    Selection.AutoFill Destination:=Range(myRange), Type:=xlFillDefault
    setSkinTime (MY_TIME)
    
    Range("A27").Select
End Sub

Sub setSkinTime(i As Integer)
    Application.ScreenUpdating = False
    
    shW_aSkinFactor.Activate
    Range("G16").Value = gDicMyTime(i)
    shW_LongTEST.Activate
    
    Application.ScreenUpdating = True
End Sub

Sub cellRED(ByVal strcell As String)
    Range(strcell).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13209
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
End Sub

Sub cellBLACK(ByVal strcell As String)
    Range(strcell).Select
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.499984740745262
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
End Sub

Sub resetValue()
    Range("p3").ClearContents
    Range("t1").Value = 0.1
    Range("l6").Value = 0.2
    
    '2023/4/3
    If Not IsNumeric(Range("GoalSeekTarget").Value) Then
        Range("GoalSeekTarget").Value = 0
    End If
    
    Range("o3:o14").ClearContents
    
    ' ActiveSheet.OptionButton1.Value = True
End Sub

Function isPositive(ByVal data As Double) As Double
    If data < 0 Then
        isPositive = False
    Else
        isPositive = True
    End If
End Function

Function CellReverse(ByVal data As Double) As Double
    If data < 0 Then
        CellReverse = Abs(data)
    Else
        CellReverse = -data
    End If
End Function

Sub findAnswer_LongTest()
    ' 2023/4/3
    Dim GoalSeekTarget As Double
    
    
    ' 2023/4/3
    If Not IsNumeric(Range("GoalSeekTarget").Value) Then
        GoalSeekTarget = 0
    Else
        GoalSeekTarget = Range("GoalSeekTarget").Value
    End If
    
    If (Range("p3").Value <> 0) Then
        Exit Sub
    End If
    
    
  
    Range("L10").GoalSeek goal:=GoalSeekTarget, ChangingCell:=Range("T1")
    Range("p3").Value = CellReverse(Range("k10").Value)
    
    If Range("l8").Value < 0 Then
        cellRED ("l8")
    Else
        cellBLACK ("l8")
    End If
    
    shW_aSkinFactor.Range("d5").Value = Round(Range("t1").Value, 4)
End Sub

Sub check_LongTest()
    Dim igoal, k0, k1 As Double
    
    k1 = Range("l8").Value
    k0 = Range("l6").Value
    
    If k0 = k1 Then Exit Sub
    If k1 > 0 Then Exit Sub
    
    If k0 <> "" Then
        igoal = k0
    Else
        igoal = 0.3
    End If
    
    Range("l8").GoalSeek goal:=igoal, ChangingCell:=Range("o3")
    
    If Range("l8").Value < 0 Then
        cellRED ("l8")
    Else
        cellBLACK ("l8")
    End If
    
    Call drawdown_level_check
End Sub

' 2023/4/4
Sub drawdown_level_check()

    Dim check_cell(1 To 5) As String
    Dim i As Integer
    
    check_cell(1) = "G10"
    check_cell(2) = "K11"
    check_cell(3) = "K12"
    check_cell(4) = "K13"
    check_cell(5) = "K14"
    
    For i = 1 To 4
        If Range(check_cell(i)).Value < Range(check_cell(i + 1)).Value Then
            MsgBox "Water Level is Trouble ....", vbOKOnly
            Exit Sub
        End If
    Next i

End Sub



Sub findAnswer_StepTest()
    Range("Q4:Q13").ClearContents
    Range("T4").Value = 0.1
    Range("G12").GoalSeek goal:=1#, ChangingCell:=Range("T4")
    
    If Range("J11").Value < 0 Then
        cellRED ("J11")
    Else
        cellBLACK ("J11")
    End If
End Sub

Sub check_StepTest()
    Dim igoal, nj   As Double
    
    igoal = 0.12
    
    Do While (Range("J11").Value < 0 Or Range("j11").Value >= 50)
        Range("J11").GoalSeek goal:=igoal, ChangingCell:=Range("Q4")
        igoal = igoal + 0.1
    Loop
    
    If Range("J11").Value < 0 Then
        cellRED ("J11")
    Else
        cellBLACK ("J11")
    End If
End Sub
Sub show_gachae()
    shYield.Visible = True
    shYield.Select
End Sub

Sub hide_gachae()
    shYield.Visible = False
    shW_aSkinFactor.Select
End Sub

Option Explicit

Function Contains(objCollection As Object, strName As String) As Boolean
    Dim o           As Object
    
    On Error Resume Next
    Set o = objCollection(strName)
    Contains = (Err.Number = 0)
    Err.Clear
End Function

Function ConvertToLongInteger(ByVal stValue As String) As Long
    On Error GoTo ConversionFailureHandler
    ConvertToLongInteger = CLng(stValue)        'TRY to convert to an Integer value
    Exit Function        'If we reach this point, then we succeeded so exit
    
ConversionFailureHandler:
    'IF we've reached this point, then we did not succeed in conversion
    'If the error is type-mismatch, clear the error and return numeric 0 from the function
    'Otherwise, disable the error handler, and re-run the code to allow the system to
    'display the error
    If Err.Number = 13 Then        'error # 13 is Type mismatch
    Err.Clear
    ConvertToLongInteger = 0
    Exit Function
Else
    On Error GoTo 0
    Resume
End If

End Function

Function IsSheetsHasA(name As String)
    Dim sheet As Worksheet
    Dim result As Integer
    
    ' Loop through all sheets in the workbook
    For Each sheet In ThisWorkbook.Worksheets
        result = StrComp(sheet.name, name, vbTextCompare)
        If result = 0 Then
            IsSheetsHasA = True
            Exit Function
        End If
    Next sheet
    
    IsSheetsHasA = False
End Function



Function sheets_count() As Long
    Dim i, nSheetsCount, nWell  As Integer
    Dim strSheetsName(50) As String
    
    nSheetsCount = ThisWorkbook.Sheets.Count
    nWell = 0
    
    For i = 1 To nSheetsCount
        strSheetsName(i) = ThisWorkbook.Sheets(i).name
        'MsgBox (strSheetsName(i))
        If (ConvertToLongInteger(strSheetsName(i)) <> 0) Then
            nWell = nWell + 1
        End If
    Next
    
    'MsgBox (CStr(nWell))
    sheets_count = nWell
End Function

' https://www.google.com/search?q=excel+vba+how+to+get+number+from+string&oq=excel+vba+how+to+get+number+from+string&aqs=chrome..69i57&sourceid=chrome&ie=UTF-8
' https://stackoverflow.com/questions/28771802/extract-number-from-string-in-vba

Function GetNumbers(str As String) As Long
    Dim regex       As Object
    Dim matches     As Variant
    
    Set regex = CreateObject("vbscript.regexp")
    
    regex.Pattern = "(\d+)"
    regex.Global = True
    
    Set matches = regex.Execute(str)
    GetNumbers = matches(0)
End Function

Function CleanString(strIn As String) As String
    Dim objRegex
    Set objRegex = CreateObject("vbscript.regexp")
    With objRegex
        .Global = True
        .Pattern = "[^\d]+"
        CleanString = .Replace(strIn, vbNullString)
    End With
End Function

'https://stackoverflow.com/questions/40365573/excel-vba-extract-numeric-value-in-string
'Requires a reference to Microsoft VBScript Regular Expressions X.X

Public Function ExtractNumber(inValue As String) As Double
    With New regExp
        .Pattern = "(\d{1,3},?)+(\.\d{2})?"
        .Global = True
        If .test(inValue) Then
            ExtractNumber = CDbl(.Execute(inValue)(0))
        End If
    End With
End Function

'https://stackoverflow.com/questions/50994883/how-to-extract-numbers-from-a-text-string-in-vba

Sub ExtractNumbers()
    Dim str         As String, regex As regExp, matches As MatchCollection, match As match
    
    str = "ID CSys ID Set ID Set Value Set Title 7026..Plate Top MajorPrn Stress 7027..Plate Top MinorPrn Stress 7033..Plate Top VonMises Stress"
    
    Set regex = New regExp
    regex.Pattern = "\d+"        '~~~> Look for variable length numbers only
    regex.Global = True
    
    If (regex.test(str) = True) Then
        Set matches = regex.Execute(str)        '~~~> Execute search
        
        For Each match In matches
            Debug.Print match.Value        '~~~> Prints: 7026, 7027, 7033
        Next
    End If
End Sub
Option Explicit

'0 : skin factor
'1 : Re1
'2 : Re2
'3 : Re3

Public Enum ER_VALUE
    erRE0 = 0
    erRE1 = 1
    erRE2 = 2
    erRE3 = 3
End Enum

Function GetEffectiveRadius() As Integer
    Dim er, r       As String
    
    er = Range("EffectiveRadius").Value
    'MsgBox er
    r = Mid(er, 5, 1)
    
    If r = "F" Then
        GetEffectiveRadius = 0
    Else
        GetEffectiveRadius = Val(r)
    End If
End Function

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Function GetNowLast(inputDate As Date) As Date

    Dim dYear, dMonth, getDate As Date

    dYear = Year(inputDate)
    dMonth = Month(inputDate)

    getDate = DateSerial(dYear, dMonth + 1, 0)

    GetNowLast = getDate

End Function

Private Sub ComboBoxFix(ByVal SIGN As Boolean)

    Dim contr As Control
    
    If SIGN Then
        For Each contr In UserFormTS.Controls
            If TypeName(contr) = "ComboBox" Then
                contr.Style = fmStyleDropDownList
            End If
        Next
    Else
        For Each contr In UserFormTS.Controls
            If TypeName(contr) = "ComboBox" Then
                contr.Style = fmStyleDropDownCombo
            End If
        Next
    End If

End Sub

Private Function whichSection(n As Integer) As Integer

    whichSection = Round((n / 10), 0) * 10

End Function

Private Sub ComboBoxYear_Initialize()
    Dim nYear, nMonth, nDay As Integer
    Dim nHour, nMin As Integer
    
    Dim i, j As Integer
    Dim lastDay As Integer
    
    Dim sheetDate, currDate As Date
    Dim isThisYear As Boolean
    
    sheetDate = Range("c10").Value
    currDate = Now()
    
    If ((Year(currDate) - Year(sheetDate)) = 0) Then
    
        isThisYear = True
        
        nYear = Year(sheetDate)
        nMonth = Month(sheetDate)
        nDay = Day(sheetDate)
        
        nHour = Hour(sheetDate)
        nMin = Minute(sheetDate)
        
    Else
        
        isThisYear = False
        
        nYear = Year(currDate)
        nMonth = Month(currDate)
        nDay = Day(currDate)
        
        nHour = Hour(currDate)
        nMin = Minute(currDate)
            
    End If
    
    
    lastDay = Day(GetNowLast(IIf(isThisYear, sheetDate, currDate)))
    Debug.Print lastDay
    
    For i = nYear - 10 To nYear
        ComboBoxYear.AddItem (i)
    Next i
    
    For i = 1 To 12
        ComboBoxMonth.AddItem (i)
    Next i
    
    For i = 1 To lastDay
        ComboBoxDay.AddItem (i)
    Next i
    
            
    For i = 1 To 12
        ComboBoxHour.AddItem (i)
    Next i
    
    
    
    For i = 0 To 60 Step 10
        ComboBoxMinute.AddItem (i)
    Next i
    
    
    
    ComboBoxYear.Value = nYear
    ComboBoxMonth.Value = nMonth
    ComboBoxDay.Value = nDay
    
    ComboBoxHour.Value = IIf(nHour > 12, nHour - 12, nHour)
    ComboBoxMinute.Value = whichSection(IIf(isThisYear, Minute(sheetDate), Minute(currDate)))
    
   
    If nHour > 12 Then
        OptionButtonPM.Value = True
    Else
        OptionButtonAM.Value = True
    End If
    
    Debug.Print nYear

End Sub

Sub ComboboxDay_ChangeItem(nYear As Integer, nMonth As Integer)
    Dim lasday, i As Integer
    
    lasday = Day(GetNowLast(DateSerial(nYear, nMonth, 1)))
    ComboBoxDay.Clear
    
    For i = 1 To lasday
        ComboBoxDay.AddItem (i)
    Next i
    
    ComboBoxDay.Value = 1

End Sub

Private Sub ComboBoxHour_Change()
    ComboBoxMinute.Value = 0
End Sub

Private Sub ComboBoxMonth_Change()
    '2019-11-26 change
    On Error GoTo Errcheck
    Call ComboboxDay_ChangeItem(ComboBoxYear.Value, ComboBoxMonth.Value)
Errcheck:
        
End Sub

Private Sub EnterButton_Click()
    Dim nYear, nMonth, nDay As Integer
    Dim nHour, nMinute As Integer
    
    Dim nDate, ntime As Date
    
    
    On Error GoTo Errcheck
    nYear = ComboBoxYear.Value
    nMonth = ComboBoxMonth.Value
    nDay = ComboBoxDay.Value
        
    nHour = ComboBoxHour.Value
    nMinute = ComboBoxMinute.Value
            
            
    nHour = nHour + IIf(OptionButtonPM.Value, 12, 0)
            
    nDate = DateSerial(nYear, nMonth, nDay)
    ntime = TimeSerial(nHour, nMinute, 0)
            
    nDate = nDate + ntime
         
    shW_LongTEST.Range("c10").Value = nDate
         
Errcheck:
     
    Unload Me
     
End Sub

Private Sub UserForm_Initialize()
    Call ComboBoxYear_Initialize
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
End Sub




Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Function GetNowLast(inputDate As Date) As Date

    Dim dYear, dMonth, getDate As Date

    dYear = Year(inputDate)
    dMonth = Month(inputDate)

    getDate = DateSerial(dYear, dMonth + 1, 0)

    GetNowLast = getDate

End Function

Private Sub ComboBoxFix(ByVal SIGN As Boolean)

    Dim contr As Control
    
    If SIGN Then
        For Each contr In UserFormTS.Controls
            If TypeName(contr) = "ComboBox" Then
                contr.Style = fmStyleDropDownList
            End If
        Next
    Else
        For Each contr In UserFormTS.Controls
            If TypeName(contr) = "ComboBox" Then
                contr.Style = fmStyleDropDownCombo
            End If
        Next
    End If

End Sub

Private Function whichSection(n As Integer) As Integer

    whichSection = Round((n / 10), 0) * 10

End Function

Private Sub ComboBoxYear_Initialize()
    Dim nYear, nMonth, nDay As Integer
    Dim nHour, nMin As Integer
    
    Dim i, j As Integer
    Dim lastDay As Integer
    
    Dim sheetDate, currDate As Date
    Dim isThisYear As Boolean
    
    sheetDate = Range("c6").Value
    'MsgBox (sheetDate)
    currDate = Now()
    
    If ((Year(currDate) - Year(sheetDate)) = 0) Then
    
        isThisYear = True
        
        nYear = Year(sheetDate)
        nMonth = Month(sheetDate)
        nDay = Day(sheetDate)
        
        nHour = Hour(sheetDate)
        nMin = Minute(sheetDate)
        
    Else
        
        isThisYear = False
        
        nYear = Year(currDate)
        nMonth = Month(currDate)
        nDay = Day(currDate)
        
        nHour = Hour(currDate)
        nMin = Minute(currDate)
            
    End If
    
    
    lastDay = Day(GetNowLast(IIf(isThisYear, sheetDate, currDate)))
    ' Debug.Print lastDay
    
    For i = nYear - 10 To nYear
        ComboBoxYear.AddItem (i)
    Next i
    
    For i = 1 To 12
        ComboBoxMonth.AddItem (i)
    Next i
    
    For i = 1 To lastDay
        ComboBoxDay.AddItem (i)
    Next i
    
            
    For i = 1 To 12
        ComboBoxHour.AddItem (i)
    Next i
    
    
    
    For i = 0 To 60 Step 10
        ComboBoxMinute.AddItem (i)
    Next i
    
    
    
    ComboBoxYear.Value = nYear
    ComboBoxMonth.Value = nMonth
    ComboBoxDay.Value = nDay
    
    ComboBoxHour.Value = IIf(nHour > 12, nHour - 12, nHour)
    ComboBoxMinute.Value = whichSection(IIf(isThisYear, Minute(sheetDate), Minute(currDate)))
    
   
    If nHour > 12 Then
        OptionButtonPM.Value = True
    Else
        OptionButtonAM.Value = True
    End If
    
    ' Debug.Print nYear

End Sub

Sub ComboboxDay_ChangeItem(nYear As Integer, nMonth As Integer)
    Dim lasday, i As Integer
    
    lasday = Day(GetNowLast(DateSerial(nYear, nMonth, 1)))
    ComboBoxDay.Clear
    
    For i = 1 To lasday
        ComboBoxDay.AddItem (i)
    Next i
    
    ComboBoxDay.Value = 1

End Sub

Private Sub ComboBoxHour_Change()
    ComboBoxMinute.Value = 0
End Sub

Private Sub ComboBoxMonth_Change()
    '2019-11-26 change
    On Error GoTo Errcheck
    Call ComboboxDay_ChangeItem(ComboBoxYear.Value, ComboBoxMonth.Value)
Errcheck:
        
End Sub

Private Sub EnterButton_Click()
    Dim nYear, nMonth, nDay As Integer
    Dim nHour, nMinute As Integer
    
    Dim nDate, ntime As Date
    
    
    On Error GoTo Errcheck
    nYear = ComboBoxYear.Value
    nMonth = ComboBoxMonth.Value
    nDay = ComboBoxDay.Value
        
    nHour = ComboBoxHour.Value
    nMinute = ComboBoxMinute.Value
            
            
    nHour = nHour + IIf(OptionButtonPM.Value, 12, 0)
            
    nDate = DateSerial(nYear, nMonth, nDay)
    ntime = TimeSerial(nHour, nMinute, 0)
            
    nDate = nDate + ntime
         
    Range("c6").Value = nDate
         
Errcheck:
     
    Unload Me
     
End Sub

Private Sub UserForm_Initialize()
    Call ComboBoxYear_Initialize
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
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


Sub set_CB_ALL()

    Call set_CB1
    MsgBox "SetCB1 Complete and Next setCB2 .... ", vbOKOnly
    Call set_CB2
    
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


Option Explicit


Private Sub OptionButton1_Click()
    TextBox1.Text = "60 - 17"
    MY_TIME = 17
End Sub

Private Sub OptionButton2_Click()
    TextBox1.Text = "75 - 18"
    MY_TIME = 18
End Sub

Private Sub OptionButton3_Click()
    TextBox1.Text = "90 - 19"
    MY_TIME = 19
End Sub

Private Sub OptionButton4_Click()
    TextBox1.Text = "105 - 20"
    MY_TIME = 20
End Sub

Private Sub OptionButton5_Click()
    TextBox1.Text = "120 - 21"
    MY_TIME = 21
End Sub

Private Sub OptionButton6_Click()
    TextBox1.Text = "140 - 22"
    MY_TIME = 22
End Sub

Private Sub OptionButton7_Click()
    TextBox1.Text = "160 - 23"
    MY_TIME = 23
End Sub

Private Sub OptionButton8_Click()
    TextBox1.Text = "180 - 24"
    MY_TIME = 24
End Sub

Private Sub OptionButton9_Click()
    TextBox1.Text = "240 - 25"
    MY_TIME = 25
End Sub


Private Sub OptionButton10_Click()
    TextBox1.Text = "300 - 26"
    MY_TIME = 26
End Sub

Private Sub OptionButton11_Click()
    TextBox1.Text = "360 - 27"
    MY_TIME = 27
End Sub

Private Sub OptionButton12_Click()
    TextBox1.Text = "420 - 28"
    MY_TIME = 28
End Sub

Private Sub OptionButton13_Click()
    TextBox1.Text = "480 - 29"
    MY_TIME = 29
End Sub

Private Sub OptionButton14_Click()
    TextBox1.Text = "540 - 30"
    MY_TIME = 30
End Sub

Private Sub OptionButton15_Click()
    TextBox1.Text = "600 - 31"
    MY_TIME = 31
End Sub

Private Sub OptionButton16_Click()
    TextBox1.Text = "660 - 32"
    MY_TIME = 32
End Sub

Private Sub OptionButton17_Click()
    TextBox1.Text = "720 - 33"
    MY_TIME = 33
End Sub

Private Sub OptionButton18_Click()
    TextBox1.Text = "780 - 34"
    MY_TIME = 34
End Sub

Private Sub OptionButton19_Click()
    TextBox1.Text = "840 - 35"
    MY_TIME = 35
End Sub

Private Sub OptionButton20_Click()
    TextBox1.Text = "900 - 36"
    MY_TIME = 36
End Sub

Private Sub OptionButton21_Click()
    TextBox1.Text = "960 - 37"
    MY_TIME = 37
End Sub

Private Sub OptionButton22_Click()
    TextBox1.Text = "1020 - 38"
    MY_TIME = 38
End Sub

Private Sub OptionButton23_Click()
    TextBox1.Text = "1080 - 39"
    MY_TIME = 39
End Sub

Private Sub OptionButton24_Click()
    TextBox1.Text = "1140 - 40"
    MY_TIME = 40
End Sub

Private Sub OptionButton25_Click()
    TextBox1.Text = "1200 - 41"
    MY_TIME = 41
End Sub

Private Sub OptionButton26_Click()
    TextBox1.Text = "1260 - 42"
    MY_TIME = 42
End Sub

Private Sub OptionButton27_Click()
    TextBox1.Text = "1320 - 43"
    MY_TIME = 43
End Sub

Private Sub OptionButton28_Click()
    TextBox1.Text = "1380 - 43"
    MY_TIME = 44
End Sub

Private Sub OptionButton29_Click()
    TextBox1.Text = "1440 - 45"
    MY_TIME = 45
End Sub


Private Sub OptionButton30_Click()
    TextBox1.Text = "1600 - 46"
    MY_TIME = 46
End Sub


Private Sub CancelButton_Click()
    Dim i As Integer
    
    TextBox1.Text = shW_LongTEST.ComboBox1.Value
    i = gDicStableTime(CInt(shW_LongTEST.ComboBox1.Value)) ' - 16
    Call SetOptionButtonClick(i)
    Unload Me
End Sub


Private Sub EnterButton_Click()
    On Error GoTo Errcheck
        shW_LongTEST.ComboBox1.Value = gDicMyTime(MY_TIME)
         
Errcheck:
    Unload Me
End Sub


Private Sub SetOptionButtonClick(i As Integer)

 Select Case i
        Case 17:
           Call OptionButton1_Click
        Case 18:
           Call OptionButton2_Click
        Case 19:
           Call OptionButton3_Click
        Case 20:
           Call OptionButton4_Click
        Case 21:
           Call OptionButton5_Click
        Case 22:
           Call OptionButton6_Click
        Case 23:
           Call OptionButton7_Click
        Case 24:
           Call OptionButton8_Click
        Case 25:
           Call OptionButton9_Click
        Case 26:
           Call OptionButton10_Click
        Case 27:
           Call OptionButton11_Click
        Case 28:
           Call OptionButton12_Click
        Case 29:
           Call OptionButton13_Click
        Case 30:
           Call OptionButton14_Click
        Case 31:
           Call OptionButton15_Click
        Case 32:
           Call OptionButton16_Click
        Case 33:
           Call OptionButton17_Click
        Case 34:
           Call OptionButton18_Click
        Case 35:
           Call OptionButton19_Click
        Case 36:
           Call OptionButton20_Click
        Case 37:
           Call OptionButton21_Click
        Case 38:
           Call OptionButton22_Click
        Case 39:
           Call OptionButton23_Click
        Case 40:
           Call OptionButton24_Click
        Case 41:
           Call OptionButton25_Click
        Case 42:
            Call OptionButton26_Click
        Case 43:
           Call OptionButton27_Click
        Case 44:
           Call OptionButton28_Click
        Case 45:
           Call OptionButton29_Click
        Case 46:
           Call OptionButton30_Click
        
        Case Else:
             Call OptionButton17_Click
    End Select

End Sub

Private Sub SetOptionButton(i As Integer)

 Select Case i
        Case 17:
            OptionButton1.Value = True
        Case 18:
            OptionButton2.Value = True
        Case 19:
            OptionButton3.Value = True
        Case 20:
            OptionButton4.Value = True
        Case 21:
           OptionButton5.Value = True
        Case 22:
           OptionButton6.Value = True
        Case 23:
           OptionButton7.Value = True
        Case 24:
            OptionButton8.Value = True
        Case 25:
            OptionButton9.Value = True
        Case 26:
            OptionButton10.Value = True
        Case 27:
            OptionButton11.Value = True
        Case 28:
            OptionButton12.Value = True
        Case 29:
           OptionButton13.Value = True
        Case 30:
           OptionButton14.Value = True
        Case 31:
           OptionButton15.Value = True
        Case 32:
            OptionButton16.Value = True
        Case 33:
            OptionButton17.Value = True
        Case 34:
            OptionButton18.Value = True
        Case 35:
            OptionButton19.Value = True
        Case 36:
            OptionButton20.Value = True
        Case 37:
           OptionButton21.Value = True
        Case 38:
           OptionButton22.Value = True
        Case 39:
           OptionButton23.Value = True
        Case 40:
            OptionButton24.Value = True
        Case 41:
            OptionButton25.Value = True
        Case 42:
            OptionButton26.Value = True
        Case 43:
            OptionButton27.Value = True
        Case 44:
            OptionButton28.Value = True
        Case 45:
           OptionButton29.Value = True
        Case 46:
           OptionButton30.Value = True
        
        Case Else:
             OptionButton17.Value = True
    End Select

End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer
    
    Call initDictionary
    TextBox1.Text = shW_LongTEST.ComboBox1.Value
    i = gDicStableTime(CInt(shW_LongTEST.ComboBox1.Value)) ' - 16
    Call SetOptionButton(i)
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
      
End Sub






Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Function GetNowLast(inputDate As Date) As Date

    Dim dYear, dMonth, getDate As Date

    dYear = Year(inputDate)
    dMonth = Month(inputDate)

    getDate = DateSerial(dYear, dMonth + 1, 0)

    GetNowLast = getDate

End Function

Private Sub ComboBoxFix(ByVal SIGN As Boolean)

    Dim contr As Control
    
    If SIGN Then
        For Each contr In UserFormTS.Controls
            If TypeName(contr) = "ComboBox" Then
                contr.Style = fmStyleDropDownList
            End If
        Next
    Else
        For Each contr In UserFormTS.Controls
            If TypeName(contr) = "ComboBox" Then
                contr.Style = fmStyleDropDownCombo
            End If
        Next
    End If

End Sub

Private Function whichSection(n As Integer) As Integer

    whichSection = Round((n / 10), 0) * 10

End Function

Private Sub ComboBoxYear_Initialize()
    Dim nYear, nMonth, nDay As Integer
    Dim nHour, nMin As Integer
    
    Dim i, j As Integer
    Dim lastDay As Integer
    
    Dim sheetDate, currDate As Date
    Dim isThisYear As Boolean
    
    sheetDate = Range("c12").Value
    'MsgBox (sheetDate)
    currDate = Now()
    
    If ((Year(currDate) - Year(sheetDate)) = 0) Then
    
        isThisYear = True
        
        nYear = Year(sheetDate)
        nMonth = Month(sheetDate)
        nDay = Day(sheetDate)
        
        nHour = Hour(sheetDate)
        nMin = Minute(sheetDate)
        
    Else
        
        isThisYear = False
        
        nYear = Year(currDate)
        nMonth = Month(currDate)
        nDay = Day(currDate)
        
        nHour = Hour(currDate)
        nMin = Minute(currDate)
            
    End If
    
    
    lastDay = Day(GetNowLast(IIf(isThisYear, sheetDate, currDate)))
    ' Debug.Print lastDay
    
    For i = nYear - 10 To nYear
        ComboBoxYear.AddItem (i)
    Next i
    
    For i = 1 To 12
        ComboBoxMonth.AddItem (i)
    Next i
    
    For i = 1 To lastDay
        ComboBoxDay.AddItem (i)
    Next i
    
            
    For i = 1 To 12
        ComboBoxHour.AddItem (i)
    Next i
    
    
    
    For i = 0 To 60 Step 10
        ComboBoxMinute.AddItem (i)
    Next i
    
    
    
    ComboBoxYear.Value = nYear
    ComboBoxMonth.Value = nMonth
    ComboBoxDay.Value = nDay
    
    ComboBoxHour.Value = IIf(nHour > 12, nHour - 12, nHour)
    ComboBoxMinute.Value = whichSection(IIf(isThisYear, Minute(sheetDate), Minute(currDate)))
    
   
    If nHour > 12 Then
        OptionButtonPM.Value = True
    Else
        OptionButtonAM.Value = True
    End If
    
    ' Debug.Print nYear

End Sub

Sub ComboboxDay_ChangeItem(nYear As Integer, nMonth As Integer)
    Dim lasday, i As Integer
    
    lasday = Day(GetNowLast(DateSerial(nYear, nMonth, 1)))
    ComboBoxDay.Clear
    
    For i = 1 To lasday
        ComboBoxDay.AddItem (i)
    Next i
    
    ComboBoxDay.Value = 1

End Sub

Private Sub ComboBoxHour_Change()
    ComboBoxMinute.Value = 0
End Sub

Private Sub ComboBoxMonth_Change()
    '2019-11-26 change
    On Error GoTo Errcheck
    Call ComboboxDay_ChangeItem(ComboBoxYear.Value, ComboBoxMonth.Value)
Errcheck:
        
End Sub

Private Sub EnterButton_Click()
    Dim nYear, nMonth, nDay As Integer
    Dim nHour, nMinute As Integer
    
    Dim nDate, ntime As Date
    
    
    On Error GoTo Errcheck
    nYear = ComboBoxYear.Value
    nMonth = ComboBoxMonth.Value
    nDay = ComboBoxDay.Value
        
    nHour = ComboBoxHour.Value
    nMinute = ComboBoxMinute.Value
            
            
    nHour = nHour + IIf(OptionButtonPM.Value, 12, 0)
            
    nDate = DateSerial(nYear, nMonth, nDay)
    ntime = TimeSerial(nHour, nMinute, 0)
            
    nDate = nDate + ntime
         
    Range("c12").Value = nDate
         
Errcheck:
     
    Unload Me
     
End Sub

Private Sub UserForm_Initialize()
    Call ComboBoxYear_Initialize
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub



Sub GitSave()
    
    DeleteAndMake
    ExportModules
    PrintAllCode
    PrintAllContainers
    
End Sub

Sub DeleteAndMake()
        
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim parentFolder As String: parentFolder = ThisWorkbook.Path & "\VBA"
    Dim childA As String: childA = parentFolder & "\VBA-Code_Together"
    Dim childB As String: childB = parentFolder & "\VBA-Code_By_Modules"
        
    On Error Resume Next
    fso.DeleteFolder parentFolder
    On Error GoTo 0
    
    MkDir parentFolder
    MkDir childA
    MkDir childB
    
End Sub

Sub PrintAllCode()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    Dim fName As String
    
    Dim pathToExport As String
    pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_Together\"
    If Dir(pathToExport) <> "" Then Kill pathToExport & "*.*"
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        
        
        If item.CodeModule.CountOfLines <> 0 Then
            lineToPrint = item.CodeModule.Lines(1, item.CodeModule.CountOfLines)
        Else
            lineToPrint = "'This Module is Empty "
        End If
        
        
        fName = item.CodeModule.name
        Debug.Print lineToPrint
        SaveTextToFile lineToPrint, pathToExport & fName & ".bas"
        
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    
    SaveTextToFile textToPrint, pathToExport & "all_code.vb"
    
End Sub

Sub PrintAllContainers()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.name
        Debug.Print lineToPrint
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_Together\"
    SaveTextToFile textToPrint, pathToExport & "all_modules.vb"
    
End Sub

Sub ExportModules()
       
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_By_Modules\"
    
    If Dir(pathToExport) <> "" Then
        Kill pathToExport & "*.*"
    End If
     
    Dim wkb As Workbook: Set wkb = Excel.Workbooks(ThisWorkbook.name)
    
    Dim unitsCount As Long
    Dim filePath As String
    Dim component As VBIDE.VBComponent
    Dim tryExport As Boolean

    For Each component In wkb.VBProject.VBComponents
        tryExport = True
        filePath = component.name
       
        Select Case component.Type
            Case vbext_ct_ClassModule
                filePath = filePath & ".cls"
            Case vbext_ct_MSForm
                filePath = filePath & ".frm"
            Case vbext_ct_StdModule
                filePath = filePath & ".bas"
            Case vbext_ct_Document
                tryExport = False
        End Select
        
        If tryExport Then
            Debug.Print unitsCount & " exporting " & filePath
            component.Export pathToExport & "\" & filePath
        End If
    Next

    Debug.Print "Exported at " & pathToExport
    
End Sub

Sub SaveTextToFile(dataToPrint As String, pathToExport As String)
    
    Dim fileSystem As Object
    Dim textObject As Object
    Dim fileName As String
    Dim newFile  As String
    Dim shellPath  As String
    
    If Dir(ThisWorkbook.Path & newFile, vbDirectory) = vbNullString Then MkDir ThisWorkbook.Path & newFile
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set textObject = fileSystem.CreateTextFile(pathToExport, True)
    
    textObject.WriteLine dataToPrint
    textObject.Close
        
    On Error GoTo 0
    Exit Sub

CreateLogFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateLogFile of Sub mod_TDD_Export"

End Sub


Sub 매크로1()
'
' 매크로1 매크로
'

'
    ActiveSheet.VPageBreaks(1).DragOff Direction:=xlToRight, RegionIndex:=1
End Sub
