VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
    Call adjustChartGraph
End Sub

Private Sub CommandButton7_Click()
    Call Make2880Document
    Call make1440sheet
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








