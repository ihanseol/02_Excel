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


Private Sub CommandButton1_Click()
    step_pumping_test
    vertical_copy
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
    Call make_long_document
End Sub

Private Sub CommandButton6_Click()
    Call adjustChartGraph
End Sub



Private Sub Worksheet_Activate()
  
  Dim gong As Integer
  
  WB_NAME = ThisWorkbook.name
  gong = Right(Range("J48").Value, 1)
  
  Call SetChartTitleText(gong)

End Sub

'2019/11/24 - set gong bun change
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim KeyCell As Range
    Dim gong As Integer
    
    Dim g1, g2 As String
    
    Set KeyCell = Range("J48")
    
    If Not Application.Intersect(KeyCell, Range(Target.Address)) Is Nothing Then
        
        gong = Right(KeyCell.Value, 1)
        Range("i54").Value = "W-" & gong
        
        Call SetChartTitleText(gong)
    End If
    
End Sub


Private Sub SetChartTitleText(ByVal i As Integer)
    
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







