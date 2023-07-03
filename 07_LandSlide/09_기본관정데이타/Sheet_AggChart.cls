VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_AggChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit


Private Sub CommandButton1_Click()
    Sheets("AggChart").Visible = False
    Sheets("Well").Select
End Sub


Private Sub CommandButton2_Click()
    If ActiveSheet.Name <> "AggChart" Then Sheets("AggChart").Select
    Call WriteAllCharts
End Sub


Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub





