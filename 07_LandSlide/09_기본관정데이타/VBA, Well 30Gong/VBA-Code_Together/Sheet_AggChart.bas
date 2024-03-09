Option Explicit


Private Sub CommandButton1_Click()
    Sheets("AggChart").Visible = False
    Sheets("Well").Select
End Sub


Private Sub CommandButton2_Click()
    If ActiveSheet.name <> "AggChart" Then Sheets("AggChart").Select
    Call WriteAllCharts
End Sub


Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub





