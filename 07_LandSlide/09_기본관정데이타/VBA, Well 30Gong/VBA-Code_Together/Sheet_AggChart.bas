Option Explicit


Private Sub CommandButton1_Click()
'Hide Aggregate

    Sheets("AggChart").Visible = False
    Sheets("Well").Select
End Sub


Private Sub CommandButton2_Click()
'Collect Data

    If ActiveSheet.name <> "AggChart" Then Sheets("AggChart").Select
    Call WriteAllCharts(999, False)

End Sub


Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub


Private Sub CommandButton3_Click()
'single well import

Dim singleWell  As Integer
Dim WB_NAME As String


WB_NAME = GetOtherFileName
'MsgBox WB_NAME

singleWell = CInt(ExtractNumberFromString(WB_NAME))
'MsgBox (SingleWell)

Call WriteAllCharts(singleWell, True)

End Sub
