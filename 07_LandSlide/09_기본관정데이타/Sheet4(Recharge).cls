VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Call getMotorPower
End Sub

Private Sub CommandButton2_Click()
    Rows("36:53").Select
    Selection.Delete Shift:=xlUp
    Range("g34").Select
End Sub

Private Sub ShowLocation_Click()
      Sheets("location").Visible = True
      Sheets("location").Activate
End Sub


Private Sub CommandButton3_Click()
    Dim i As Integer
    Dim max, min As Single
    
    max = Range("o15").value
    min = Range("o16").value
    
    Range("B5:P14").Select
    Selection.Font.Bold = False
     
    Range("a1").Activate
    
    For i = 5 To 14
        If Cells(i, "O").value = max Or Cells(i, "O").value = min Then
            Union(Cells(i, "B"), Cells(i, "O")).Select
            Selection.Font.Bold = True
        End If
    Next i
End Sub


