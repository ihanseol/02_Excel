VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Call clear_30year_data
End Sub

Private Sub CommandButton2_Click()
    Call BackupData
End Sub

Private Sub CommandButton3_Click()
' get 30 year data by Selenium

   Call get_weather_data
   Call import30RecentData
   Range("A1").Select
   
End Sub


Private Sub CommandButton4_Click()
    Call importFromArray
End Sub

Private Sub CommandButton5_Click()
    Call ShiftNewYear
End Sub

