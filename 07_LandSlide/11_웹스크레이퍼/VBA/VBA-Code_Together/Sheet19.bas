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

