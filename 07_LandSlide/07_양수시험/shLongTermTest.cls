VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shLongTermTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CommandButton2_Click()
    Call frame_time_setting
    Call TimeSetting
End Sub

Private Sub frame_time_setting()
    Dim i As Integer
    
    For i = 6 To 19
        If (Frame1.Controls("OptionButton" & CStr(i)).Value) Then MY_TIME = (38 + i - 11)
    Next i
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

Private Sub Frame1_Click()
    Frame1.Controls("optionbutton13").Value = True
End Sub


Private Sub Worksheet_Activate()
    Dim gong1, gong2 As String
    Dim gong, occur As Long
    Dim KeyCell As Range
    
    Debug.Print ActiveSheet.name
       
    If ActiveSheet.name <> "장기양수시험" Then
        Exit Sub
    End If
   
    
    If MY_TIME = 0 Then
        MY_TIME = initialize_myTime
        OptionButton_Setting (MY_TIME)
    End If

   Set KeyCell = shInput.Range("J48")
   
   gong = Val(CleanString(KeyCell.Value))
     
    gong1 = "W-" & CStr(gong)
    gong2 = shInput.Range("i54").Value
 
    If gong1 <> gong2 Then
        shInput.Range("i54").Value = gong1
    End If
    
End Sub

