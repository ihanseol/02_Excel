VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormSTime 
   Caption         =   "Set Stable Time - 안정수위도달시간"
   ClientHeight    =   8625.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12330
   OleObjectBlob   =   "UserFormSTime.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserFormSTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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
    
    TextBox1.Text = shLongTermTest.ComboBox1.Value
    i = gDicStableTime(CInt(shLongTermTest.ComboBox1.Value)) ' - 16
    Call SetOptionButtonClick(i)
    Unload Me
End Sub


Private Sub EnterButton_Click()
    On Error GoTo Errcheck
        shLongTermTest.ComboBox1.Value = gDicMyTime(MY_TIME)
         
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
    TextBox1.Text = shLongTermTest.ComboBox1.Value
    i = gDicStableTime(CInt(shLongTermTest.ComboBox1.Value)) ' - 16
    Call SetOptionButton(i)
    
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
      
End Sub


