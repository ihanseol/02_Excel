' ***************************************************************
' UserForm_SS
'
' ***************************************************************

' Optionbutton1 - 가정용
' Optionbutton2 - 일반용
' Optionbutton3 - 청소용
' Optionbutton4 - 민방위용
' Optionbutton5 - 학교용
' Optionbutton6 - 공동주택용
' Optionbutton7 - 간이상수도
' Optionbutton8 - 농생활겸용
' Optionbutton9 - 기타


Private Sub CommandButton1_Click()
    Dim i As Integer
    Dim options() As Variant
    Dim selectedOption As String
    
    ' Assign captions to an array
    options = Array("가정용", "일반용", "청소용", "민방위용", "학교용", "공동주택용", "간이상수도", "농생활겸용", "기타")
    
    ' Loop through OptionButtons to find the selected one
    For i = 0 To 8
        If Controls("OptionButton" & i + 1).Value Then
            selectedOption = options(i)
            Exit For
        End If
    Next i
    
    ' Set the value of the active cell
    If selectedOption <> "" Then
        ActiveCell.Value = selectedOption
        Unload Me
    Else
        MsgBox "Please select an option."
    End If
End Sub

Private Sub CommandButton2_Click()
  Unload Me
End Sub

Private Sub UserForm_Initialize()
    Dim i As Integer
    
    Me.StartUpPosition = 0
    
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
   OptionButton1.Value = True
End Sub

'Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = 27 Then
'        Unload Me
'    End If
'End Sub
