VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_AA 
   Caption         =   "AA, Contents Selection"
   ClientHeight    =   2280
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7035
   OleObjectBlob   =   "UserForm_AA.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm_AA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Optionbutton1 - 답작용
' Optionbutton2 - 전작용
' Optionbutton3 - 원예용
' Optionbutton4 - 축산용
' Optionbutton5 - 양어장용
' Optionbutton6 - 기타


'Private Sub CommandButton1_Click()
'    If OptionButton1.Value Then
'        ActiveCell.Value = "답작용"
'        Unload Me
'    End If
'
'    If OptionButton2.Value Then
'        ActiveCell.Value = "전작용"
'        Unload Me
'    End If
'
'    If OptionButton3.Value Then
'        ActiveCell.Value = "원예용"
'        Unload Me
'    End If
'
'    If OptionButton4.Value Then
'        ActiveCell.Value = "축산용"
'        Unload Me
'    End If
'
'    If OptionButton5.Value Then
'        ActiveCell.Value = "양어장용"
'        Unload Me
'    End If
'
'
'    If OptionButton6.Value Then
'        ActiveCell.Value = "기타"
'        Unload Me
'    End If
'
'
'End Sub

Private Sub CommandButton1_Click()
    Dim i As Integer
    Dim options() As Variant
    Dim selectedOption As String
    
    ' Assign captions to an array
    options = Array("답작용", "전작용", "원예용", "축산용", "양어장용", "기타")
    
    ' Loop through OptionButtons to find the selected one
    For i = 0 To 5
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




