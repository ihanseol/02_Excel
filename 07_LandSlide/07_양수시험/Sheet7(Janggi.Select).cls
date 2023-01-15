VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CommandButton1_Click()
    Call janggi_01
End Sub

Private Sub CommandButton2_Click()
    Call janggi_02
End Sub

Private Sub CommandButton3_Click()
    Call save_original
End Sub

Private Sub CommandButton4_Click()
    
    Call ToggleWellRadius
End Sub



'0 : skin factor, cell, C8
'1 : Re1,         cell, E8
'2 : Re2,         cell, H8
'3 : Re3,         cell, G10


Private Sub ToggleWellRadius()
    Dim er As Integer
    Dim cellformula As String
    
    er = GetEffectiveRadius
        
     Select Case er
        Case erRE1
            cellformula = "=SkinFactor!K8"
        
        Case erRE2
            cellformula = "=SkinFactor!K9"
            
        Case erRE3
            cellformula = "=SkinFactor!K10"
        
        Case Else
            cellformula = "=SkinFactor!C8"
    End Select

    If (Range("A27").Formula = cellformula) Then
        Range("A27").Formula = 0
    Else
        Range("A27").Formula = cellformula
    End If

End Sub

Private Sub SetEffectiveRadius()
     er = GetEffectiveRadius
            
     Select Case er
        Case erRE1
            cellformula = "=SkinFactor!K8"
        
        Case erRE2
            cellformula = "=SkinFactor!K9"
            
        Case erRE3
            cellformula = "=SkinFactor!K10"
        
        Case Else
            cellformula = "=SkinFactor!C8"
    End Select

    Range("A27").Formula = cellformula
End Sub


Private Sub Worksheet_Activate()
    Dim KeyCell As Range
    Dim gong1, gong2 As String
    Dim gong As Long
    
    Set KeyCell = shInput.Range("J48")
    
     gong = Val(CleanString(KeyCell.Value))
    
    
    gong1 = "W-" & CStr(gong)
    gong2 = shInput.Range("i54").Value
 
    If gong1 <> gong2 Then
        'MsgBox "different : " & g1 & " g2 : " & g2
        shInput.Range("i54").Value = gong1
    End If
    
    Call SetEffectiveRadius
End Sub


