VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet71"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Call recover_01
End Sub

Private Sub CommandButton2_Click()
    Call save_original
End Sub

Private Sub Worksheet_Activate()
    Dim KeyCell As Range
    Dim gong1, gong2 As String
    Dim gong As Long
    Dim er As Integer
    Dim cellformula As String
    
    
    Set KeyCell = shInput.Range("J48")
    
      gong = Val(CleanString(KeyCell.Value))
    
    gong1 = "W-" & CStr(gong)
    gong2 = shInput.Range("i54").Value
 
    If gong1 <> gong2 Then
        'MsgBox "different : " & g1 & " g2 : " & g2
        shInput.Range("i54").Value = gong1
    End If
    

    er = GetEffectiveRadius
        
     Select Case er
        Case erRE1
            cellformula = "=SkinFactor!E8"
        
        Case erRE2
            cellformula = "=SkinFactor!H8"
            
        Case erRE3
            cellformula = "=SkinFactor!G10"
        
        Case Else
            cellformula = "=SkinFactor!C8"
    End Select
    
    Range("A28").Formula = cellformula
    
End Sub

