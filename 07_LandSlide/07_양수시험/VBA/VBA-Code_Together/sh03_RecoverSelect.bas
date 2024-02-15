Private Sub CommandButton1_Click()
    Call recover_01
End Sub

Private Sub CommandButton2_Click()
    Call save_original
End Sub

Private Sub Worksheet_Activate()
    Dim gong1, gong2 As String
    Dim gong As Long
    Dim er As Integer
    Dim cellformula As String
    

'    gong = Val(CleanString(shInput.Range("J48").Value))
'
'    gong1 = "W-" & CStr(gong)
'    gong2 = shInput.Range("i54").Value
'
'    If gong1 <> gong2 Then
'        'MsgBox "different : " & g1 & " g2 : " & g2
'        shInput.Range("i54").Value = gong1
'    End If
    

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
    
    Range("A28").Formula = cellformula
    
End Sub


