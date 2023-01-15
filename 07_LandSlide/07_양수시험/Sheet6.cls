VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CommandButton1_Click()
    Call step_01
End Sub

Private Sub CommandButton2_Click()
    Call save_original
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
    
End Sub

