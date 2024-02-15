Private Sub CommandButton1_Click()
  UserFormTS1.Show
End Sub


Private Sub CommandButton2_Click()
    Call make_adjust_value
End Sub

Private Sub CommandButton3_Click()
    Range("L14:N23").Select
    Selection.Copy
    Range("H14").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("K9").Select
    Application.CutCopyMode = False
End Sub


Private Sub SetWellTitle(ByVal gong As Integer)

    Dim strText As String
    
    strText = "W-" & CStr(gong)
    
    Range("b4").Value = "¼öÁú " & CStr(gong) & "¹ø"
    Range("c4").Value = strText
    Range("d12").Value = strText
    Range("h12").Value = strText
    Range("l12").Value = strText
    
End Sub

Private Sub Worksheet_Activate()
 '   Dim gong As Integer

'    Range("C6").Select
'    ActiveCell.FormulaR1C1 = "=LongTest!R[4]C"

'    gong = Val(CleanString(shInput.Range("J48").Value))
'    Call SetWellTitle(gong)
End Sub






   
  
  




