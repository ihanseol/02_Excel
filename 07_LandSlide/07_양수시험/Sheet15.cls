VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
        
    Dim gong As Integer
    Dim KeyCell As Range
  
    Range("C6").Select
    ActiveCell.FormulaR1C1 = "=LongTest!R[4]C"
    
    
    Set KeyCell = shInput.Range("J48")
    gong = Val(CleanString(KeyCell.Value))
    Call SetWellTitle(gong)
    
End Sub




   
  
  
