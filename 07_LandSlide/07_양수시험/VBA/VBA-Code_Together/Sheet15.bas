Private Sub CommandButton1_Click()
    UserFormTS1.Show
End Sub


Private Sub CommandButton2_Click()
    Dim i As Integer
    
    For i = 14 To 23
        ' Temp
        Cells(i, "h").Value = Round(myRandBetween(1, 3, 10), 1)
        
        ' EC
        Cells(i, "i").Value = myRandBetween(1, 3, 1)
        
        ' PH
        Cells(i, "j").Value = Round(myRandBetween(7, 13, 100), 2)
    Next i

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

'
' Random Generator
'

Private Sub CommandButton4_Click()
' 2024,03,11
' Random Generation by Button ...


    Dim i As Integer
    
    For i = 14 To 23
        'Temperature
        Cells(i, "L").Value = myRandBetween(1, 3, 10)
        
        'EC
        Cells(i, "M").Value = myRandBetween(1, 20, 1)
        
        'PH
        Cells(i, "N").Value = myRandBetween(8, 12, 100)
    Next i
    
End Sub



