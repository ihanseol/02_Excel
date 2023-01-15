Attribute VB_Name = "Chapter03"
Sub quickcopypaste()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("3").Activate
    'Range("A1", Range("A1").End(xlDown)).Select
    Range("A1", Range("A1").End(xlDown)).Copy
    Range("B1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
End Sub

Sub boxes()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("3").Activate
    inputboxage = InputBox("prompt how old are you", "title ask age", "default 1000")
    Range("A10").Value = inputboxage
    
    messagebox = MsgBox("Hello vbExclamation message box", vbExclamation, "title greeting")
    messageboxyesno = MsgBox("Do you like bananas?  6=vbYes 7=vbNo", vbYesNo, "title Fruits question")
    Range("A11").Value = messageboxyesno
    
    If messageboxyesno = 6 Then
        Range("A12").Value = "It's a 6 yes"
    Else
        Range("A12").Value = "It's a 7 no"
    End If
    messageboxyesnocancel = MsgBox("Do you want to save this file 2=vbCancel", _
                                   vbYesNoCancel, "title three")
    Range("A13").Value = messageboxyesnocancel
    
    If messageboxyesnocancel = 2 Then
        Range("A14").Value = "It's a 2 cancel"
    Else
        Range("A14").Value = "Already covered 6 yes 7 no"
    End If
    '1 vbOK, 3 vbAbort, 4 vbRetry, 5 vbIgnore
End Sub

Sub expandedmath()
    Application.Workbooks("vbaforexcelmadesimple.xlsm").Worksheets("3").Activate
    'Exponentiation
    Range("D1").Value = (7 ^ 4)                  'print 240
    'Integer division
    Range("D2").Value = (7 \ 2)                  'print 3
    'Integer division
    Range("E2").Value = (7 \ 3)                  'print 2
    'Modulo
    Range("D3").Value = 7 Mod 2                  'print 1
End Sub


