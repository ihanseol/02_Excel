Attribute VB_Name = "Chapter10"
Sub boxes()
Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("10").Activate
numberinput = InputBox("Enter the number", "Title", "Default3")
Range("A1").Value = numberinput

mymessage = "vbExclamation Exclamation message box and vbYesNoCancel buttons"
response = MsgBox(mymessage, vbExclamation + vbYesNoCancel, "Title")
Select Case response
    Case Is = vbYes
        Range("A2").Value = "Yes"
    Case Is = vbNo
        Range("A2").Value = "No"
    Case Is = vbCancel
        Range("A2").Value = "Cancel"
End Select
End Sub

