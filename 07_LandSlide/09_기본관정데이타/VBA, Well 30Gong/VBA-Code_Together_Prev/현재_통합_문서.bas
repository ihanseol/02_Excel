Private Sub Workbook_Open()
    Call InitialSetColorValue
    Sheets("Well").SingleColor.value = True
    Sheets("Recharge").cbCheSoo.value = True
End Sub

Private Sub Workbook_SheetActivate(ByVal sh As Object)
 ' Call InitialSetColorValue
End Sub


