VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "현재_통합_문서"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Call InitialSetColorValue
    Sheets("Well").SingleColor.value = True
    Sheets("Recharge").cbCheSoo.value = True
End Sub

Private Sub Workbook_SheetActivate(ByVal sh As Object)
 ' Call InitialSetColorValue
End Sub


