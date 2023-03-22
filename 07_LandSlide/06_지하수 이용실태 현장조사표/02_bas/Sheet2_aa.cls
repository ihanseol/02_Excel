VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2_aa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Call MainMoudleGenerateCopy
End Sub

Private Sub CommandButton2_Click()
    Call SubModuleCleanCopySection
End Sub

Private Sub CommandButton3_Click()
    Call insertRow
End Sub

Private Sub CommandButton4_Click()
    Call ComputeQ
    Sheets("aa").Activate
End Sub

Private Sub CommandButton5_Click()
    ' 지하수 이용실태 현장조사표
    ' Groundwater Availability Field Survey Sheet
    
    Call MakeFieldList
    Sheets("aa").Activate
End Sub

Private Sub Worksheet_Activate()
    Call clearRowA
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
