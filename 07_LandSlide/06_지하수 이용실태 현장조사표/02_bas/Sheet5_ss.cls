VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5_ss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub CommandButton1_Click()
    Call insertRow
End Sub

Private Sub CommandButton2_Click()
    ' 지하수 이용실태 현장조사표
    ' Groundwater Availability Field Survey Sheet
    
    Call MakeFieldList
    Sheets("ss").Activate
    
End Sub

Private Sub CommandButton3_Click()
    Call ComputeQ
    Sheets("ss").Activate
End Sub

Private Sub CommandButtonCopy_Click()
    Call MainMoudleGenerateCopy
End Sub

Private Sub CommandButtonDelete_Click()
    Call SubModuleCleanCopySection
End Sub

Private Sub Worksheet_Activate()
   Call clearRowA
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
