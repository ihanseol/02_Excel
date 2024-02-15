' ***************************************************************
' Sheet2_aa(aa)
'
' ***************************************************************



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

Private Sub CommandButtonInitialClear_Click()
 Call SubModuleInitialClear
End Sub


Private Sub Worksheet_Activate()
    
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    
End Sub
