VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5_ss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub combobox_initialize()

    Dim tbl As ListObject
    Dim tableNAME, shNAME As String
    
    
    If ISIT_FIRST Then
        comboAREA.Clear
        
        If chkboxJIYEOL.Value = True Then
            tableNAME = "tableJIYEOL"
            shNAME = "ref1"
        Else
            tableNAME = "tableCNU"
            shNAME = "ref"
        End If
        
        Set tbl = Sheets(shNAME).ListObjects(tableNAME)
             
        Dim cell As Range
        Dim i As Integer
        Dim isFirst As Boolean: isFirst = True
        
        i = 0
        For Each cell In tbl.HeaderRowRange.Cells
            If isFirst Then
                isFirst = False
                GoTo NEXT_ITER
            End If
            
             comboAREA.AddItem cell.Value
NEXT_ITER:
        Next cell
    Else
        ISIT_FIRST = False
    End If
End Sub


Private Sub Worksheet_Activate()
    Call combobox_initialize
    Call clearRowA
End Sub

Private Sub chkboxJIYEOL_Click()
    ISIT_FIRST = True
    Call combobox_initialize
    ISIT_FIRST = False
End Sub


Private Sub comboAREA_DropButtonClick()
    'Call combobox_initialize
End Sub

Private Sub comboAREA_GotFocus()
   'Call combobox_initialize
End Sub


Private Sub comboAREA_Change()
    ' Dim selectedHeader As String
    ' selectedHeader = comboAREA.Value
    ' Range("S21").Value = selectedHeader
End Sub


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

Private Sub CommandButton4_Click()

   If Sheets("ref").Visible Then
        Sheets("ref").Visible = False
        Sheets("ref1").Visible = False
    Else
        Sheets("ref").Visible = True
        Sheets("ref1").Visible = True
    End If
    
End Sub

Private Sub CommandButtonCopy_Click()
    Call MainMoudleGenerateCopy
End Sub

Private Sub CommandButtonDelete_Click()
    Call SubModuleCleanCopySection
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub
