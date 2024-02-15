Private Sub CommandButton1_Click()
    If Workbooks.Count = 1 Then
        MsgBox "Please Open 기사용관정, 파일 ", vbOKOnly
        Exit Sub
    End If

   WB_NAME = GetOtherFileName
   
   ' Call mDeleteAllActiveXButtons(WB_NAME)
   
   
   Call DeleteAllActiveXControls(WB_NAME)
   Call SaveJustXLSX(WB_NAME)

End Sub


Private Sub CommandButton2_Click()
    If Workbooks.Count = 1 Then
        MsgBox "Please Open 기사용관정, 파일 ", vbOKOnly
        Exit Sub
    End If

   WB_NAME = GetOtherFileName
   Call DeleteHiddenSheets(WB_NAME)

End Sub


Sub DeleteHiddenSheets(ByVal fName As String)

    Dim ws As Worksheet

    If MsgBox("Are you sure you want to delete all hidden sheets in this workbook?", vbYesNo, "Delete Hidden Sheets?") <> vbYes Then Exit Sub

    Workbooks(fName).Activate
    
    Application.DisplayAlerts = False
    For Each ws In Worksheets
        If ws.Visible = xlSheetHidden Or ws.Visible = xlSheetVeryHidden Then
            ws.Delete
        End If
    Next ws
    Application.DisplayAlerts = True
    
    ThisWorkbook.Activate

End Sub


Public Function MyDownPath() As String
    MyDownPath = Environ$("USERPROFILE") & "\" & "Downloads"
    Debug.Print MyDownPath
End Function



Sub SaveJustXLSX(ByVal fName As String)
    Dim mypath, fname0 As String
    Dim fso As New Scripting.FileSystemObject
        
    mypath = MyDownPath
    Debug.Print "path" + mypath
    
    
    Workbooks(fName).Activate
    fname0 = fso.GetBaseName(fName)
    
    ActiveWorkbook.SaveAs fileName:= _
        mypath & "\" & fname0 & ".xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
        
        
    On Error Resume Next
    'Workbooks(fname0).Close SaveChanges:=False
    Workbooks("'" & fname0 & "'").Close SaveChanges:=False
    On Error GoTo 0
    
    
    
    ThisWorkbook.Activate
End Sub

Sub DeleteAllActiveXControls(ByVal fName As String)
    Dim myControl As Object
    
    
    Workbooks(fName).Activate
    
    For Each ws In Workbooks(fName).Worksheets
    
        For Each myControl In ws.OLEObjects
            myControl.Delete
        Next myControl
        
     Next ws
    
    ThisWorkbook.Activate
End Sub



Sub mDeleteAllActiveXButtons(ByVal fName As String)
    Dim ws As Worksheet
    Dim obj As OLEObject
    
    Workbooks(fName).Activate
          
    For Each ws In Workbooks(fName).Worksheets
        For Each obj In ws.OLEObjects
            If TypeName(obj.Object) = "CommandButton" Then
                obj.Delete
            End If
        Next obj
    Next ws
    
    ThisWorkbook.Activate
    
End Sub

Sub ListOpenWorkbookNames()
    Dim Workbook As Workbook
    Dim workbookNames As String
    Dim i As Long
        
    workbookNames = ""
    
    For Each Workbook In Application.Workbooks
        workbookNames = workbookNames & Workbook.Name & vbCrLf
    Next
    
    Cells(1, 1).Value = workbookNames
End Sub

'Function GetOtherFileName() As String
'    Dim Workbook As Workbook
'    Dim workbookNames As String
'    Dim bool As Boolean
'    Dim i As Long
'
'    workbookNames = ""
'    bool = False
'
'    For Each Workbook In Application.Workbooks
'        If StrComp(ThisWorkbook.Name, Workbook.Name, vbTextCompare) = 0 Then
'            bool = True
'            GoTo NEXT_ITERATION
'        End If
'
'        If bool Then
'            Exit For
'        End If
'
'NEXT_ITERATION:
'    Next
'
'    GetOtherFileName = Workbook.Name
'End Function


Function GetOtherFileName() As String
' refactor by instr function

    Dim OtherWorkbook As Workbook
    Dim ThisWorkbookName As String
    Dim OtherWorkbookName As String

    ThisWorkbookName = ThisWorkbook.Name
    
    For Each OtherWorkbook In Application.Workbooks
        If InStr(1, OtherWorkbook.Name, ThisWorkbookName, vbTextCompare) = 0 Then
            OtherWorkbookName = OtherWorkbook.Name
            Exit For
        End If
    Next OtherWorkbook
    
    GetOtherFileName = OtherWorkbookName
End Function


'
' refactor function
'
'Function GetOtherFileName() As String
'    Dim OtherWorkbook As Workbook
'    Dim ThisWorkbookName As String
'    Dim OtherWorkbookName As String
'
'    ThisWorkbookName = ThisWorkbook.Name
'
'    For Each OtherWorkbook In Application.Workbooks
'        If StrComp(ThisWorkbookName, OtherWorkbook.Name, vbTextCompare) <> 0 Then
'            OtherWorkbookName = OtherWorkbook.Name
'            Exit For
'        End If
'    Next OtherWorkbook
'
'    GetOtherFileName = OtherWorkbookName
'End Function
'




'Function CheckSubstring(str As String, chk As String) As Boolean
'
'    If InStr(str, chk) > 0 Then
'        ' The string contains "chk"
'        CheckSubstring = True
'    Else
'        ' The string does not contain "chk"
'        CheckSubstring = False
'    End If
'End Function



Function CheckSubstring(str As String, chk As String) As Boolean
    CheckSubstring = (InStr(str, chk) > 0)
End Function




