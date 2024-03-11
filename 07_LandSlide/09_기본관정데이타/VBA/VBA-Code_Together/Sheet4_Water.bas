Private Sub CommandButton1_Click()
   Sheets("water").Visible = False
    Sheets("Well").Select
End Sub

Private Sub CommandButton2_Click()
    Dim WB_NAME, cpRange  As String

    If Workbooks.count = 1 Then
        MsgBox "Please Open 기사용관정, 파일 ", vbOKOnly
        Exit Sub
    End If
   
    
    WB_NAME = GetOtherFileName
    
    cpRange = GetCopyPoint(WB_NAME)
    Call CopyFromGWAN_JUNG(WB_NAME, cpRange)
    Call FormulaInjection
    
End Sub

' 2024-01-14
' inject formula ...

Private Sub FormulaInjection()
    Dim nofwell, i As Integer
    
    nofwell = GetNumberOfWell()
    For i = 4 To nofwell + 3
        Sheets("Well").Cells(i, "O").formula = "=ROUND(water!$F$7, 1)"
    Next i

End Sub


Function GetOtherFileName() As String
    Dim Workbook As Workbook
    Dim workbookNames As String
    Dim i As Long

    workbookNames = ""
    
    For Each Workbook In Application.Workbooks
        If StrComp(ThisWorkbook.name, Workbook.name, vbTextCompare) = 0 Then
            GoTo NEXT_ITERATION
        End If
        
        If CheckSubstring(Workbook.name, "관정") Then
            Exit For
        End If
        
NEXT_ITERATION:
    Next
    
    GetOtherFileName = Workbook.name
End Function


Function CheckSubstring(str As String, chk As String) As Boolean
    
    If InStr(str, chk) > 0 Then
        ' The string contains "chk"
        CheckSubstring = True
    Else
        ' The string does not contain "chk"
        CheckSubstring = False
    End If
End Function


'
'Function lastRowByKey(cell As String) As Long
'    lastRowByKey = Range(cell).End(xlDown).Row
'End Function


Function GetCopyPoint(ByVal fName As String) As String

  Dim ip1, ip2 As Integer

  ip1 = Workbooks(fName).Worksheets("ss").Range("b1").End(xlDown).Row + 4
  ip2 = ip1 + 2
  
  GetCopyPoint = "B" & ip1 & ":J" & ip2
  ThisWorkbook.Activate

End Function


Sub CopyFromGWAN_JUNG(ByVal fName As String, ByVal cpRange As String)

    Workbooks(fName).Worksheets("ss").Activate
    Workbooks(fName).Worksheets("ss").Range(cpRange).Select
    Selection.Copy
    
    ThisWorkbook.Sheets("water").Activate
    
    Range("d6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End Sub

Sub ListOpenWorkbookNames()
    Dim Workbook As Workbook
    Dim workbookNames As String
    Dim i As Long
        
    workbookNames = ""
    
    For Each Workbook In Application.Workbooks
        workbookNames = workbookNames & Workbook.name & vbCrLf
    Next
    
    Cells(1, 1).value = workbookNames
End Sub

