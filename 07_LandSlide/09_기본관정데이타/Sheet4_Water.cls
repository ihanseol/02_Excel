VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4_Water"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
End Sub

'
'Function lastRowByKey(cell As String) As Long
'    lastRowByKey = Range(cell).End(xlDown).Row
'End Function


Function GetCopyPoint(ByVal fname As String) As String

  Dim ip1, ip2 As Integer

  ip1 = Workbooks(fname).Worksheets("ss").Range("b1").End(xlDown).Row + 4
  ip2 = ip1 + 2
  
  GetCopyPoint = "B" & ip1 & ":J" & ip2
  ThisWorkbook.Activate

End Function


Sub CopyFromGWAN_JUNG(ByVal fname As String, ByVal cpRange As String)

    Workbooks(fname).Worksheets("ss").Activate
    Workbooks(fname).Worksheets("ss").Range(cpRange).Select
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
