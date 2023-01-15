Attribute VB_Name = "printworkbooksworksheets"
Sub listopenworkbooks()
    'list output in Immediate Windows View-->Immediate Windows or Ctrl+G
    Dim wk As Workbook
    For Each wk In Workbooks
        Debug.Print wk.Name & ", " & wk.FullName
    Next
End Sub

Sub listworksheetscurrenetworkbook()
    'list output in Immediate Windows View-->Immediate Windows or Ctrl+G
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        Debug.Print sh.Name
    Next
End Sub

Sub listworksheetsallopenworkbooks()
    'list output in Immediate Windows View-->Immediate Windows or Ctrl+G
    Dim sh As Worksheet, wk As Workbook
    For Each wk In Workbooks
        For Each sh In ThisWorkbook.Worksheets
            Debug.Print wk.Name & ":" & sh.Name
        Next sh
    Next wk
End Sub

Sub listworksheetscurrenetworkbookreverseorder()
    'list output in Immediate Windows View-->Immediate Windows or Ctrl+G
    Dim i As Long
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Debug.Print ThisWorkbook.Worksheets(i).Name
    Next
End Sub

Sub listworksheetscurrenetworkbookexceptfirstsheet()
    'list output in Immediate Windows View-->Immediate Windows or Ctrl+G
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name <> "breakcombinestatements" Then
            Debug.Print sh.Name
        End If
    Next sh
End Sub

Sub printnamepathfullname()
    'list output in Immediate Windows View-->Immediate Windows or Ctrl+G
    Debug.Print Application.UserName
    Debug.Print ThisWorkbook.Name
    Debug.Print ThisWorkbook.Path
    Debug.Print ThisWorkbook.FullName
End Sub

