VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Call importRainfall
End Sub

Private Sub CommandButton2_Click()
    Range("b5:n34").ClearContents
End Sub

Private Sub importRainfall()

    Dim a, i&, fn$, prefix$
    ReDim a(1 To Workbooks.count)
    Dim fname_this, fname_that As String
        
    prefix = "report to excel."
    
    If Workbooks.count = 1 Then
        MsgBox "open the file A2_웹스크레이퍼 by Selenium V1.0 NewForm.xlsm"
        Exit Sub
    End If
    
    For i = 1 To Workbooks.count
      a(i) = Workbooks(i).Name
      If Left(a(i), Len(prefix)) = prefix Then fn = a(i)
    Next i
    
    If (InStr(1, a(1), "기본관정") <> 0) Then
        fname_this = a(1)
        fname_that = a(2)
    Else
        fname_this = a(2)
        fname_that = a(1)
    End If
    
    
    On Error GoTo Errcheck
    Workbooks(fname_that).Worksheets("main").Activate
    
    Workbooks(fname_that).Worksheets("main").Range("b6:n35").Select
    Selection.Copy
    
    Windows(fname_this).Activate
    Range("B5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    
    Range("b2").value = Workbooks(fname_that).Worksheets("main").Range("b2").value
    Range("a1").Select
    Application.CutCopyMode = False
    
Errcheck:
    Err.Clear

End Sub
