VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    
    Sheets("Aggregate1").Visible = False
    Sheets("Well").Select
    
End Sub

'q1 - 한계양수량 - b13
'q2 - 가채수량 - b7
'q3 - 취수계획량 - b15
'ratio - b11

Private Sub CommandButton2_Click()
    Dim fname As String
    Dim nofWell, i As Integer
    Dim q1() As Double
    Dim q2() As Double
    Dim q3() As Double
    Dim ratio() As Double
    Dim C() As Double
    Dim B() As Double
    
    nofWell = GetNumberOfWell()
    Sheets("Aggregate1").Select
    
    ReDim q1(1 To nofWell)
    ReDim q2(1 To nofWell)
    ReDim q3(1 To nofWell)
    ReDim ratio(1 To nofWell)
    ReDim C(1 To nofWell)
    ReDim B(1 To nofWell)
    
    
    For i = 1 To nofWell
        fname = "A" & CStr(i) & "_ge_OriginalSaveFile.xlsm"
        If Not IsWorkBookOpen(fname) Then
            MsgBox "Please open the yangsoo data ! " & fname
            Exit Sub
        End If
        
        q1(i) = Workbooks(fname).Worksheets("SafeYield").Range("b13").value
        q2(i) = Workbooks(fname).Worksheets("SafeYield").Range("b7").value
        q3(i) = Workbooks(fname).Worksheets("SafeYield").Range("b15").value
        ratio(i) = Workbooks(fname).Worksheets("SafeYield").Range("b11").value
        
        C(i) = Workbooks(fname).Worksheets("Input").Range("A31").value
        B(i) = Workbooks(fname).Worksheets("Input").Range("B31").value
        
    Next i

    Call WriteWellData(q1, q2, q3, ratio, C, B, nofWell)
    Call TransPoseWellData(nofWell)
    Application.CutCopyMode = False
End Sub

Sub TransPoseWellData(ByVal nofWell As Integer)
    
    Range("i3:i" & (nofWell + 2)).Select
    Selection.Copy
    Range("M23").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True

    Range("j3:j" & (nofWell + 2)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("M24").Select
    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True

End Sub


Sub WriteWellData(q1 As Variant, q2 As Variant, q3 As Variant, ratio As Variant, C As Variant, B As Variant, ByVal nofWell As Integer)
    
    Dim i As Integer
    
    For i = 1 To nofWell
        Range("G" & (i + 2)).value = "W-" & i
        Range("H" & (i + 2)).value = q1(i)
        Range("I" & (i + 2)).value = q2(i)
        Range("J" & (i + 2)).value = q3(i)
        Range("K" & (i + 2)).value = ratio(i)
        
        Range("Q" & (i + 2)).value = "W-" & i
        Range("R" & (i + 2)).value = C(i)
        Range("S" & (i + 2)).value = B(i)
    Next i
   
    Range("N3").value = Application.min(ratio)
    Range("O3").value = Application.max(ratio)
    
    Range("N4").value = Application.min(q2)
    Range("O4").value = Application.max(q2)
    
    Range("N5").value = Application.min(q3)
    Range("O5").value = Application.max(q3)

End Sub

Private Sub CommandButton3_Click()
    Range("G3:K29").Select
    Selection.ClearContents
    
    Range("n3:o29").Select
    Selection.ClearContents
     
    Range("q3:s29").Select
    Selection.ClearContents
    
    Range("B24").Select
    Application.CutCopyMode = False
End Sub
