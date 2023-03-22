VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_Agg1"
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
' Collect Data

    Dim fname As String
    Dim nofwell, i As Integer
    Dim q1() As Double
    Dim q2() As Double
    Dim q3() As Double
    
    Dim ratio() As Double
    
    Dim C() As Double
    Dim B() As Double
    
    Dim S1() As Double
    Dim S2() As Double
    
    
    nofwell = GetNumberOfWell()
    Sheets("Aggregate1").Select
    
    ReDim q1(1 To nofwell) '한계양수량
    ReDim q2(1 To nofwell) '적정취수량
    ReDim q3(1 To nofwell) '취수계획량
    
    ReDim ratio(1 To nofwell)
    
    ReDim C(1 To nofwell)
    ReDim B(1 To nofwell)
    
    ReDim S1(1 To nofwell)
    ReDim S2(1 To nofwell)
    
    
     With Range("g3:k19")
        .value = ""
    End With
        
     With Range("q3:s19")
        .value = ""
    End With
    
    
    For i = 1 To nofwell
'        fname = "A" & CStr(i) & "_ge_OriginalSaveFile.xlsm"
'        If Not IsWorkBookOpen(fname) Then
'            MsgBox "Please open the yangsoo data ! " & fname
'            Exit Sub
'        End If
        
'        q1(i) = Workbooks(fname).Worksheets("SafeYield").Range("b13").value  ' 한계양수량
'        q2(i) = Workbooks(fname).Worksheets("SafeYield").Range("b7").value   ' 가채수량
'        q3(i) = Workbooks(fname).Worksheets("SafeYield").Range("b15").value   ' 취수계획량
'        ratio(i) = Workbooks(fname).Worksheets("SafeYield").Range("b11").value
'
'        S1(i) = Workbooks(fname).Worksheets("SafeYield").Range("b3").value   ' sd1
'        S2(i) = Workbooks(fname).Worksheets("SafeYield").Range("b4").value   ' sd2
'
'        C(i) = Workbooks(fname).Worksheets("Input").Range("A31").value
'        B(i) = Workbooks(fname).Worksheets("Input").Range("B31").value


        q1(i) = Worksheets("YangSoo").Cells(4 + i, "aa").value
        q2(i) = Worksheets("YangSoo").Cells(4 + i, "ab").value
        q3(i) = Worksheets("YangSoo").Cells(4 + i, "k").value
        
        ratio(i) = Worksheets("YangSoo").Cells(4 + i, "ah").value
        
        S1(i) = Worksheets("YangSoo").Cells(4 + i, "ad").value
        S2(i) = Worksheets("YangSoo").Cells(4 + i, "ae").value
        
        C(i) = Worksheets("YangSoo").Cells(4 + i, "af").value
        B(i) = Worksheets("YangSoo").Cells(4 + i, "ag").value
        
    Next i

    Call WriteWellData36(q1, q2, q3, ratio, C, B, nofwell)
    Call TransPoseWellData(nofwell)
    Call Write_Tentative_water_intake(q1, S2, S1, q2, nofwell)
    
    
    Application.CutCopyMode = False
End Sub


Sub Write_Tentative_water_intake(q1 As Variant, S2 As Variant, S1 As Variant, q2 As Variant, nofwell As Variant)
    Dim i As Integer
    
    ' Range("E32:F67").Clear
    
    With Range("f27:i49")
        .value = ""
    End With
        
    For i = 1 To nofwell
        Cells(27 + (i - 1) * 2, "F").value = "W-" & CStr(i)
    
        Cells(27 + (i - 1) * 2, "G").value = q1(i)
        
        Cells(27 + (i - 1) * 2, "H").value = S2(i)
        Cells(28 + (i - 1) * 2, "H").value = S1(i)
    
        Cells(27 + (i - 1) * 2, "I").value = q2(i)
    Next i
End Sub

Sub TransPoseWellData(ByVal nofwell As Integer)
    
'    Range("i3:i" & (nofwell + 2)).Select
'    Selection.Copy
'    Range("M23").Select
'    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
'        False, Transpose:=True
'
'    Range("j3:j" & (nofwell + 2)).Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Range("M24").Select
'    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
'        False, Transpose:=True

End Sub


Sub WriteWellData36(q1 As Variant, q2 As Variant, q3 As Variant, ratio As Variant, C As Variant, B As Variant, ByVal nofwell As Integer)
    
    Dim i As Integer
    
    For i = 1 To nofwell
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
    Range("g3:k19").Select
    Selection.ClearContents
    
    Range("n3:o5").Select
    Selection.ClearContents
     
    Range("q3:s19").Select
    Selection.ClearContents
    
    Range("B24").Select
    Application.CutCopyMode = False
End Sub
