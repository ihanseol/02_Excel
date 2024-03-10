Private Sub CommandButton1_Click()
    
    Sheets("Aggregate1").Visible = False
    Sheets("Well").Select
    
End Sub

'q1 - 한계양수량 - b13
'q2 - 가채수량 - b7
'q3 - 취수계획량 - b15
'ratio - b11
'qq1 - 1단계 양수량


' Agg1_Tentative_Water_Intake : 적정취수량의 계산




Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub


Private Sub CommandButton2_Click()
' Collect Data

    Dim fName As String
    Dim nofwell, i As Integer
    Dim q1() As Double
    Dim qq1() As Double
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
    ReDim qq1(1 To nofwell) '1단계 양수량
    
    ReDim ratio(1 To nofwell)
    
    ReDim C(1 To nofwell)
    ReDim B(1 To nofwell)
    
    ReDim S1(1 To nofwell)
    ReDim S2(1 To nofwell)
    
    
    Call EraseCellData("G3:K35")
    Call EraseCellData("Q3:S35")
    
    
    For i = 1 To nofwell
        q1(i) = Worksheets("YangSoo").Cells(4 + i, "aa").value
        qq1(i) = Worksheets("YangSoo").Cells(4 + i, "ac").value
        
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
    Call Write_Tentative_water_intake(qq1, S2, S1, q2, nofwell)
    
    
    Application.CutCopyMode = False
End Sub


'적정취수량의 계산
Sub Write_Tentative_water_intake(q1 As Variant, S2 As Variant, S1 As Variant, q2 As Variant, nofwell As Variant)
    
'****************************************
' ip = 43
'****************************************
' Call EraseCellData("F43:I102")

    
    Dim i, ip, remainder As Variant
    Dim Values As Variant
    
    Values = GetRowColumn("Agg1_Tentative_Water_Intake")
    ip = Values(2)
    
    Call EraseCellData("F" & ip & ":I" & (ip + nofwell - 1))
    
        
    For i = 1 To nofwell
        Cells((ip + 0) + (i - 1) * 2, "F").value = "W-" & CStr(i)
    
        Cells((ip + 0) + (i - 1) * 2, "G").value = q1(i)
        
        Cells((ip + 0) + (i - 1) * 2, "H").value = S2(i)
        Cells((ip + 1) + (i - 1) * 2, "H").value = S1(i)
    
        Cells((ip + 0) + (i - 1) * 2, "I").value = q2(i)
        
        
        remainder = i Mod 2
        If remainder = 0 Then
                Call BackGroundFill(Range(Cells((ip + 0) + (i - 1) * 2, "F"), Cells((ip + 0) + (i - 1) * 2 + 1, "I")), True)
        Else
                Call BackGroundFill(Range(Cells((ip + 0) + (i - 1) * 2, "F"), Cells((ip + 0) + (i - 1) * 2 + 1, "I")), False)
        End If
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


'3-6, 조사공의 적정취수량및 취수계획량
Sub WriteWellData36(q1 As Variant, q2 As Variant, q3 As Variant, ratio As Variant, C As Variant, B As Variant, ByVal nofwell As Integer)
    
    Dim i, remainder As Integer
    
    For i = 1 To nofwell
        Range("G" & (i + 2)).value = "W-" & i
        Range("H" & (i + 2)).value = q1(i)
        Range("I" & (i + 2)).value = q2(i)
        Range("J" & (i + 2)).value = q3(i)
        Range("K" & (i + 2)).value = ratio(i)
        
        Range("Q" & (i + 2)).value = "W-" & i
        Range("R" & (i + 2)).value = C(i)
        Range("S" & (i + 2)).value = B(i)
        
        remainder = i Mod 2
        If remainder = 0 Then
                Call BackGroundFill(Range(Cells(i + 2, "G"), Cells(i + 2, "K")), True)
                Call BackGroundFill(Range(Cells(i + 2, "Q"), Cells(i + 2, "S")), True)
        Else
                Call BackGroundFill(Range(Cells(i + 2, "G"), Cells(i + 2, "K")), False)
                Call BackGroundFill(Range(Cells(i + 2, "Q"), Cells(i + 2, "S")), False)
        End If
        
    Next i
   
    Range("N3").value = Application.min(ratio)
    Range("O3").value = Application.max(ratio)
    
    Range("N4").value = Application.min(q2)
    Range("O4").value = Application.max(q2)
    
    Range("N5").value = Application.min(q3)
    Range("O5").value = Application.max(q3)

End Sub


'
'Private Sub CommandButton3_Click()
'    Range("g3:k19").Select
'    Selection.ClearContents
'
'    Range("n3:o5").Select
'    Selection.ClearContents
'
'    Range("q3:s19").Select
'    Selection.ClearContents
'
'    Range("B24").Select
'    Application.CutCopyMode = False
'End Sub
