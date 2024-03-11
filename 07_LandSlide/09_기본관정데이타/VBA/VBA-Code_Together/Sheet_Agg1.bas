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

Call AggregateOne_Import(999, False)

End Sub



Private Sub CommandButton3_Click()
    ' SingleWell Import
        
    Dim singleWell  As Integer
    Dim WB_NAME As String
    
    
    WB_NAME = GetOtherFileName
    'MsgBox WB_NAME
    
    'If Workbook Is Nothing Then
    '    GetOtherFileName = "Empty"
    'Else
    '    GetOtherFileName = Workbook.name
    'End If
        
    If WB_NAME = "Empty" Then
        MsgBox "WorkBook is Empty"
        Exit Sub
    Else
        singleWell = CInt(ExtractNumberFromString(WB_NAME))
    '   MsgBox (SingleWell)
    End If
    
    Call AggregateOne_Import(singleWell, True)

End Sub


Private Sub AggregateOne_Import(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
' isSingleWellImport = True ---> SingleWell Import
' isSingleWellImport = False ---> AllWell Import
        
    Dim fName As String
    Dim nofwell, i As Integer
    Dim q1, qq1, q2, q3, ratio, C, B, S1, S2 As Double
    Dim wsYangSoo As Worksheet
    
    nofwell = GetNumberOfWell()
    Sheets("Aggregate1").Select
    
    Set wsYangSoo = Worksheets("YangSoo")
    
    
    If Not isSingleWellImport Then
        Call EraseCellData("G3:K35")
        Call EraseCellData("Q3:S35")
        Call EraseCellData("F43:I102")
    End If
    
    
    For i = 1 To nofwell

        If Not isSingleWellImport Or (isSingleWellImport And i = singleWell) Then
            GoTo SINGLE_ITERATION
        Else
            GoTo NEXT_ITERATION
        End If
        
SINGLE_ITERATION:

        q1 = wsYangSoo.Cells(4 + i, "aa").value
        qq1 = wsYangSoo.Cells(4 + i, "ac").value
        
        q2 = wsYangSoo.Cells(4 + i, "ab").value
        q3 = wsYangSoo.Cells(4 + i, "k").value
        
        ratio = wsYangSoo.Cells(4 + i, "ah").value
        
        S1 = wsYangSoo.Cells(4 + i, "ad").value
        S2 = wsYangSoo.Cells(4 + i, "ae").value
        
        C = wsYangSoo.Cells(4 + i, "af").value
        B = wsYangSoo.Cells(4 + i, "ag").value
        
        
        TurnOffStuff
        
        Call WriteWellData36_Single(q1, q2, q3, ratio, C, B, i, isSingleWellImport)
        Call Write_Tentative_water_intake_Single(qq1, S2, S1, q2, i, isSingleWellImport)
        
        TurnOnStuff
        
NEXT_ITERATION:
        
    Next i

    Application.CutCopyMode = False
    Range("L1").Select
    
End Sub



'3-6, 조사공의 적정취수량및 취수계획량
Sub WriteWellData36_Single(q1 As Variant, q2 As Variant, q3 As Variant, ratio As Variant, C As Variant, B As Variant, ByVal i As Integer, isSingleWellImport)
    
    Dim remainder As Integer
        
    If isSingleWellImport Then
        Call EraseCellData("G" & (i + 2) & ":K" & (i + 2))
        Call EraseCellData("Q" & (i + 2) & ":S" & (i + 2))
    End If
        
    Range("G" & (i + 2)).value = "W-" & i
    Range("H" & (i + 2)).value = q1
    Range("I" & (i + 2)).value = q2
    Range("J" & (i + 2)).value = q3
    Range("K" & (i + 2)).value = ratio
    
    Range("Q" & (i + 2)).value = "W-" & i
    Range("R" & (i + 2)).value = C
    Range("S" & (i + 2)).value = B
    
    remainder = i Mod 2
    If remainder = 0 Then
            Call BackGroundFill(Range(Cells(i + 2, "G"), Cells(i + 2, "K")), True)
            Call BackGroundFill(Range(Cells(i + 2, "Q"), Cells(i + 2, "S")), True)
    Else
            Call BackGroundFill(Range(Cells(i + 2, "G"), Cells(i + 2, "K")), False)
            Call BackGroundFill(Range(Cells(i + 2, "Q"), Cells(i + 2, "S")), False)
    End If

End Sub




'적정취수량의 계산
Sub Write_Tentative_water_intake_Single(q1 As Variant, S2 As Variant, S1 As Variant, q2 As Variant, i As Variant, isSingleWellImport)
    
'****************************************
' ip = 43
'****************************************
' Call EraseCellData("F43:I102")

    
    Dim ip, remainder As Variant
    Dim Values As Variant
    
    Values = GetRowColumn("Agg1_Tentative_Water_Intake")
    ip = Values(2)
    
    'Call EraseCellData("F" & ip & ":I" & (ip + nofwell - 1))
    If isSingleWellImport Then
        Call EraseCellData("F" & (ip + i - 1) & ":I" & (ip + (i - 1) * 2 + 1))
    End If
    
    Cells((ip + 0) + (i - 1) * 2, "F").value = "W-" & CStr(i)
    Cells((ip + 0) + (i - 1) * 2, "G").value = q1
    Cells((ip + 0) + (i - 1) * 2, "H").value = S2
    Cells((ip + 1) + (i - 1) * 2, "H").value = S1
    Cells((ip + 0) + (i - 1) * 2, "I").value = q2
    
    
    remainder = i Mod 2
    If remainder = 0 Then
            Call BackGroundFill(Range(Cells((ip + 0) + (i - 1) * 2, "F"), Cells((ip + 0) + (i - 1) * 2 + 1, "I")), True)
    Else
            Call BackGroundFill(Range(Cells((ip + 0) + (i - 1) * 2, "F"), Cells((ip + 0) + (i - 1) * 2 + 1, "I")), False)
    End If
    
End Sub


