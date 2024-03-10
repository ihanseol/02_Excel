Private Sub CommandButton1_Click()
    Sheets("Aggregate2").Visible = False
    Sheets("Well").Select
End Sub

Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub


Private Sub CommandButton2_Click()
' Collect All Data

Call ImportWellSpec(999, False)

End Sub


Private Sub CommandButton3_Click()
' SingleWell Import
    
Dim singleWell  As Integer
Dim WB_NAME As String


WB_NAME = GetOtherFileName
'MsgBox WB_NAME

singleWell = CInt(ExtractNumberFromString(WB_NAME))
'MsgBox (SingleWell)

Call ImportWellSpec(singleWell, True)

End Sub



Private Sub ImportWellSpec(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
    Dim fName As String
    Dim nofwell, i As Integer
    
    Dim Q() As Double          '양수량
    Dim natural() As Double    '자연수위
    Dim stable() As Double      '안정수위
    Dim recover() As Double     '회복수위
    
    Dim radius() As Double       ' 공반경
    Dim deltas() As Double       ' deltas
    Dim deltah() As Double       ' deltah : 수위강하량
    Dim daeSoo() As Double       ' 대수층 두께
    
    Dim T1() As Double            ' T1
    Dim T2() As Double            ' T2
    Dim TA() As Double            ' TA - (T1+T2)/2, TAverage
    
    Dim K() As Double
    Dim time_() As Double           ' 안정수위도달시간
    
    Dim S1() As Double            ' S1
    Dim S2() As Double            ' S2 - 스킨팩터 해석, s값
    
    Dim schultz() As Double
    Dim webber() As Double
    Dim jcob() As Double
    
    Dim skin() As Double ' skin factor
    Dim er() As Double   ' effective radius
    

    nofwell = GetNumberOfWell()
    Sheets("Aggregate2").Select
    
    ' --------------------------------------------------------------------------------------
    
    ReDim Q(1 To nofwell)
    ReDim natural(1 To nofwell)
    ReDim stable(1 To nofwell)
    ReDim recover(1 To nofwell)
    
    ReDim radius(1 To nofwell)
    ReDim deltas(1 To nofwell)
    ReDim deltah(1 To nofwell)
    ReDim daeSoo(1 To nofwell)
    
    ReDim T1(1 To nofwell)
    ReDim T2(1 To nofwell)
    ReDim TA(1 To nofwell)
    ReDim K(1 To nofwell)
    ReDim time_(1 To nofwell)
    
    ReDim S1(1 To nofwell)
    ReDim S2(1 To nofwell)
    
    ReDim shultz(1 To nofwell)
    ReDim webber(1 To nofwell)
    ReDim jcob(1 To nofwell)
    
    ReDim skin(1 To nofwell) ' skin factor
    ReDim er(1 To nofwell)   ' effective radius
    
    ' --------------------------------------------------------------------------------------
    
    If Not isSingleWellImport Then
        Call EraseCellData("C3:J33")
        Call EraseCellData("L3:Q33")
        Call EraseCellData("S3:U33")
    End If
            
            
    For i = 1 To nofwell
        ' isSingleWellImport = True ---> SingleWell Import
        ' isSingleWellImport = False ---> AllWell Import
        
        If isSingleWellImport Then
            If i = singleWell Then
                GoTo SINGLE_ITERATION
            Else
                GoTo NEXT_ITERATION
            End If
        End If
        
SINGLE_ITERATION:
   
        Q(i) = Worksheets("YangSoo").Cells(4 + i, "k").value
        
        natural(i) = Worksheets("YangSoo").Cells(4 + i, "b").value
        stable(i) = Worksheets("YangSoo").Cells(4 + i, "c").value
        recover(i) = Worksheets("YangSoo").Cells(4 + i, "d").value
        
        radius(i) = Worksheets("YangSoo").Cells(4 + i, "h").value
        
        deltas(i) = Worksheets("YangSoo").Cells(4 + i, "l").value
        deltah(i) = Worksheets("YangSoo").Cells(4 + i, "f").value
        daeSoo(i) = Worksheets("YangSoo").Cells(4 + i, "n").value
        
        
        T1(i) = Worksheets("YangSoo").Cells(4 + i, "o").value
        T2(i) = Worksheets("YangSoo").Cells(4 + i, "p").value
        TA(i) = Worksheets("YangSoo").Cells(4 + i, "q").value
        
        time_(i) = Worksheets("YangSoo").Cells(4 + i, "u").value
                
        S1(i) = Worksheets("YangSoo").Cells(4 + i, "r").value
        S2(i) = Worksheets("YangSoo").Cells(4 + i, "s").value
        K(i) = Worksheets("YangSoo").Cells(4 + i, "t").value
        
        shultz(i) = Worksheets("YangSoo").Cells(4 + i, "v").value
        webber(i) = Worksheets("YangSoo").Cells(4 + i, "w").value
        jcob(i) = Worksheets("YangSoo").Cells(4 + i, "x").value
        
        
        skin(i) = Worksheets("YangSoo").Cells(4 + i, "y").value
        er(i) = Worksheets("YangSoo").Cells(4 + i, "z").value
        
        Call WriteWellData_Single(Q(i), natural(i), stable(i), recover(i), radius(i), deltas(i), daeSoo(i), T1(i), S1(i), i)
        Call WriteData37_RadiusOfInfluence_Single(TA(i), K(i), S2(i), time_(i), deltah(i), daeSoo(i), i)
        Call WriteData36_TS_Analysis_Single(T1(i), T2(i), TA(i), S2(i), i)
        Call Write38_RadiusOfInfluence_Result_Single(shultz(i), webber(i), jcob(i), i)
        Call Wrote34_SkinFactor_Single(skin(i), er(i), i)
        
    
NEXT_ITERATION:
    
    Next i

    Range("a1").Select
    Application.CutCopyMode = False
    
End Sub


' 3-3, 3-4, 3-5 결과출력
Sub WriteWellData_Single(Q As Variant, natural As Variant, stable As Variant, recover As Variant, radius As Variant, deltas As Variant, daeSoo As Variant, T1 As Variant, S1 As Variant, ByVal i As Integer)
    
    Dim remainder As Integer
    
    ' 3-3, 장기양수시험결과 (Collect from yangsoo data)

    Range("C" & (i + 2)).value = "W-" & i
    Range("D" & (i + 2)).value = 2880
    
    Range("e" & (i + 2)).value = Q
    Range("l" & (i + 2)).value = Q
    
    Range("f" & (i + 2)).value = natural
    Range("g" & (i + 2)).value = stable
    Range("h" & (i + 2)).value = stable - natural
    
    Range("i" & (i + 2)).value = radius
    Range("j" & (i + 2)).value = deltas
    
    
    ' 3-4, aqtesolv 해석결과
    Range("m" & (i + 2)).value = radius
    Range("n" & (i + 2)).value = radius
    Range("o" & (i + 2)).value = daeSoo
    Range("p" & (i + 2)).value = T1
    Range("q" & (i + 2)).value = S1
    
    
    '3-5, 수위회복시험 결과
    Range("s" & (i + 2)).value = stable
    Range("t" & (i + 2)).value = recover
    Range("u" & (i + 2)).value = stable - recover
    
    remainder = i Mod 2
    If remainder = 0 Then
            Call BackGroundFill(Range(Cells(i + 2, "c"), Cells(i + 2, "j")), True)
            Call BackGroundFill(Range(Cells(i + 2, "l"), Cells(i + 2, "q")), True)
            Call BackGroundFill(Range(Cells(i + 2, "s"), Cells(i + 2, "u")), True)
            
    Else
            Call BackGroundFill(Range(Cells(i + 2, "c"), Cells(i + 2, "j")), False)
            Call BackGroundFill(Range(Cells(i + 2, "l"), Cells(i + 2, "q")), False)
            Call BackGroundFill(Range(Cells(i + 2, "s"), Cells(i + 2, "u")), False)
    End If
   
End Sub


' 3-7, 조사공별 수리상수
Sub WriteData37_RadiusOfInfluence_Single(TA As Variant, K As Variant, S2 As Variant, time_ As Variant, deltah As Variant, daeSoo As Variant, i As Variant)

'****************************************
'    ip = 37 'W-1 point
'****************************************

    Dim ip, remainder As Variant
    Dim unit, rngString As String
    Dim Values As Variant
    
    Values = GetRowColumn("agg2_37_roi")
    ip = Values(2)
    
    Call EraseCellData(ColumnNumberToLetter(4 + i) & ip & ":" & ColumnNumberToLetter(4 + i) & (ip + 6))
    
    
    Cells((ip + 0), (4 + i)).value = "W-" & i
    
    Cells((ip + 1), (4 + i)).value = TA
    Cells((ip + 1), (4 + i)).NumberFormat = "0.0000"
    
    Cells((ip + 2), (4 + i)).value = K
    Cells((ip + 2), (4 + i)).NumberFormat = "0.0000"
    
    
    Cells((ip + 3), (4 + i)).value = S2
    Cells((ip + 3), (4 + i)).NumberFormat = "0.0000000"
    
    Cells((ip + 4), (4 + i)).value = time_
    Cells((ip + 4), (4 + i)).NumberFormat = "0.0000"
    
    Cells((ip + 5), (4 + i)).value = deltah
    Cells((ip + 5), (4 + i)).NumberFormat = "0.00"
    
    Cells((ip + 6), (4 + i)).value = daeSoo
    
    
    remainder = i Mod 2
    If remainder = 0 Then
            Call BackGroundFill(Range(Cells(ip + 1, (i + 4)), Cells(ip + 6, (i + 4))), True)
    Else
            Call BackGroundFill(Range(Cells(ip + 1, (i + 4)), Cells(ip + 6, (i + 4))), False)
    End If
    

End Sub




' 3-6, 수리상수산정결과
Sub WriteData36_TS_Analysis_Single(T1 As Variant, T2 As Variant, TA As Variant, S2 As Variant, i As Variant)
    
'****************************************
'    ip = 48
'****************************************
' Call EraseCellData("C48:F137")
' 137 - 48 = 89

    Dim ip, remainder As Variant
    Dim unit, rngString As String
    Dim Values As Variant
    Dim nofwell As Integer
    
    
    Values = GetRowColumn("agg2_36_surisangsoo")
    ip = Values(2)
    
    
    Call EraseCellData("C" & (ip + (i - 1) * 3) & ":F" & (ip + (i - 1) * 3 + 2))
        
    
    Cells(ip + (i - 1) * 3, "C").value = "W-" & i
            
    Cells((ip + 0) + (i - 1) * 3, "D").value = "장기양수시험"
    Cells((ip + 1) + (i - 1) * 3, "D").value = "수위회복시험"
    Cells((ip + 2) + (i - 1) * 3, "D").value = "선택치"

    Cells((ip + 0) + (i - 1) * 3, "E").value = T1
    Cells((ip + 0) + (i - 1) * 3, "E").NumberFormat = "0.0000"
    
    Cells((ip + 1) + (i - 1) * 3, "E").value = T2
    Cells((ip + 1) + (i - 1) * 3, "E").NumberFormat = "0.0000"
    
    Cells((ip + 2) + (i - 1) * 3, "E").value = TA
    Cells((ip + 2) + (i - 1) * 3, "E").NumberFormat = "0.0000"
    Cells((ip + 2) + (i - 1) * 3, "E").Font.Bold = True
    
    Cells((ip + 0) + (i - 1) * 3, "F").value = S2
    Cells((ip + 0) + ip + (i - 1) * 3, "F").NumberFormat = "0.0000000"
    
    Cells((ip + 2) + (i - 1) * 3, "F").value = S2
    Cells((ip + 2) + (i - 1) * 3, "F").NumberFormat = "0.0000000"
    Cells((ip + 2) + (i - 1) * 3, "F").Font.Bold = True
    
    
    remainder = i Mod 2
    If remainder = 0 Then
            Call BackGroundFill(Range(Cells(ip + (i - 1) * 3, "C"), Cells((ip + 2) + (i - 1) * 3, "F")), True)
    Else
            Call BackGroundFill(Range(Cells(ip + (i - 1) * 3, "C"), Cells((ip + 2) + (i - 1) * 3, "F")), False)
    End If

End Sub



'3.8 영향반경
Sub Write38_RadiusOfInfluence_Result_Single(shultz As Variant, webber As Variant, jcob As Variant, i As Variant)
 
'****************************************
'    ip = 48 'W-1 point
'****************************************
' Call EraseCellData("H48:N77")
' 77 - 48 = 29


    Dim ip, remainder As Variant
    Dim unit, rngString As String
    Dim Values As Variant
    
    Values = GetRowColumn("agg2_38_roi_result")
    ip = Values(2)
    
    'Call EraseCellData("H" & ip & ":N" & (ip + nofwell - 1))
    Call EraseCellData("H" & (ip + i - 1) & ":N" & (ip + i - 1))
    
    Cells(ip + (i - 1), "h").value = "W-" & i
    Cells(ip + (i - 1), "h").NumberFormat = "0.0"
    
    Cells(ip + (i - 1), "i").value = shultz
    Cells(ip + (i - 1), "i").NumberFormat = "0.0"
    
    Cells(ip + (i - 1), "j").value = webber
    Cells(ip + (i - 1), "j").NumberFormat = "0.0"
    
    Cells(ip + (i - 1), "k").value = jcob
    Cells(ip + (i - 1), "k").NumberFormat = "0.0"

    Cells(ip + (i - 1), "l").value = Round((shultz + webber + jcob) / 3, 1)
    Cells(ip + (i - 1), "l").NumberFormat = "0.0"
    
    Cells(ip + (i - 1), "m").value = Application.WorksheetFunction.max(shultz, webber, jcob)
    Cells(ip + (i - 1), "m").NumberFormat = "0.0"
    
    Cells(ip + (i - 1), "n").value = Application.WorksheetFunction.min(shultz, webber, jcob)
    Cells(ip + (i - 1), "n").NumberFormat = "0.0"
    
    
    remainder = i Mod 2
    If remainder = 0 Then
            Call BackGroundFill(Range(Cells(ip + (i - 1), "h"), Cells(ip + (i - 1), "n")), True)
    Else
            Call BackGroundFill(Range(Cells(ip + (i - 1), "h"), Cells(ip + (i - 1), "n")), False)
    End If


End Sub



' 3.4 스킨계수
Sub Wrote34_SkinFactor_Single(skin As Variant, er As Variant, i As Variant)
    
'****************************************
'   ip = 48
'****************************************
' Call EraseCellData("P48:R77")
'****************************************

    Dim ip As Variant
    Dim unit, rngString As String
    Dim Values As Variant
    
    Values = GetRowColumn("agg2_34_skinfactor")
    ip = Values(2)
   
    Call EraseCellData("P" & (ip + i - 1) & ":R" & (ip + i - 1))
    
    Cells(ip + (i - 1), "p").value = "W-" & i
    Cells(ip + (i - 1), "q").value = skin
    Cells(ip + (i - 1), "q").NumberFormat = "0.0000"
    Cells(ip + (i - 1), "r").value = er
    Cells(ip + (i - 1), "r").NumberFormat = "0.000"
    
    remainder = i Mod 2
    If remainder = 0 Then
            Call BackGroundFill(Range(Cells(ip + (i - 1), "p"), Cells(ip + (i - 1), "r")), True)
    Else
            Call BackGroundFill(Range(Cells(ip + (i - 1), "p"), Cells(ip + (i - 1), "r")), False)
    End If

End Sub







