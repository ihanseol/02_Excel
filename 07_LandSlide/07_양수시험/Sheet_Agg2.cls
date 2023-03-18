VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_Agg2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
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
' Collect Data

    Dim fname As String
    Dim nofwell, i As Integer
    
    Dim q() As Double          '양수량
    Dim Natural() As Double    '자연수위
    Dim Stable() As Double      '안정수위
    Dim Recover() As Double     '회복수위
    
    Dim Radius() As Double       ' 공반경
    Dim deltas() As Double       ' deltas
    Dim deltah() As Double       ' deltah : 수위강하량
    Dim daeSoo() As Double       ' 대수층 두께
    
    Dim T1() As Double            ' T1
    Dim T2() As Double            ' T2
    Dim TA() As Double            ' TA - (T1+T2)/2, TAverage
    
    Dim K() As Double
    Dim time_() As Double           ' 안정수위도달시간
    
    Dim S1() As Double            ' S1
    Dim s2() As Double            ' S2 - 스킨팩터 해석, s값
    
    Dim schultz() As Double
    Dim webber() As Double
    Dim jcob() As Double
    
    Dim skin() As Double ' skin factor
    Dim er() As Double   ' effective radius
    

    nofwell = GetNumberOfWell()
    Sheets("Aggregate2").Select
    
    ' --------------------------------------------------------------------------------------
    
    ReDim q(1 To nofwell)
    ReDim Natural(1 To nofwell)
    ReDim Stable(1 To nofwell)
    ReDim Recover(1 To nofwell)
    
    ReDim Radius(1 To nofwell)
    ReDim deltas(1 To nofwell)
    ReDim deltah(1 To nofwell)
    ReDim daeSoo(1 To nofwell)
    
    ReDim T1(1 To nofwell)
    ReDim T2(1 To nofwell)
    ReDim TA(1 To nofwell)
    ReDim K(1 To nofwell)
    ReDim time_(1 To nofwell)
    
    ReDim S1(1 To nofwell)
    ReDim s2(1 To nofwell)
    
    ReDim shultz(1 To nofwell)
    ReDim webber(1 To nofwell)
    ReDim jcob(1 To nofwell)
    
    ReDim skin(1 To nofwell) ' skin factor
    ReDim er(1 To nofwell)   ' effective radius
    
    ' --------------------------------------------------------------------------------------
    
    Call EraseCellData("d3:j16")
    Call EraseCellData("l3:q16")
    Call EraseCellData("s3:u16")
            
            
    For i = 1 To nofwell
    
        fname = "A" & CStr(i) & "_ge_OriginalSaveFile.xlsm"
        If Not IsWorkBookOpen(fname) Then
            MsgBox "Please open the yangsoo data ! " & fname
            Exit Sub
        End If
        
        q(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c4").value
        Natural(i) = Workbooks(fname).Worksheets("SkinFactor").Range("i4").value
        Stable(i) = Workbooks(fname).Worksheets("SkinFactor").Range("i6").value
        Recover(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c10").value
        
        Radius(i) = Workbooks(fname).Worksheets("SkinFactor").Range("e4").value
        
        deltas(i) = Workbooks(fname).Worksheets("SkinFactor").Range("b4").value
        deltah(i) = Workbooks(fname).Worksheets("SkinFactor").Range("b16").value
        daeSoo(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c16").value
        
        
        T1(i) = Workbooks(fname).Worksheets("SkinFactor").Range("d5").value
        T2(i) = Workbooks(fname).Worksheets("SkinFactor").Range("h13").value
        TA(i) = Workbooks(fname).Worksheets("SkinFactor").Range("d16").value
        
        time_(i) = Workbooks(fname).Worksheets("SkinFactor").Range("h16").value
                
        S1(i) = Workbooks(fname).Worksheets("SkinFactor").Range("e10").value
        s2(i) = Workbooks(fname).Worksheets("SkinFactor").Range("i16").value
        K(i) = Workbooks(fname).Worksheets("SkinFactor").Range("E16").value
        
        shultz(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c13").value
        webber(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c18").value
        jcob(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c23").value
        
        
        skin(i) = Workbooks(fname).Worksheets("SkinFactor").Range("g6").value
        er(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c8").value
    
    Next i

    Call WriteWellData(q, Natural, Stable, Recover, Radius, deltas, daeSoo, T1, S1, nofwell)
    Call WriteData37_RadiusOfInfluence(T2, K, s2, time_, deltah, daeSoo, nofwell)
    Call WriteData36_TS_Analysis(T1, T2, TA, s2, nofwell)
    Call Write38_RadiusOfInfluence_Result(shultz, webber, jcob, nofwell)
    Call Wrote34_SkinFactor(skin, er, nofwell)
    
    Range("a1").Select
    Application.CutCopyMode = False
    
End Sub


Sub Wrote34_SkinFactor(skin As Variant, er As Variant, nofwell As Variant)
    Dim i As Integer
    
    Call EraseCellData("o32:q45")
    
    For i = 1 To nofwell
        Cells(32 + (i - 1), "o").value = "W-" & CStr(i)
'        Cells(32 + (i - 1), "p").value = Application.WorksheetFunction.Text(skin(i), "#0.0000")
'        Cells(32 + (i - 1), "q").value = Application.WorksheetFunction.Text(er(i), "#0.000")
            
         Cells(32 + (i - 1), "p").value = skin(i)
         Cells(32 + (i - 1), "p").NumberFormat = "0.0000"
         
         Cells(32 + (i - 1), "q").value = er(i)
         Cells(32 + (i - 1), "q").NumberFormat = "0.000"
    Next i
End Sub

Sub Write38_RadiusOfInfluence_Result(shultz As Variant, webber As Variant, jcob As Variant, nofwell As Variant)

    Dim i As Integer
    Dim sum, average As Double
    
    Call EraseCellData("h32:m45")
    
    For i = 1 To nofwell
        Cells(32 + (i - 1), "h").value = "W-" & CStr(i)
        Cells(32 + (i - 1), "i").value = shultz(i)
        Cells(32 + (i - 1), "j").value = webber(i)
        Cells(32 + (i - 1), "k").value = jcob(i)
    
        Cells(32 + (i - 1), "l").value = Round((shultz(i) + webber(i) + jcob(i)) / 3, 1)
        Cells(32 + (i - 1), "m").value = Application.WorksheetFunction.max(shultz(i), webber(i), jcob(i))
    Next i

End Sub



Sub WriteData36_TS_Analysis(T1 As Variant, T2 As Variant, TA As Variant, s2 As Variant, nofwell As Variant)
    Dim i As Integer
    
    Call EraseCellData("c32:f67")
    
    For i = 1 To nofwell
        Cells(32 + (i - 1) * 3, "C").value = "W-" & CStr(i)
        
        
        Cells(32 + (i - 1) * 3, "D").value = "장기양수시험"
        Cells(33 + (i - 1) * 3, "D").value = "수위회복시험"
        Cells(34 + (i - 1) * 3, "D").value = "선택치"
    
        Cells(32 + (i - 1) * 3, "E").value = T1(i)
        Cells(33 + (i - 1) * 3, "E").value = T2(i)
        Cells(34 + (i - 1) * 3, "E").value = TA(i)
        
        Cells(32 + (i - 1) * 3, "F").value = s2(i)
        Cells(34 + (i - 1) * 3, "F").value = s2(i)
    Next i
End Sub



Sub WriteData37_RadiusOfInfluence(T2 As Variant, K As Variant, s2 As Variant, time_ As Variant, deltah As Variant, daeSoo As Variant, nofwell As Variant)

    Dim i As Integer
       
    Call EraseCellData("e21:p27")
    
    For i = 1 To nofwell
        Cells(21, Chr(Asc("D") + i)).value = "W-" & CStr(i)
        Cells(22, Chr(Asc("D") + i)).value = T2(i)
        Cells(23, Chr(Asc("D") + i)).value = K(i)
        Cells(24, Chr(Asc("D") + i)).value = s2(i)
        Cells(25, Chr(Asc("D") + i)).value = time_(i)
        Cells(26, Chr(Asc("D") + i)).value = deltah(i)
        Cells(27, Chr(Asc("D") + i)).value = daeSoo(i)
    Next i

End Sub


Sub WriteWellData(q As Variant, Natural As Variant, Stable As Variant, Recover As Variant, Radius As Variant, deltas As Variant, daeSoo As Variant, T1 As Variant, S1 As Variant, ByVal nofwell As Integer)
    
    Dim i As Integer
    
    For i = 1 To nofwell
        Range("D" & (i + 2)).value = "W-" & i
        
        Range("e" & (i + 2)).value = q(i)
        Range("l" & (i + 2)).value = q(i)
        
        Range("f" & (i + 2)).value = Natural(i)
        Range("g" & (i + 2)).value = Stable(i)
        Range("h" & (i + 2)).value = Stable(i) - Natural(i)
        
        Range("i" & (i + 2)).value = Radius(i)
        Range("j" & (i + 2)).value = deltas(i)
        
        Range("m" & (i + 2)).value = Radius(i)
        Range("n" & (i + 2)).value = Radius(i)
        Range("o" & (i + 2)).value = daeSoo(i)
        Range("p" & (i + 2)).value = T1(i)
        Range("q" & (i + 2)).value = S1(i)
        
        Range("s" & (i + 2)).value = Stable(i)
        Range("t" & (i + 2)).value = Recover(i)
        Range("u" & (i + 2)).value = Stable(i) - Recover(i)
    Next i
   
End Sub

Private Sub CommandButton3_Click()
    Range("d3:u16").Select
    Selection.ClearContents
    
    Range("B24").Select
    Application.CutCopyMode = False
End Sub


