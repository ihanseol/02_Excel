VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Sheets("Aggregate2").Visible = False
    Sheets("Well").Select
End Sub


Private Sub CommandButton2_Click()
' Collect Data

    Dim fname As String
    Dim nofWell, i As Integer
    
    Dim Q() As Double          '양수량
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
    
        
    
    nofWell = GetNumberOfWell()
    Sheets("Aggregate2").Select
    
    ' --------------------------------------------------------------------------------------
    
    ReDim Q(1 To nofWell)
    ReDim Natural(1 To nofWell)
    ReDim Stable(1 To nofWell)
    ReDim Recover(1 To nofWell)
    
    ReDim Radius(1 To nofWell)
    ReDim deltas(1 To nofWell)
    ReDim deltah(1 To nofWell)
    ReDim daeSoo(1 To nofWell)
    
    ReDim T1(1 To nofWell)
    ReDim T2(1 To nofWell)
    ReDim TA(1 To nofWell)
    ReDim K(1 To nofWell)
    ReDim time_(1 To nofWell)
    
    ReDim S1(1 To nofWell)
    ReDim s2(1 To nofWell)
    
    ReDim shultz(1 To nofWell)
    ReDim webber(1 To nofWell)
    ReDim jcob(1 To nofWell)
    
    
    ' --------------------------------------------------------------------------------------
    
    For i = 1 To nofWell
    
        fname = "A" & CStr(i) & "_ge_OriginalSaveFile.xlsm"
        If Not IsWorkBookOpen(fname) Then
            MsgBox "Please open the yangsoo data ! " & fname
            Exit Sub
        End If
        
        Q(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c4").value
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
        
        
    Next i

    Call WriteWellData(Q, Natural, Stable, Recover, Radius, deltas, daeSoo, T1, S1, nofWell)
    Call WriteData37_RadiusOfInfluence(T2, K, s2, time_, deltah, daeSoo, nofWell)
    Call WriteData36_TS_Analysis(T1, T2, TA, s2, nofWell)
    Call Write38_RadiusOfInfluence_Result(shultz, webber, jcob, nofWell)
    
    Range("a1").Select
    Application.CutCopyMode = False
    
End Sub


Sub Write38_RadiusOfInfluence_Result(shultz As Variant, webber As Variant, jcob As Variant, nofWell As Variant)

    Dim i As Integer
    Dim sum, average As Double
    
    With Range("i32:m45")
        .value = ""
    End With
            
    
    For i = 1 To nofWell
        Cells(32 + (i - 1), "h").value = "W-" & CStr(i)
        Cells(32 + (i - 1), "i").value = shultz(i)
        Cells(32 + (i - 1), "j").value = webber(i)
        Cells(32 + (i - 1), "k").value = jcob(i)
    
        Cells(32 + (i - 1), "l").value = Round((shultz(i) + webber(i) + jcob(i)) / 3, 1)
        Cells(32 + (i - 1), "m").value = Application.WorksheetFunction.max(shultz(i), webber(i), jcob(i))
    Next i

End Sub



Sub WriteData36_TS_Analysis(T1 As Variant, T2 As Variant, TA As Variant, s2 As Variant, nofWell As Variant)
    Dim i As Integer
    
    With Range("E32:F67")
        .value = ""
    End With
        
    For i = 1 To nofWell
        Cells(32 + (i - 1) * 3, "C").value = "W-" & CStr(i)
    
        Cells(32 + (i - 1) * 3, "E").value = T1(i)
        Cells(33 + (i - 1) * 3, "E").value = T2(i)
        Cells(34 + (i - 1) * 3, "E").value = TA(i)
        
        Cells(32 + (i - 1) * 3, "F").value = s2(i)
        Cells(34 + (i - 1) * 3, "F").value = s2(i)
    Next i
End Sub



Sub WriteData37_RadiusOfInfluence(T2 As Variant, K As Variant, s2 As Variant, time_ As Variant, deltah As Variant, daeSoo As Variant, nofWell As Variant)

    Dim i As Integer
       
    With Range("E22:P27")
        .value = ""
    End With
    
    For i = 1 To nofWell
        Cells(21, Chr(Asc("D") + i)).value = "W-" & CStr(i)
        Cells(22, Chr(Asc("D") + i)).value = T2(i)
        Cells(23, Chr(Asc("D") + i)).value = K(i)
        Cells(24, Chr(Asc("D") + i)).value = s2(i)
        Cells(25, Chr(Asc("D") + i)).value = time_(i)
        Cells(26, Chr(Asc("D") + i)).value = deltah(i)
        Cells(27, Chr(Asc("D") + i)).value = daeSoo(i)
    Next i

End Sub


Sub WriteWellData(Q As Variant, Natural As Variant, Stable As Variant, Recover As Variant, Radius As Variant, deltas As Variant, daeSoo As Variant, T1 As Variant, S1 As Variant, ByVal nofWell As Integer)
    
    Dim i As Integer
    
    For i = 1 To nofWell
        Range("D" & (i + 2)).value = "W-" & i
        
        Range("e" & (i + 2)).value = Q(i)
        Range("l" & (i + 2)).value = Q(i)
        
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


