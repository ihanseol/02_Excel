VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_AggYangSoo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Sheets("YangSoo").Visible = False
    Sheets("Well").Select
End Sub

Private Sub CommandButton2_Click()
    Call GetBaseDataFromYangSoo
End Sub

Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub


Sub GetBaseDataFromYangSoo()
    Dim fname As String
    Dim nofwell, i As Integer
        
    Dim natural() As Double ' 자연수위, natural depth
    Dim stable() As Double  ' 안정수위, stable depth
    Dim recover() As Double ' 회복수위, recover depth
    Dim Sw() As Double ' 수위회복량 - 안정수위 - 회복수위
    
    Dim delta_h() As Double ' deltah : 수위강하량
    
    Dim radius() As Double ' 공반경
    Dim Rw() As Double      ' 공반경 / 2000
    
    Dim well_depth() As Double     ' 관정심도, well depth
    Dim casing() As Double  ' 케이싱심도
    
    Dim Q() As Double       '취수계획량
    Dim delta_s() As Double
    Dim hp() As Double
    
    Dim daeSoo() As Double  ' 대수층 두께
    
    Dim T1() As Double      ' T1
    Dim T2() As Double      ' T2
    Dim TA() As Double      ' TA - (T1+T2)/2, TAverage
    
    Dim S1() As Double      ' S1
    Dim S2() As Double      ' S2 - 스킨팩터 해석, s값
    
    Dim K() As Double
    Dim time_() As Double   ' 안정수위도달시간
    
    Dim shultze() As Double
    Dim webber() As Double
    Dim jacob() As Double
    
    Dim skin() As Double ' skin factor
    Dim er() As Double   ' effective radius, 유효우물반경
        
    
    Dim qh() As Double ' 한계양수량
    Dim qg() As Double ' 가채수량
    Dim q1() As Double ' Q1
    
    
    Dim sd1() As Double ' 1단계 수위강하량
    Dim sd2() As Double ' 4단계 수위강하량
    
    
    Dim C() As Double
    Dim B() As Double
    
    Dim ratio() As Double
    
    
 ' --------------------------------------------------------------------------------------

    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
    
 ' --------------------------------------------------------------------------------------
    
    ReDim natural(1 To nofwell)
    ReDim stable(1 To nofwell)
    ReDim recover(1 To nofwell)
    ReDim delta_h(1 To nofwell)
    ReDim Sw(1 To nofwell)
    
    
    ReDim radius(1 To nofwell)
    ReDim Rw(1 To nofwell)
    
    ReDim well_depth(1 To nofwell)
    ReDim casing(1 To nofwell)
    
    ReDim Q(1 To nofwell)
    ReDim delta_s(1 To nofwell)
    ReDim hp(1 To nofwell)
    
    ReDim daeSoo(1 To nofwell)
    
    ReDim T1(1 To nofwell)
    ReDim T2(1 To nofwell)
    ReDim TA(1 To nofwell)
    
    ReDim S1(1 To nofwell)
    ReDim S2(1 To nofwell)
    
    ReDim K(1 To nofwell)
    ReDim time_(1 To nofwell)
    
    ReDim shultze(1 To nofwell)
    ReDim webber(1 To nofwell)
    ReDim jacob(1 To nofwell)
    
    ReDim skin(1 To nofwell)
    ReDim er(1 To nofwell)
        
    
    ReDim qh(1 To nofwell)
    ReDim qg(1 To nofwell)
    
    
    ReDim sd1(1 To nofwell)
    ReDim sd2(1 To nofwell)
    ReDim q1(1 To nofwell)
    
    ReDim C(1 To nofwell)
    ReDim B(1 To nofwell)
    
    ReDim ratio(1 To nofwell)
    
    Call EraseCellData("a5:ah19")
    
    For i = 1 To nofwell
    
        fname = "A" & CStr(i) & "_ge_OriginalSaveFile.xlsm"
        If Not IsWorkBookOpen(fname) Then
            MsgBox "Please open the yangsoo data ! " & fname
            Exit Sub
        End If
        
        
        Q(i) = Workbooks(fname).Worksheets("Input").Range("m51").value
        hp(i) = Workbooks(fname).Worksheets("Input").Range("i48").value
        
        natural(i) = Workbooks(fname).Worksheets("Input").Range("m48").value
        stable(i) = Workbooks(fname).Worksheets("Input").Range("m49").value
        radius(i) = Workbooks(fname).Worksheets("Input").Range("m44").value
        Rw(i) = radius(i) / 2000
        
        well_depth(i) = Workbooks(fname).Worksheets("Input").Range("m45").value
        casing(i) = Workbooks(fname).Worksheets("Input").Range("i52").value
        
        
        C(i) = Workbooks(fname).Worksheets("Input").Range("A31").value
        B(i) = Workbooks(fname).Worksheets("Input").Range("B31").value
        
        
        
        recover(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c10").value
        Sw(i) = stable(i) - recover(i)
        
        delta_h(i) = Workbooks(fname).Worksheets("SkinFactor").Range("b16").value
        delta_s(i) = Workbooks(fname).Worksheets("SkinFactor").Range("b4").value
        
        daeSoo(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c16").value
        
        '----------------------------------------------------------------------------------
        
        T1(i) = Workbooks(fname).Worksheets("SkinFactor").Range("d5").value
        T2(i) = Workbooks(fname).Worksheets("SkinFactor").Range("h13").value
        TA(i) = (T1(i) + T2(i)) / 2
        
        S1(i) = Workbooks(fname).Worksheets("SkinFactor").Range("e10").value
        S2(i) = Workbooks(fname).Worksheets("SkinFactor").Range("i16").value
        
        K(i) = Workbooks(fname).Worksheets("SkinFactor").Range("e16").value
        time_(i) = Workbooks(fname).Worksheets("SkinFactor").Range("h16").value
        
        shultze(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c13").value
        webber(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c18").value
        jacob(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c23").value
        
        skin(i) = Workbooks(fname).Worksheets("SkinFactor").Range("g6").value
        er(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c8").value
        
        '----------------------------------------------------------------------------------
        
        qh(i) = Workbooks(fname).Worksheets("SafeYield").Range("b13").value
        qg(i) = Workbooks(fname).Worksheets("SafeYield").Range("b7").value
        
        sd1(i) = Workbooks(fname).Worksheets("SafeYield").Range("b3").value
        sd2(i) = Workbooks(fname).Worksheets("SafeYield").Range("b4").value
        q1(i) = Workbooks(fname).Worksheets("SafeYield").Range("b2").value
        
        ratio(i) = Workbooks(fname).Worksheets("SafeYield").Range("b11").value
        
        '*****************************************************************************************
        
        Cells(4 + i, "a").value = "W-" & CStr(i)
        Cells(4 + i, "b").value = natural(i)
        Cells(4 + i, "c").value = stable(i)
        
        Cells(4 + i, "d").value = recover(i)
        Cells(4 + i, "d").NumberFormat = "0.00"
        
        Cells(4 + i, "e").value = Sw(i)
        Cells(4 + i, "e").NumberFormat = "0.00"
        
        Cells(4 + i, "f").value = delta_h(i)
        Cells(4 + i, "f").NumberFormat = "0.00"
        
        Cells(4 + i, "g").value = radius(i)
        Cells(4 + i, "h").value = Rw(i)
        Cells(4 + i, "i").value = well_depth(i)
        Cells(4 + i, "j").value = casing(i)
        Cells(4 + i, "k").value = Q(i)
        
        Cells(4 + i, "l").value = delta_s(i)
        Cells(4 + i, "l").NumberFormat = "0.00"
        
        Cells(4 + i, "m").value = hp(i)
        Cells(4 + i, "n").value = daeSoo(i)
        
        Cells(4 + i, "o").value = T1(i)
        Cells(4 + i, "o").NumberFormat = "0.0000"
         
        Cells(4 + i, "p").value = T2(i)
        Cells(4 + i, "p").NumberFormat = "0.0000"
         
        Cells(4 + i, "q").value = TA(i)
        Cells(4 + i, "q").NumberFormat = "0.0000"
        
        Cells(4 + i, "r").value = S1(i)
        
        Cells(4 + i, "s").value = S2(i)
        Cells(4 + i, "s").NumberFormat = "0.0000000"
        
        Cells(4 + i, "t").value = K(i)
        Cells(4 + i, "t").NumberFormat = "0.0000"
        
        Cells(4 + i, "u").value = time_(i)
        
        Cells(4 + i, "v").value = shultze(i)
        Cells(4 + i, "v").NumberFormat = "0.0"
        
        Cells(4 + i, "w").value = webber(i)
        Cells(4 + i, "w").NumberFormat = "0.0"
        
        Cells(4 + i, "x").value = jacob(i)
        Cells(4 + i, "x").NumberFormat = "0.0"
        
        
        
        Cells(4 + i, "y").value = Format(skin(i), "0.0000")
        
        Cells(4 + i, "z").value = er(i)
        Cells(4 + i, "z").NumberFormat = "0.0000"
        
        Cells(4 + i, "aa").value = Format(qh(i), "0.")
        Cells(4 + i, "ab").value = Format(qg(i), "0.00")
        Cells(4 + i, "ac").value = Format(q1(i), "0.")
        
        Cells(4 + i, "ad").value = Format(sd1(i), "0.00")
        Cells(4 + i, "ae").value = Format(sd2(i), "0.00")
        
        Cells(4 + i, "af").value = C(i)
        Cells(4 + i, "ag").value = B(i)
        
        Cells(4 + i, "ah").value = ratio(i)
        Cells(4 + i, "ah").NumberFormat = "0.0%"
        
    Next i
End Sub


Public Sub MyDebug(sPrintStr As String, Optional bClear As Boolean = False)
   If bClear = True Then
      Application.SendKeys "^g^{END}", True

      DoEvents '  !!! DoEvents is VERY IMPORTANT here !!!

      Debug.Print String(30, vbCrLf)
   End If

   Debug.Print sPrintStr
End Sub


Private Sub CommandButton3_Click()
    Dim formula1, formula2 As String
    Dim nofwell As String
    Dim i As Integer
    Dim T, Q, radius, skin, er As Double
    
    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
        
        
    Call MyDebug("Formula SkinFactor ... ", True)
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    
    For i = 1 To nofwell
    
        T = Cells(4 + i, "o").value
        Q = Cells(4 + i, "k").value
        delta_s = Format(Cells(4 + i, "l").value, "0.00")
        radius = Format(Cells(4 + i, "h").value, "0.000")
        skin = Cells(4 + i, "y").value
        er = Cells(4 + i, "z").value
        
        formula1 = "W-" & i & "~~ sigma  _{w-" & i & "} = {2 pi  TIMES  " & T & " TIMES  " & delta_s & " } over {" & Q & "} -1.15 TIMES  log {2.25 TIMES  " & T & " TIMES  (1/1440)} over {0.0005 TIMES  (" & radius & " TIMES  " & radius & ")} =`" & skin
        formula2 = "W-" & i & "~~r _{e-" & i & "} `=~r _{w} e ^{- sigma  _{w-" & i & "}} =" & radius & " TIMES e ^{-(" & skin & ")} =" & er & "m"
        Debug.Print formula1
        Debug.Print formula2
        Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    
    Call FormulaChwiSoo
    Call FormulaRadiusOfInfluence
End Sub


Sub FormulaChwiSoo()
    Dim formula As String
    Dim nofwell As String
    Dim i As Integer
    Dim q1, S1, S2, res As Double
    
    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
        
        
    For i = 1 To 7
        Debug.Print " "
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    
    For i = 1 To nofwell
        q1 = Cells(4 + i, "ac").value
 
        S1 = Format(Cells(4 + i, "ad").value, "0.00")
        S2 = Format(Cells(4 + i, "ae").value, "0.00")
        
        res = Cells(4 + i, "ab").value
        formula = "W-" & i & "~~Q _{ & 2} `＝" & q1 & "` TIMES  `(` {" & S2 & "} over {" & S1 & "} `) ^{2/3} `＝" & res & "㎥/일"
        
        Debug.Print formula
        Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
End Sub


Sub FormulaRadiusOfInfluence()

    Dim formula1, formula2, formula3 As String
    ' 슐츠, 웨버, 제이콥
    
    Dim nofwell As String
    Dim i As Integer
    Dim shultze, webber, jacob, T, K, S, time_, delta_h As String
    
    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
        
        
    For i = 1 To 7
        Debug.Print " "
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    
    For i = 1 To nofwell
        shultze = CStr(Cells(4 + i, "v").value)
        webber = CStr(Cells(4 + i, "w").value)
        jacob = CStr(Cells(4 + i, "x").value)
        
        T = CStr(Cells(4 + i, "q").value)
        S = CStr(Cells(4 + i, "s").value)
        K = CStr(Cells(4 + i, "t").value)
    
        delta_h = CStr(Cells(4 + i, "f").value)
        time_ = CStr(Cells(4 + i, "u").value)
        
        formula1 = "W-" & i & "호공~~R _{W-" & i & "} ``=`` sqrt {6 TIMES  " & delta_h & " TIMES  " & K & " TIMES  " & time_ & "/" & S & "} ``=~" & shultze & "m"
        formula2 = "W-" & i & "호공~~R _{W-" & i & "} ``=3`` sqrt {" & delta_h & " TIMES " & K & " TIMES " & time_ & "/" & S & "} `=`" & webber & "`m"
        formula3 = "W-" & i & "호공~~r _{0(W-" & i & ")} `=~ sqrt {{2.25 TIMES  " & T & " TIMES  " & time_ & "} over {" & S & "}} `=~" & jacob & "m"
        
        Debug.Print ""
        Debug.Print formula1
        Debug.Print formula2
        Debug.Print formula3
        Debug.Print ""
        
        Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    
End Sub







