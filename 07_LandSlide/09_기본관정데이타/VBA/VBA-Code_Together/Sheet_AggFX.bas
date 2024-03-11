
Private Sub CommandButton1_Click()
'Hide Aggregate

    Sheets("YangSoo").Visible = False
    Sheets("Well").Select
End Sub

Private Sub CommandButton2_Click()
'Collect Data
    
    Call GetBaseDataFromYangSoo(999, False)
End Sub

Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub



Private Sub CommandButton4_Click()
'single well import

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

Call GetBaseDataFromYangSoo(singleWell, True)

End Sub


Sub GetBaseDataFromYangSoo(ByVal singleWell As Integer, ByVal isSingleWellImport As Boolean)
' isSingleWellImport = True ---> SingleWell Import
' isSingleWellImport = False ---> AllWell Import
'
' SingleWell --> ImportWell Number
' 999 & False --> 모든관정을 임포트
'
    Dim fName As String
    Dim nofwell, i As Integer
    Dim rngString As String
        
    Dim natural, stable, recover, Sw, delta_h, radius, Rw, well_depth As Double
    Dim casing, Q, delta_s, hp, daeSoo As Double
    Dim T1, T2, TA, S1, S2, T0, S0, ER_MODE, ER1, ER2, ER3 As Double
    Dim K, time_, shultze, webber, jacob, skin, er As Double
    Dim qh, qg, q1, sd1, sd2, C, B, ratio As Double
    
    Dim wb As Workbook
    Dim wsInput As Worksheet
    Dim wsSkinFactor As Worksheet
    Dim wsSafeYield As Worksheet
    
    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
    
    If isSingleWellImport Then
        rngString = "A" & (singleWell + 5 - 1) & ":AN" & (singleWell + 5 - 1)
        Call EraseCellData(rngString)
    Else
        rngString = "A5:AN37"
        Call EraseCellData(rngString)
    End If
        
    
    For i = 1 To nofwell
        
        If Not isSingleWellImport Or (isSingleWellImport And i = singleWell) Then
            GoTo SINGLE_ITERATION
        Else
            GoTo NEXT_ITERATION
        End If
        
SINGLE_ITERATION:
        
        fName = "A" & CStr(i) & "_ge_OriginalSaveFile.xlsm"
        If Not IsWorkBookOpen(fName) Then
            MsgBox "Please open the yangsoo data ! " & fName
            Exit Sub
        End If
        
        Set wb = Workbooks(fName)
        Set wsInput = wb.Worksheets("Input")
        Set wsSkinFactor = wb.Worksheets("SkinFactor")
        Set wsSafeYield = wb.Worksheets("SafeYield")
        
        
        Q = wsInput.Range("m51").value
        hp = wsInput.Range("i48").value
        
        natural = wsInput.Range("m48").value
        stable = wsInput.Range("m49").value
        radius = wsInput.Range("m44").value
        Rw = radius / 2000
        
        well_depth = wsInput.Range("m45").value
        casing = wsInput.Range("i52").value
        
        
        C = wsInput.Range("A31").value
        B = wsInput.Range("B31").value
        
        
        
        recover = wsSkinFactor.Range("c10").value
        Sw = stable - recover
        
        delta_h = wsSkinFactor.Range("b16").value
        delta_s = wsSkinFactor.Range("b4").value
        
        daeSoo = wsSkinFactor.Range("c16").value
        
        '----------------------------------------------------------------------------------
        
        T0 = wsSkinFactor.Range("d4").value
        S0 = wsSkinFactor.Range("f4").value
        ER_MODE = wsSkinFactor.Range("h10").value
        
        T1 = wsSkinFactor.Range("d5").value
        T2 = wsSkinFactor.Range("h13").value
        TA = (T1 + T2) / 2
        
        S1 = wsSkinFactor.Range("e10").value
        S2 = wsSkinFactor.Range("i16").value
        
        K = wsSkinFactor.Range("e16").value
        time_ = wsSkinFactor.Range("h16").value
        
        shultze = wsSkinFactor.Range("c13").value
        webber = wsSkinFactor.Range("c18").value
        jacob = wsSkinFactor.Range("c23").value
        
        skin = wsSkinFactor.Range("g6").value
        er = wsSkinFactor.Range("c8").value
        
        
        ' 경험식, 1번, 2번, 3번의 유효우물반경
        ER1 = wsSkinFactor.Range("K8").value
        ER2 = wsSkinFactor.Range("K9").value
        ER3 = wsSkinFactor.Range("K10").value
        
        '----------------------------------------------------------------------------------
        
        qh = wsSafeYield.Range("b13").value
        qg = wsSafeYield.Range("b7").value
        
        sd1 = wsSafeYield.Range("b3").value
        sd2 = wsSafeYield.Range("b4").value
        q1 = wsSafeYield.Range("b2").value
        
        ratio = wsSafeYield.Range("b11").value
        
        '*****************************************************************************************
        
        Cells(4 + i, "a").value = "W-" & i
        Cells(4 + i, "b").value = natural
        Cells(4 + i, "c").value = stable
        
        Cells(4 + i, "d").value = recover
        Cells(4 + i, "d").NumberFormat = "0.00"
        
        Cells(4 + i, "e").value = Sw
        Cells(4 + i, "e").NumberFormat = "0.00"
        
        Cells(4 + i, "f").value = delta_h
        Cells(4 + i, "f").NumberFormat = "0.00"
        
        Cells(4 + i, "g").value = radius
        Cells(4 + i, "h").value = Rw
        Cells(4 + i, "i").value = well_depth
        Cells(4 + i, "j").value = casing
        Cells(4 + i, "k").value = Q
        
        Cells(4 + i, "l").value = delta_s
        Cells(4 + i, "l").NumberFormat = "0.00"
        
        Cells(4 + i, "m").value = hp
        Cells(4 + i, "n").value = daeSoo
        
        Cells(4 + i, "o").value = T1
        Cells(4 + i, "o").NumberFormat = "0.0000"
         
        Cells(4 + i, "p").value = T2
        Cells(4 + i, "p").NumberFormat = "0.0000"
         
        Cells(4 + i, "q").value = TA
        Cells(4 + i, "q").NumberFormat = "0.0000"
        
        Cells(4 + i, "r").value = S1
        
        Cells(4 + i, "s").value = S2
        Cells(4 + i, "s").NumberFormat = "0.0000000"
        
        Cells(4 + i, "t").value = K
        Cells(4 + i, "t").NumberFormat = "0.0000"
        
        Cells(4 + i, "u").value = time_
        
        Cells(4 + i, "v").value = shultze
        Cells(4 + i, "v").NumberFormat = "0.0"
        
        Cells(4 + i, "w").value = webber
        Cells(4 + i, "w").NumberFormat = "0.0"
        
        Cells(4 + i, "x").value = jacob
        Cells(4 + i, "x").NumberFormat = "0.0"
        
        
        
        Cells(4 + i, "y").value = Format(skin, "0.0000")
        
        Cells(4 + i, "z").value = er
        Cells(4 + i, "z").NumberFormat = "0.0000"
        
        Cells(4 + i, "aa").value = Format(qh, "0.")
        Cells(4 + i, "ab").value = Format(qg, "0.00")
        Cells(4 + i, "ac").value = Format(q1, "0.")
        
        Cells(4 + i, "ad").value = Format(sd1, "0.00")
        Cells(4 + i, "ae").value = Format(sd2, "0.00")
        
        Cells(4 + i, "af").value = C
        Cells(4 + i, "ag").value = B
        
        Cells(4 + i, "ah").value = ratio
        Cells(4 + i, "ah").NumberFormat = "0.0%"
        
        
        ' 2023/09/22 새로 추가한 값들 ...
        Cells(4 + i, "AI").value = Format(T0, "0.0000")
        Cells(4 + i, "AJ").value = Format(S0, "0.0000")
        Cells(4 + i, "AK").value = ER_MODE
        
        Cells(4 + i, "AL").value = Format(ER1, "0.0000")
        Cells(4 + i, "AM").value = Format(ER2, "0.0000")
        Cells(4 + i, "AN").value = Format(ER3, "0.0000")
        
NEXT_ITERATION:

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


'0 : skin factor, cell, C8
'1 : Re1,         cell, E8
'2 : Re2,         cell, H8
'3 : Re3,         cell, G10

Function DetermineEffectiveRadius(ERMode As String) As Integer
    Dim er, r As String
    
    er = ERMode
    'MsgBox er
    r = Mid(er, 5, 1)
    
    If r = "F" Then
        DetermineEffectiveRadius = erRE0
    Else
        DetermineEffectiveRadius = val(r)
    End If
End Function



Private Sub CommandButton3_Click()
' Formula Button

Dim formula1, formula2 As String
Dim nofwell As String
Dim i As Integer
Dim T, Q, radius, skin, er As Double
Dim T0, S0 As Double
Dim ERMode As String
Dim ER1, ER2, ER3, B, S1 As Double

' Save array to a file
Dim filePath As String
Dim FileNum As Integer



nofwell = GetNumberOfWell()
Sheets("YangSoo").Select


filePath = ThisWorkbook.Path & "\" & ActiveSheet.name & ".csv"
FileNum = FreeFile

Open filePath For Output As FileNum

    
Call MyDebug("Formula SkinFactor ... ", True)

Debug.Print "************************************************************************************************************************************************************************************************"
Print #FileNum, "************************************************************************************************************************************************************************************************"

For i = 1 To nofwell

    T = Format(Cells(4 + i, "o").value, "0.0000")
    Q = Cells(4 + i, "k").value
    
    T0 = Format(Cells(4 + i, "AI").value, "0.0000")
    S0 = Format(Cells(4 + i, "AJ").value, "0.0000")
    S1 = Cells(4 + i, "R").value
            
    delta_s = Format(Cells(4 + i, "l").value, "0.00")
    radius = Format(Cells(4 + i, "h").value, "0.000")
    skin = Cells(4 + i, "y").value
    er = Cells(4 + i, "z").value
    
    
    B = Format(Cells(4 + i, "AG").value, "0.0000")
    ER1 = Cells(4 + i, "AL").value
    ER2 = Cells(4 + i, "AM").value
    ER3 = Cells(4 + i, "AN").value
    
    
    Select Case DetermineEffectiveRadius(Cells(4 + i, "AK").value)
    ' 경험식 1번
    Case erRE1
        formula1 = "W-" & i & "호공~~r _{e-" & i & "} `=~ sqrt {{2.25 TIMES  " & T0 & " TIMES  0.0833333} over {" & S0 & " TIMES  10 ^{5.46 TIMES  " & T & " TIMES  " & B & "}}} `=~" & ER1 & "m"
        formula2 = "erRE1, 경험식 1번"
        
    ' 경험식 2번
    Case erRE2
        formula1 = "W-" & i & "호공~~r _{e-" & i & "} `=~ sqrt {{2.25 TIMES  " & T0 & " TIMES  0.0833333} over {" & S0 & " TIMES  10 ^{4 pi TIMES " & T & " TIMES  " & B & "}}} `=~" & ER2 & "m"
        formula2 = "erRE2, 경험식 2번"
    ' 경험식 3번
    Case erRE3
        formula1 = "W-" & i & "호공~~r _{e-" & i & "} `=~" & radius & " TIMES  sqrt {{" & S1 & "} over {" & S0 & "}} `=~" & ER3 & "m"
        formula2 = "erRE3, 경험식 3번"
        
    Case Else
        ' 스킨계수
        formula1 = "W-" & i & "호공~~ sigma  _{w-" & i & "} = {2 pi  TIMES  " & T & " TIMES  " & delta_s & " } over {" & Q & "} -1.15 TIMES  log {2.25 TIMES  " & T & " TIMES  (1/1440)} over {0.0005 TIMES  (" & radius & " TIMES  " & radius & ")} =`" & skin
        ' 유효우물반경
        formula2 = "W-" & i & "호공~~r _{e-" & i & "} `=~r _{w} e ^{- sigma  _{w-" & i & "}} =" & radius & " TIMES e ^{-(" & skin & ")} =" & er & "m"
    End Select
    
    Debug.Print formula1
    Debug.Print formula2
    Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    
    Print #FileNum, formula1
    Print #FileNum, formula2
    Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    

Next i

Debug.Print "************************************************************************************************************************************************************************************************"
Print #FileNum, "************************************************************************************************************************************************************************************************"


Call FormulaChwiSoo(FileNum)
' 3-7, 적정취수량

Call FormulaRadiusOfInfluence(FileNum)
' 양수영향반경
End Sub


Sub FormulaChwiSoo(FileNum As Integer)
' 3-7, 적정취수량

    Dim formula As String
    Dim nofwell As String
    Dim i As Integer
    Dim q1, S1, S2, res As Double
    
    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
        
        
    For i = 1 To 3
        Debug.Print " "
        Print #FileNum, " "
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
    For i = 1 To nofwell
        q1 = Cells(4 + i, "ac").value
 
        S1 = Format(Cells(4 + i, "ad").value, "0.00")
        S2 = Format(Cells(4 + i, "ae").value, "0.00")
        
        res = Cells(4 + i, "ab").value
        formula = "W-" & i & "호공~~Q _{ & 2} `＝" & q1 & "` TIMES  `(` {" & S2 & "} over {" & S1 & "} `) ^{2/3} `＝" & res & "㎥/일"
        
        Debug.Print formula
        Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
        Print #FileNum, formula
        Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
    
End Sub


Sub FormulaRadiusOfInfluence(FileNum As Integer)
' 양수영향반경

    Dim formula1, formula2, formula3 As String
    ' 슐츠, 웨버, 제이콥
    
    Dim nofwell As String
    Dim i As Integer
    Dim shultze, webber, jacob, T, K, S, time_, delta_h As String
    
    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
        
        
    For i = 1 To 3
        Debug.Print " "
        Print #FileNum, " "
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "************************************************************************************************************************************************************************************************"
    
    
    For i = 1 To nofwell
        shultze = CStr(Cells(4 + i, "v").value)
        webber = CStr(Cells(4 + i, "w").value)
        jacob = CStr(Cells(4 + i, "x").value)
        
        T = CStr(Format(Cells(4 + i, "q").value, "0.0000"))
        S = CStr(Format(Cells(4 + i, "s").value, "0.0000000"))
        K = CStr(Format(Cells(4 + i, "t").value, "0.0000"))
    
        delta_h = CStr(Cells(4 + i, "f").value)
        time_ = CStr(Format(Cells(4 + i, "u").value, "0.0000"))
        
        
        ' Cells(4 + i, "y").value = Format(skin(i), "0.0000")
        
        formula1 = "W-" & i & "호공~~R _{W-" & i & "} ``=`` sqrt {6 TIMES  " & delta_h & " TIMES  " & K & " TIMES  " & time_ & "/" & S & "} ``=~" & shultze & "m"
        formula2 = "W-" & i & "호공~~R _{W-" & i & "} ``=3`` sqrt {" & delta_h & " TIMES " & K & " TIMES " & time_ & "/" & S & "} `=`" & webber & "`m"
        formula3 = "W-" & i & "호공~~r _{0(W-" & i & ")} `=~ sqrt {{2.25 TIMES  " & T & " TIMES  " & time_ & "} over {" & S & "}} `=~" & jacob & "m"
        
        Debug.Print ""
        Debug.Print formula1
        Debug.Print formula2
        Debug.Print formula3
        Debug.Print ""
        
        Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
        
         Print #FileNum, " "
         Print #FileNum, formula1
         Print #FileNum, formula2
         Print #FileNum, formula3
         Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
        
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    Print #FileNum, "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    
End Sub






