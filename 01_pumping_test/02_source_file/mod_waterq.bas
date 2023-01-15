Attribute VB_Name = "mod_waterq"
Public SS(1 To 5, 1 To 2) As Double
Public AA(1 To 6, 1 To 2) As Double

Public SS_CITY As Double

Public Enum SS_VALUE

    svGAJUNG = 1
    svILBAN = 2
    svSCHOOL = 3
    svGONGDONG = 4
    svMAEUL = 5

End Enum

Public Enum AA_VALUE
    
    avJEONJAK = 1
    avDAPJAK = 2
    avWONYE = 3
    avCOW = 4
    avPIG = 5
    avCHICKEN = 6
    
End Enum


Sub init_nonsan()

   
    SS(svGAJUNG, 1) = 0.173
    SS(svGAJUNG, 2) = 0.21
    SS_CITY = 2.63
    
    SS(svILBAN, 1) = 3.154
    SS(svILBAN, 2) = 0.023
    
    SS(svSCHOOL, 1) = 7.986
    SS(svSCHOOL, 2) = 0.005
    
    SS(svGONGDONG, 1) = 0.173
    SS(svGONGDONG, 2) = 0.21
    
    SS(svMAEUL, 1) = 7.13
    SS(svMAEUL, 2) = 0.001
    
'----------------------------------------

    AA(avJEONJAK, 1) = 5.66
    AA(avJEONJAK, 2) = 0.014
    
    AA(avDAPJAK, 1) = 1.98
    AA(avDAPJAK, 2) = 0.044
    
    AA(avWONYE, 1) = 2.789
    AA(avWONYE, 2) = 0.011
    
    AA(avCOW, 1) = 3.48
    AA(avCOW, 2) = 0.009
    
    AA(avPIG, 1) = 4.719
    AA(avPIG, 2) = 0.001
    
    AA(avCHICKEN, 1) = 5.492
    AA(avCHICKEN, 2) = 0.041
    
End Sub


'충청남도 연기군
Sub init_sejong()

   
    SS(svGAJUNG, 1) = 0.173
    SS(svGAJUNG, 2) = 0.21
    SS_CITY = 2.75

    SS(svILBAN, 1) = 3.521
    SS(svILBAN, 2) = 0.011
    
    SS(svSCHOOL, 1) = 11.687
    SS(svSCHOOL, 2) = 0.007
    
    SS(svGONGDONG, 1) = 0.265
    SS(svGONGDONG, 2) = 0.181
    
    SS(svMAEUL, 1) = 7.287
    SS(svMAEUL, 2) = 0.001
    
'----------------------------------------

    AA(avJEONJAK, 1) = 5.66
    AA(avJEONJAK, 2) = 0.014
    
    AA(avDAPJAK, 1) = 1.98
    AA(avDAPJAK, 2) = 0.044
    
    AA(avWONYE, 1) = 2.789
    AA(avWONYE, 2) = 0.011
    
    AA(avCOW, 1) = 3.48
    AA(avCOW, 2) = 0.009
    
    AA(avPIG, 1) = 4.719
    AA(avPIG, 2) = 0.001
    
    AA(avCHICKEN, 1) = 5.492
    AA(avCHICKEN, 2) = 0.041

End Sub

Sub init_daejeon()

   
    SS(svGAJUNG, 1) = 0.173
    SS(svGAJUNG, 2) = 0.21
    SS_CITY = 2.73

    SS(svILBAN, 1) = 3.154
    SS(svILBAN, 2) = 0.023
    
    SS(svSCHOOL, 1) = 7.986
    SS(svSCHOOL, 2) = 0.005
    
    SS(svGONGDONG, 1) = 0.173
    SS(svGONGDONG, 2) = 0.21
    
    SS(svMAEUL, 1) = 7.13
    SS(svMAEUL, 2) = 0.001
    
'----------------------------------------

    AA(avJEONJAK, 1) = 5.66
    AA(avJEONJAK, 2) = 0.014
    
    AA(avDAPJAK, 1) = 1.98
    AA(avDAPJAK, 2) = 0.044
    
    AA(avWONYE, 1) = 2.789
    AA(avWONYE, 2) = 0.011
    
    AA(avCOW, 1) = 3.48
    AA(avCOW, 2) = 0.009
    
    AA(avPIG, 1) = 4.719
    AA(avPIG, 2) = 0.001
    
    AA(avCHICKEN, 1) = 5.492
    AA(avCHICKEN, 2) = 0.041

End Sub

Sub init_boryoung()

   
    SS(svGAJUNG, 1) = 0.173
    SS(svGAJUNG, 2) = 0.21
    SS_CITY = 2.52
    
    SS(svILBAN, 1) = 3.154
    SS(svILBAN, 2) = 0.023
    
    SS(svSCHOOL, 1) = 7.986
    SS(svSCHOOL, 2) = 0.005
    
    SS(svGONGDONG, 1) = 0.173
    SS(svGONGDONG, 2) = 0.21
    
    SS(svMAEUL, 1) = 7.13
    SS(svMAEUL, 2) = 0.001
    
'----------------------------------------

    AA(avJEONJAK, 1) = 6.964
    AA(avJEONJAK, 2) = 0.013
    
    AA(avDAPJAK, 1) = 2.089
    AA(avDAPJAK, 2) = 0.043
    
    AA(avWONYE, 1) = 2.789
    AA(avWONYE, 2) = 0.011
    
    AA(avCOW, 1) = 3.48
    AA(avCOW, 2) = 0.009
    
    AA(avPIG, 1) = 4.719
    AA(avPIG, 2) = 0.001
    
    AA(avCHICKEN, 1) = 5.492
    AA(avCHICKEN, 2) = 0.041
    
End Sub


Sub initialize()
        
       'Call init_nonsan
       'Call init_daejeon
       'Call init_boryoung
       Call init_sejong
       
End Sub

Function instr2(ByVal strPurpose, ByVal sstr As String) As Boolean

    Dim mypos As Integer
    
    mypos = InStr(1, strPurpose, sstr)
    If (mypos <> 0) Then
        instr2 = True
        Exit Function
    End If
    
    instr2 = False

End Function

Function ss_water(ByVal qhp As Integer, ByVal strPurpose As String, Optional ByVal npopulation As Integer = 60) As Double

    Dim mypos As Integer


    mypos = InStr(1, strPurpose, "일") '일반용
    If (mypos <> 0) Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "가") '가정용
    If (mypos <> 0) Then
        ss_water = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "기") '기타
    If (mypos <> 0) Then
        ss_water = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "농") '농생활겸용
    If (mypos <> 0) Then
        ss_water = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "청") '청소용
    If (mypos <> 0) Then
        ss_water = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "학") '학교용
    If (mypos <> 0) Then
        ss_water = Round(SS(svSCHOOL, 1) + 3 * npopulation * SS(svSCHOOL, 2), 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "상") '간이상수도
    If (mypos <> 0) Then
        ss_water = Round(SS(svMAEUL, 1) + npopulation * SS(svMAEUL, 2), 2)
        Exit Function
    End If
    
    
   ss_water = 900
      
End Function

Function ss_water1(ByVal qhp As Integer, ByVal strPurpose As String, Optional ByVal npopulation As Integer = 60) As Double

    Dim mypos As Integer


    If (instr2(strPurpose, "일")) Then '일반용
        ss_water1 = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    
    If (instr2(strPurpose, "가") Or instr2(strPurpose, "기") Or instr2(strPurpose, "농") Or instr2(strPurpose, "청")) Then  '가정용
        ss_water1 = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
        
    
    If (instr2(strPurpose, "학")) Then '학교용
        ss_water1 = Round(SS(svSCHOOL, 1) + 3 * npopulation * SS(svSCHOOL, 2), 2)
        Exit Function
    End If
    
    If (instr2(strPurpose, "상") Or instr2(strPurpose, "마")) Then  '학교용
        ss_water1 = Round(SS(svMAEUL, 1) + npopulation * SS(svMAEUL, 2), 2)
        Exit Function
    End If
    
    If (instr2(strPurpose, "공")) Then '공동주택용
        ss_water1 = Round(SS(svGONGDONG, 1) + npopulation * SS(svGONGDONG, 2), 2)
        Exit Function
    End If
        
    
   ss_water1 = 900
      
End Function




Function aa_water(qhp As Integer, strPurpose As String, Optional ByVal nhead As Integer = 30) As Double

    'nhead - 축산업의 두수 ....


    Dim mypos As Integer


    mypos = InStr(1, strPurpose, "전") '전작용
    If (mypos <> 0) Then
        aa_water = Round(AA(avJEONJAK, 1) + qhp * AA(avJEONJAK, 2), 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "답") '답작용
    If (mypos <> 0) Then
        aa_water = Round(AA(avDAPJAK, 1) + qhp * AA(avDAPJAK, 2), 2)
        Exit Function
    End If
    
    mypos = InStr(1, strPurpose, "원") '원예용
    If (mypos <> 0) Then
        aa_water = Round(AA(avWONYE, 1) + qhp * AA(avWONYE, 2), 2)
        Exit Function
    End If
    
    '농생활겸용
    mypos = InStr(1, strPurpose, "농")
    If (mypos <> 0) Then
        aa_water = Round(AA(avJEONJAK, 1) + qhp * AA(avJEONJAK, 2), 2)
        Exit Function
    End If
    
    '양어장용
    mypos = InStr(1, strPurpose, "양")
    If (mypos <> 0) Then
        aa_water = Round(AA(avDAPJAK, 1) + qhp * AA(avDAPJAK, 2), 2)
        Exit Function
    End If
    
    
    mypos = InStr(1, strPurpose, "축") '축산업
    If (mypos <> 0) Then
        aa_water = Round(AA(avCOW, 1) + nhead * AA(avCOW, 2), 2)
        Exit Function
    End If
    
   aa_water = 900
      
End Function


Function aa_water1(qhp As Integer, strPurpose As String, Optional ByVal nhead As Integer = 30) As Double

    'nhead - 축산업의 두수 ....


    If (instr2(strPurpose, "전") Or instr2(strPurpose, "농")) Then '전작용, 농생활겸용
        aa_water1 = Round(AA(avJEONJAK, 1) + qhp * AA(avJEONJAK, 2), 2)
        Exit Function
    End If
        
    If (instr2(strPurpose, "답") Or instr2(strPurpose, "양")) Then '답작용, 양어장용
        aa_water1 = Round(AA(avDAPJAK, 1) + qhp * AA(avDAPJAK, 2), 2)
        Exit Function
    End If
 
    If (instr2(strPurpose, "원")) Then '원예용
        aa_water1 = Round(AA(avWONYE, 1) + qhp * AA(avWONYE, 2), 2)
        Exit Function
    End If
 
    If (instr2(strPurpose, "축")) Then '축산업
        aa_water1 = Round(AA(avCOW, 1) + nhead * AA(avCOW, 2), 2)
        Exit Function
    End If
    
   aa_water1 = 900
      
End Function








