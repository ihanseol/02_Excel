Attribute VB_Name = "water_q"
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

Sub initialize()
        
    '전라남도, 목포시, 2020 환경부 지하수업무수행지침
    SS(svGAJUNG, 1) = 0.173
    SS(svGAJUNG, 2) = 0.21
    SS_CITY = 2.44
    
    SS(svILBAN, 1) = 2.119
    SS(svILBAN, 2) = 0.021
    
    SS(svSCHOOL, 1) = 7.986
    SS(svSCHOOL, 2) = 0.005
    
    SS(svGONGDONG, 1) = 7.13
    SS(svGONGDONG, 2) = 0.001
    
    SS(svMAEUL, 1) = 6.463
    SS(svMAEUL, 2) = 0.178
    
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


Private Function lastRowByKey(cell As String) As Long

    lastRowByKey = Range(cell).End(xlDown).Row

End Function


' 물량계산
Sub ComputeQ()
    Dim i As Integer
    Dim lastRow As Long

    Call initialize
    
    Sheets("ss").Activate
    lastRow = lastRowByKey("A1")
    
    For i = 2 To lastRow
        Cells(i, "L").Value = ss_water(Range("I" & CStr(i)).Value, Range("K" & CStr(i)).Value, 100)
    Next i
    
    Sheets("aa").Activate
    lastRow = lastRowByKey("A1")
    
    For i = 2 To lastRow
        Cells(i, "L").Value = aa_water(Range("I" & CStr(i)).Value, Range("K" & CStr(i)).Value, 100)
    Next i
End Sub


Function ss_water(ByVal qhp As Integer, ByVal strPurpose As String, Optional ByVal npopulation As Integer = 60) As Double
    '지열 냉난방
    If CheckSubstring(strPurpose, "냉") Then
        ss_water = qhp * 0.01
        Exit Function
    End If
    
    ' 일반용
    If CheckSubstring(strPurpose, "일") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    
    ' 가정용
    If CheckSubstring(strPurpose, "가") Then
        ss_water = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
    
    ' 기타
    If CheckSubstring(strPurpose, "기") Then
        ss_water = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
    
    ' 농생활겸용
    If CheckSubstring(strPurpose, "농") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    ' 청소용
    If CheckSubstring(strPurpose, "청") Then
        ss_water = Round(SS(svGAJUNG, 1) + SS_CITY * SS(svGAJUNG, 2), 2)
        Exit Function
    End If
    
    '간이상수도
    If CheckSubstring(strPurpose, "상") Then
        ss_water = Round(SS(svMAEUL, 1) + npopulation * SS(svMAEUL, 2), 2)
        Exit Function
    End If
    
    ' 공사용
    If CheckSubstring(strPurpose, "공사") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    ' 공동주택용
    If CheckSubstring(strPurpose, "공동") Then
        ss_water = Round(SS(svGONGDONG, 1) + npopulation * SS(svGONGDONG, 2), 2)
        Exit Function
    End If
        
    ' 민방위용
    If CheckSubstring(strPurpose, "민방") Then
        ss_water = Round(SS(svILBAN, 1) + qhp * SS(svILBAN, 2), 2)
        Exit Function
    End If
    
    ' 학교용
    If CheckSubstring(strPurpose, "학교") Then
        ss_water = Round(SS(svSCHOOL, 1) + npopulation * SS(svSCHOOL, 2), 2)
        Exit Function
    End If
    
   ss_water = 900
      
End Function




Function aa_water(qhp As Integer, strPurpose As String, Optional ByVal nhead As Integer = 30) As Double

    'nhead - 축산업의 두수 ....

    ' 전작용
    If CheckSubstring(strPurpose, "전") Then
        aa_water = Round(AA(avJEONJAK, 1) + qhp * AA(avJEONJAK, 2), 2)
        Exit Function
    End If
    
    ' 답작용
    If CheckSubstring(strPurpose, "답") Then
        aa_water = Round(AA(avDAPJAK, 1) + qhp * AA(avDAPJAK, 2), 2)
        Exit Function
    End If
    
    
    ' 원예용
    If CheckSubstring(strPurpose, "원") Then
        aa_water = Round(AA(avWONYE, 1) + qhp * AA(avWONYE, 2), 2)
        Exit Function
    End If
    
    ' 농생활겸용
    If CheckSubstring(strPurpose, "농") Then
        aa_water = Round(AA(avJEONJAK, 1) + qhp * AA(avJEONJAK, 2), 2)
        Exit Function
    End If
    
    ' 양계장용
    If CheckSubstring(strPurpose, "양") Then
        aa_water = Round(AA(avDAPJAK, 1) + qhp * AA(avDAPJAK, 2), 2)
        Exit Function
    End If
    
    '축산용
    If CheckSubstring(strPurpose, "축") Then
        aa_water = Round(AA(avCOW, 1) + nhead * AA(avCOW, 2), 2)
        Exit Function
    End If
    
    ' 기타
    If CheckSubstring(strPurpose, "기타") Then
        aa_water = Round(AA(avDAPJAK, 1) + nhead * AA(avDAPJAK, 2), 2)
        Exit Function
    End If
    
    
   aa_water = 900
      
End Function










