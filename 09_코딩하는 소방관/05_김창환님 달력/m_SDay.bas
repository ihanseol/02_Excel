Attribute VB_Name = "m_SDay"
Option Explicit

Type SpecialDay '명절, 기념일 등
  RealDay As Double   '올해의 날짜
  y As Integer
  M As Integer
  D As Integer
  LuniSolar As Boolean
  LeapMonth As Boolean
  DayName As String
  Holy As Byte
End Type

Public SDay(200) As SpecialDay  '기념일(150개까지+절기)

Public Function GetOtherDay(ByVal Kind As Byte, ByVal cYear As Integer) As Double
  Dim j As Double, pY As Integer, R As Double, B As Boolean, N As Double, M As Double, P As Double
  
  If Kind = 0 Then '한식
    pY = cYear - 1
    j = cJunggi(CDbl(pY), 270, 350, TimeZone)  '전년의 동지 구하기
    j = GetJD0(j) + 0.5
    R = j + 105  '동지 다음날로부터 105일째 날
  
  ElseIf Kind = 1 Then '초복
    j = cJunggi(CDbl(cYear), 90, 170, TimeZone)  '당년의 하지 구하기
    j = GetJD0(j) + 0.5
    N = (j + 49) Mod 60
    M = N Mod 10
    
    If M <= 6 Then j = j + (6 - M) Else j = j + (16 - M)
    R = j + 20
    
  ElseIf Kind = 2 Then '중복
    j = cJunggi(CDbl(cYear), 90, 170, TimeZone)  '당년의 하지 구하기
    j = GetJD0(j) + 0.5
    N = (j + 49) Mod 60
    M = N Mod 10
    
    If M <= 6 Then j = j + (6 - M) Else j = j + (16 - M)
    R = j + 30
    
  ElseIf Kind = 3 Then '말복
    j = cJunggi(CDbl(cYear), 135, 215, TimeZone)   '당년의 입추 구하기
    j = GetJD0(j) + 0.5
    N = (j + 49) Mod 60
    M = N Mod 10
    
    If M <= 6 Then j = j + (6 - M) Else j = j + (16 - M)
    R = j
    
  ElseIf Kind = 4 Then '납향
    pY = cYear - 1
    j = cJunggi(CDbl(pY), 270, 350, TimeZone)  '전년의 동지 구하기
    j = GetJD0(j) + 0.5
    N = (j + 49) Mod 60
    M = N Mod 12
    
    If M <= 7 Then j = j + (7 - M) Else j = j + (M + 6)
    R = j + 24
    
  ElseIf Kind = 5 Then '제석
    B = InvLuniSolarCal(cYear, 1, 1, False, TimeZone, UseMeanSun, UseMeanMoon, UseJinsak, j)      '설날 구하기
    R = j - 1  '설날 전날
    
  ElseIf Kind = 6 Then  '부활절(교회)
    R = FindEaster(cYear)  '부활절
  
  ElseIf Kind = 7 Then
    j = JULIANDAY(CDbl(cYear), 5, 1, 12, 0)
    N = (j + 1) Mod 7
    If N <= 1 Then M = 1 - N Else M = 8 - N
    R = j + M + 14
  End If
  
  GetOtherDay = R
End Function

'이 함수는 창을 열 때 한 번만 수행하면 됨
Sub CalcSpecialDay(ByVal cYear As Integer)
  Dim i As Integer, a As Boolean, N As Integer, M As String, X As Integer, g As Double
  Dim E As Double, S As Double
  
  i = 0: E = JULIANDAY(CDbl(cYear), 12, 31, 12, 0): S = JULIANDAY(CDbl(cYear), 1, 1, 12, 0)
  Call ClearSDay
  Call ReadDayData
  Do While SDay(i).M > 0 And i < 150
    If SDay(i).LuniSolar Then
      If SDay(i).y > -10000 Then
        a = FindTBLInv(CInt(SDay(i).y), CByte(SDay(i).M), CByte(SDay(i).D), SDay(i).LeapMonth, SDay(i).RealDay)
      ElseIf SDay(i).y <= -10000 And SDay(i).y > -15000 Then
        a = True: SDay(i).RealDay = 1
      End If
      
      If a = False Then SDay(i).M = -1 'RealDay = 0  '없는 날이면
    Else
      If SDay(i).y > -10000 Then
        SDay(i).RealDay = JULIANDAY(CDbl(SDay(i).y), CDbl(SDay(i).M), CDbl(SDay(i).D), 12, 0)
      ElseIf SDay(i).y <= -10000 And SDay(i).y > -15000 Then
        SDay(i).RealDay = JULIANDAY(CDbl(cYear), CDbl(SDay(i).M), CDbl(SDay(i).D), 12, 0)
      End If
    End If
    
    i = i + 1
  Loop
  
  '잡절 더하기
  N = i
  For i = 0 To 7
    SDay(N).RealDay = GetOtherDay(i, cYear)
    If SDay(N).RealDay > E Then
      SDay(N).RealDay = GetOtherDay(i, cYear - 1)
    ElseIf SDay(N).RealDay < S Then
      SDay(N).RealDay = GetOtherDay(i, cYear + 1)
    End If
    Select Case i
     Case 0  '한식
       SDay(N).DayName = "한식": SDay(N).Holy = 0: SDay(N).M = 0
     Case 1  '초복
       SDay(N).DayName = "초복": SDay(N).Holy = 0: SDay(N).M = 0
     Case 2 '중복
       SDay(N).DayName = "중복": SDay(N).Holy = 0: SDay(N).M = 0
     Case 3 '말복
       SDay(N).DayName = "말복": SDay(N).Holy = 0: SDay(N).M = 0
     Case 4 '납향
       SDay(N).DayName = "납향": SDay(N).Holy = 0: SDay(N).M = 0
     Case 5 '제석
       SDay(N).DayName = "제석": SDay(N).Holy = 1: SDay(N).M = 0
     Case 6 '부활절
       SDay(N).DayName = "부활절": SDay(N).Holy = 0: SDay(N).M = 0
     Case 7
       SDay(N).DayName = "성년의 날": SDay(N).Holy = 0: SDay(N).M = 0
    End Select
    N = N + 1
  Next i
  
  '24절기 더하기
  M = "소한대한입춘우수경칩춘분청명곡우입하소만망종하지소서대서입추처서백로추분한로상강입동소설대설동지"
  For i = 0 To 23
    SDay(N).DayName = Mid$(M, 1 + 2 * i, 2)
    SDay(N).Holy = 0
    SDay(N).M = 0
    SDay(N).RealDay = GetJD0(cJunggi(CDbl(cYear), Rev(285 + i * 15), 5 + i * 15, TimeZone)) + 0.5
    If SDay(N).RealDay > E Then
      SDay(N).RealDay = GetJD0(cJunggi(CDbl(cYear - 1), Rev(285 + i * 15), 5 + i * 15, TimeZone)) + 0.5
    ElseIf SDay(N).RealDay < S Then
      SDay(N).RealDay = GetJD0(cJunggi(CDbl(cYear + 1), Rev(285 + i * 15), 5 + i * 15, TimeZone)) + 0.5
    End If
    N = N + 1
  Next i
  
  For i = 0 To N - 1
    If SDay(i).y = -15000 Then
      X = SDay(i).M
      g = CDbl(SDay(i).D)
      If X > 2 Then SDay(i).RealDay = GetJD0(SDay(X - 2).RealDay + g) + 0.5 Else SDay(i).M = -1
    End If
  Next i
End Sub

Function FindSDay(ByVal JD As Double, h As Byte) As String
  Dim i As Integer, R As String, k As Byte
  
  R = "": k = 0
  Do While SDay(i).M > -5 And i <= 200 'SDay(i).RealDay > 0 And i <= 200
    If (SDay(i).M > -1 And JD = SDay(i).RealDay) And (SDay(i).LuniSolar = False Or SDay(i).y > -10000) Then
      If R <> "" And SDay(i).DayName <> "" Then R = R & ", "
      R = R & SDay(i).DayName
      If k = 1 And SDay(i).Holy = 2 Then
        k = 3
      ElseIf k = 2 And SDay(i).Holy = 1 Then
        k = 3
      ElseIf k = 3 Then
        k = 3  'k=3이면 Holy 값이 무엇이든 상관없이 3 반환
      ElseIf k > 0 And SDay(i).Holy = 0 Then
        'k = k  '이 경우 아무 처리 안함
      Else
        k = SDay(i).Holy
      End If
    End If
    
    i = i + 1
  Loop
  
  h = k
  FindSDay = R
End Function

Function FindSDayL(ByVal STR As String, ByVal lm As Integer, ByVal LD As Integer, ByVal Leap As Boolean, h As Byte) As String
  Dim i As Integer, R As String, k As Byte
  
  R = STR: k = h
  Do While SDay(i).M > -5 And i <= 200 'SDay(i).RealDay > 0 And i <= 200
    If SDay(i).LuniSolar = True And Leap = SDay(i).LeapMonth Then
      If (SDay(i).M > -1) And SDay(i).y <= -10000 And SDay(i).y > -15000 Then
        If lm = SDay(i).M And LD = SDay(i).D Then
          If R <> "" And SDay(i).DayName <> "" Then R = R & ", "
          R = R & SDay(i).DayName
          If k = 1 And SDay(i).Holy = 2 Then
            k = 3
          ElseIf k = 2 And SDay(i).Holy = 1 Then
            k = 3
          ElseIf k = 3 Then
            k = 3  'k=3이면 Holy 값이 무엇이든 상관없이 3 반환
          ElseIf k > 0 And SDay(i).Holy = 0 Then
            'k = k  '이 경우 아무 처리 안함
          Else
            k = SDay(i).Holy
          End If
        End If
      End If
    End If
    
    i = i + 1
  Loop
  
  h = k
  FindSDayL = R
End Function

Function FindSDayA(ByVal STR As String, ByVal JD As Double, ByVal TZZ As Double) As String
  Dim i As Long, R As String
  
  R = STR: i = 0
  Do While UBound(E) >= i
    If JD = GetJD0(E(i) + TZZ / 24) + 0.5 Then
      If R <> "" And S(i) <> "" Then R = R & ", "
      R = R & S(i) & MakeTimeString(E(i) + TZZ / 24)
    End If
    
    i = i + 1
  Loop
  
  FindSDayA = R
End Function

Private Sub ClearSDay()
  Dim i As Integer
  
  For i = 0 To 200
    With SDay(i)
      .y = -10000
      .M = -5
      .D = 0
      .LuniSolar = False
      .LeapMonth = False
      .DayName = ""
      .RealDay = 0
      .Holy = 0
    End With
  Next i
End Sub

Private Sub ReadDayData()
  Dim Num As Integer
  
  Num = 2

  Do While Sheet2.Cells(Num, 2).Value <> "" And Num < 152
       With SDay(Num - 2)
         If Trim(Sheet2.Cells(Num, 1).Value) = "" Then
           .y = -10000
         ElseIf Trim(Sheet2.Cells(Num, 1).Value) = "x" Then
           .y = -15000
         Else
           .y = CInt(Sheet2.Cells(Num, 1).Value)
         End If
         .M = CInt(Sheet2.Cells(Num, 2).Value)
         .D = CInt(Sheet2.Cells(Num, 3).Value)
         .LuniSolar = Trim(Sheet2.Cells(Num, 5).Value) <> ""
         .LeapMonth = Trim(Sheet2.Cells(Num, 4).Value) <> ""
         .DayName = Trim(Sheet2.Cells(Num, 7).Value)
         .Holy = 0
         .Holy = Sheet2.Cells(Num, 6).Value
       End With
       Num = Num + 1
  Loop
End Sub

Function FindEaster(ByVal y As Integer) As Double
  Dim a As Integer, B As Integer, c As Integer, D As Integer
  Dim E As Integer, f As Integer, g As Integer, h As Integer, i As Integer
  Dim k As Integer, l As Integer, M As Integer, N As Integer, P As Integer
  Dim mon As Double, Da As Double
  
  If y >= 1583 Then
    a = y Mod 19
    B = y \ 100
    c = y Mod 100
    D = B \ 4
    E = B Mod 4
    f = (B + 8) \ 25
    g = (B - f + 1) \ 3
    h = (19 * a + B - D - g + 15) Mod 30
    i = c \ 4
    k = c Mod 4
    l = (32 + 2 * E + 2 * i - h - k) Mod 7
    M = (a + 11 * h + 22 * l) \ 451
    N = (h + l - 7 * M + 114) \ 31
    P = (h + l - 7 * M + 114) Mod 31
    mon = CDbl(N): Da = CDbl(P + 1)
  Else
    a = y Mod 4
    B = y Mod 7
    c = y Mod 19
    D = (19 * c + 15) Mod 30
    E = (2 * a + 4 * B - D + 34) Mod 7
    f = (D + E + 114) \ 31
    g = (D + E + 114) Mod 31
    mon = CDbl(f): Da = CDbl(g + 1)
  End If
  FindEaster = JULIANDAY(CDbl(y), mon, Da, 12, 0)
End Function

