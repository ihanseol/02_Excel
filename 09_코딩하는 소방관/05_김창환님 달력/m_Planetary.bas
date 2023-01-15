Attribute VB_Name = "m_Planetary"
Option Explicit

Public E() As Double, S() As String

'findppheno  julianday(2008,1,1,12,0),julianday(2008,12,31,12,0)
Sub FindPPheno(ByVal BeginJD As Double, ByVal EndJD As Double)
  Dim i As Integer, j As Integer, D As Double, X As Double, y As Double, Z As Double, PSTR As String
  Dim PN As String, PNA() As String, N As Long, l As Long, M As Long, O As Long
  Dim EE As Double, ss As String
  
  If EndJD <= BeginJD Then Exit Sub
  PN = "수성,금성,화성,목성,토성,천왕성,해왕성"
  PNA = Split(PN, ",")
  l = Int((EndJD - BeginJD) / 365.25) * 600
  If l < 600 Then l = 600
  If l > 30000 Then MsgBox "선택한 기간이 너무 깁니다. 50년 미만으로 줄여주세요.", vbOKOnly & vbExclamation, "오류": Exit Sub
  l = l + 50
  ReDim E(l): ReDim S(l)
  
  For M = 0 To l
    S(M) = "x": E(M) = 9999999999#
  Next M
  
  D = BeginJD - 90: N = 0
  Do
    For i = 1 To 7
      For j = 0 To 1
        PN = PNA(i - 1)
        
        X = FindPlanet(D, i, j) '+ TimeZone / 24
        
        If j = 0 Then
          If i < 3 Then PSTR = PN & " 내합" Else PSTR = PN & " 충"
        Else
          If i < 3 Then PSTR = PN & " 외합" Else PSTR = PN & " 합"
        End If
        If BeginJD <= X And X <= EndJD Then E(N) = X: S(N) = PSTR: N = N + 1
        
        If i < 6 Then
          y = StationPlanet(D, i, j) '+ TimeZone / 24
          PSTR = PN & " 유"
          If BeginJD <= y And y <= EndJD Then E(N) = y: S(N) = PSTR: N = N + 1
        End If
        
        If i < 3 Then
          Z = ElongPlanet(D, i, j) '+ TimeZone / 24
          If j = 0 Then
            PSTR = PN & " 동방최대이각"
          Else
            PSTR = PN & " 서방최대이각"
          End If
          If BeginJD <= Z And Z <= EndJD Then E(N) = Z: S(N) = PSTR: N = N + 1
        End If
      Next j
    Next i
    D = D + 30
  Loop Until EndJD + 90 < D
  
  D = BeginJD - 20
  Do
    For i = 0 To 3
      X = Round(GetMoon(D, i * 90, TimeZone), 4) - TimeZone / 24
      If i = 0 Then PSTR = "그믐"
      If i = 1 Then PSTR = "상현"
      If i = 2 Then PSTR = "보름"
      If i = 3 Then PSTR = "하현"
      
      If BeginJD <= X And X <= EndJD Then E(N) = X: S(N) = PSTR: N = N + 1
    Next i
    D = D + 27
  Loop Until EndJD + 50 < D
  
  Dim MM As Long, EEE As Double
  For M = 0 To l - 1
    EEE = E(M): MM = M
    For O = M + 1 To l
      If E(O) < EEE Then
        EEE = E(O): MM = O
      End If
    Next O
    EE = E(M): ss = S(M)
    E(M) = E(MM): S(M) = S(MM)
    E(MM) = EE: S(MM) = ss
  Next M
  
  For M = 0 To l - 1 'err
    If E(M) = E(M + 1) Then
      S(M) = "x"
    End If
  Next M
  
  N = 0
  For M = 0 To l - 1
    If S(M) <> "x" Then
      S(N) = S(M): E(N) = E(M)
      N = N + 1
    End If
  Next M
  
  If N > 0 Then
    ReDim Preserve E(N - 1)
    ReDim Preserve S(N - 1)
  Else
    ReDim E(0)
    ReDim S(0)
  End If
End Sub

'출력은 TT, J=0: 내합,충, 1: 외합, 합
'test: ? maketimestring3(findplanet(julianday(1631,11,7,0,0),1,0),1)
Function FindPlanet(ByVal JD As Double, ByVal P As Integer, ByVal j As Integer) As Double
  Dim k As Double, A1 As Double, B1 As Double, M1 As Double, M0 As Double
  Dim y As Double, Mo As Double, Da As Double, JDE0 As Double, M As Double, T As Double, ct As Double
  Dim a As Double, B As Double, c As Double, D As Double, E As Double, f As Double, g As Double, R As Double
  
  InvJD JD, y, Mo, Da, A1, B1: y = y + Mo / 12 + Da / 365
  SetConst P, j, A1, B1, M0, M1: R = 0
  
  k = Int((365.2425 * y + 1721060 - A1) / B1)
  JDE0 = A1 + k * B1
  M = M0 + k * M1
  
  T = (JDE0 - 2451545) / 36525
  
  a = 0: B = 0: c = 0: D = 0: E = 0: f = 0: g = 0
  If P = 4 Then
    a = 82.74 + 40.76 * T
  ElseIf P = 5 Then
    a = 82.74 + 40.76 * T
    B = 29.86 + 1181.36 * T
    c = 14.13 + 590.68 * T
    D = 220.02 + 1262.87 * T
  ElseIf P = 6 Then
    E = 207.83 + 8.51 * T
    f = 108.84 + 419.96 * T
  ElseIf P = 7 Then
    E = 207.83 + 8.51 * T
    g = 276.74 + 209.98 * T
  End If
  
  ct = CorTime(P, j, T, M, a, B, c, D, E, f, g)
  R = JDE0 + ct
  
  FindPlanet = R
End Function

'금성, 수성의 최대 이각, 출력은 TT, J=0 동방최대이각, 1: 서방최대이각
'test: ? maketimestring3(ElongPlanet(julianday(1993,11,22,0,0),1 ,1),1)
Function ElongPlanet(ByVal JD As Double, ByVal P As Integer, ByVal j As Integer) As Double
  Dim k As Double, A1 As Double, B1 As Double, M1 As Double, M0 As Double
  Dim y As Double, Mo As Double, Da As Double, JDE0 As Double, M As Double, T As Double, ct As Double, R As Double
    
  InvJD JD, y, Mo, Da, A1, B1
  y = y + Mo / 12 + Da / 365
  SetConst P, 0, A1, B1, M0, M1: R = 0
  
  k = Int((365.2425 * y + 1721060 - A1) / B1)
  JDE0 = A1 + k * B1
  M = M0 + k * M1
  
  T = (JDE0 - 2451545) / 36525
  ct = CorTime2(P, j, T, M)
  R = JDE0 + ct
  
  ElongPlanet = R
End Function

'수성~토성의 유, 출력은 TT, J=0 1st Stn., 1: 2nd Stn
'test: ? maketimestring3(StationPlanet(julianday(1631,11,7,0,0),1,0),1)
Function StationPlanet(ByVal JD As Double, ByVal P As Integer, ByVal j As Integer) As Double
  Dim k As Double, A1 As Double, B1 As Double, M1 As Double, M0 As Double
  Dim y As Double, Mo As Double, Da As Double, JDE0 As Double, M As Double, T As Double, ct As Double, R As Double
  Dim a As Double, B As Double, c As Double, D As Double
    
  InvJD JD, y, Mo, Da, A1, B1
  y = y + Mo / 12 + Da / 365
  SetConst P, 0, A1, B1, M0, M1: R = 0
  
  k = Int((365.2425 * y + 1721060 - A1) / B1)
  JDE0 = A1 + k * B1
  M = M0 + k * M1
  
  T = (JDE0 - 2451545) / 36525
  
  a = 0: B = 0: c = 0: D = 0
  If P = 4 Then
    a = 82.74 + 40.76 * T
  ElseIf P = 5 Then
    a = 82.74 + 40.76 * T
    B = 29.86 + 1181.36 * T
    c = 14.13 + 590.68 * T
    D = 220.02 + 1262.87 * T
  End If
  
  ct = CorTime3(P, j, T, M, a, B, c, D)
  R = JDE0 + ct
  
  StationPlanet = R
End Function

Function CorTime(ByVal P As Integer, ByVal j As Integer, ByVal T As Double, ByVal M As Double, ByVal a As Double, ByVal B As Double, _
                           ByVal c As Double, ByVal D As Double, ByVal E As Double, ByVal f As Double, ByVal g As Double) As Double
  Dim sm1 As Double, cm1 As Double, sm2 As Double, cm2 As Double, sm3 As Double, cm3 As Double, T2 As Double
  Dim sm4 As Double, cm4 As Double, sm5 As Double, cm5 As Double, R As Double
  
  T2 = T * T: sm1 = Sind(M): cm1 = Cosd(M): sm2 = Sind(2 * M): cm2 = Cosd(2 * M): sm3 = Sind(3 * M): cm3 = Cosd(3 * M)
  sm5 = Sind(5 * M): cm5 = Cosd(5 * M): sm4 = Sind(4 * M): cm4 = Cosd(4 * M): R = 0
  
  If P = 1 And j = 0 Then
    R = 0.0545 + 0.0002 * T
    R = R + sm1 * (-6.2008 + 0.0074 * T + 0.00003 * T2) + cm1 * (-3.275 - 0.0197 * T + 0.00001 * T2)
    R = R + sm2 * (0.4737 - 0.0052 * T - 0.00001 * T2) + cm2 * (0.8111 + 0.0033 * T - 0.00002 * T2)
    R = R + sm3 * (0.0037 + 0.0018 * T) + cm3 * (-0.1768 + 0.00001 * T2)
    R = R + sm4 * (-0.0211 - 0.0004 * T) + cm4 * (0.0326 - 0.0003 * T)
    R = R + sm5 * (0.0083 + 0.0001 * T) + cm5 * (-0.004 + 0.0001 * T)
  ElseIf P = 1 And j = 1 Then
    R = -0.0548 - 0.0002 * T
    R = R + sm1 * (7.3894 - 0.01 * T - 0.00003 * T2) + cm1 * (3.22 + 0.0197 * T - 0.00001 * T2)
    R = R + sm2 * (0.8383 - 0.0064 * T - 0.00001 * T2) + cm2 * (0.9666 + 0.0039 * T - 0.00003 * T2)
    R = R + sm3 * (0.077 - 0.0026 * T) + cm3 * (0.2758 + 0.0002 * T - 0.00002 * T2)
    R = R + sm4 * (-0.0128 - 0.0008 * T) + cm4 * (0.0734 - 0.0004 * T - 0.00001 * T2)
    R = R + sm5 * (-0.0122 - 0.0002 * T) + cm5 * (0.0173 - 0.0002 * T2)
  ElseIf P = 2 And j = 0 Then
    R = -0.0096 + 0.0002 * T - 0.00001 * T2
    R = R + sm1 * (2.0009 - 0.0033 * T - 0.00001 * T2) + cm1 * (0.598 - 0.0104 * T + 0.00001 * T2)
    R = R + sm2 * (0.0967 - 0.0016 * T - 0.00003 * T2) + cm2 * (0.0913 + 0.0009 * T - 0.00002 * T2)
    R = R + sm3 * (0.0046 - 0.0002 * T) + cm3 * (0.0079 + 0.0001 * T)
  ElseIf P = 2 And j = 1 Then
    R = 0.0099 - 0.0002 * T - 0.00001 * T2
    R = R + sm1 * (4.1991 - 0.0121 * T - 0.00003 * T2) + cm1 * (-0.6095 + 0.0102 * T - 0.00002 * T2)
    R = R + sm2 * (0.25 - 0.0028 * T - 0.00003 * T2) + cm2 * (0.0063 + 0.0025 * T - 0.00002 * T2)
    R = R + sm3 * (0.0232 - 0.0005 * T - 0.00001 * T2) + cm3 * (0.0031 + 0.0004 * T)
  ElseIf P = 3 And j = 0 Then
    R = -0.3088 + 0.00002 * T2
    R = R + sm1 * (-17.6965 + 0.0363 * T + 0.00005 * T2) + cm1 * (18.3131 + 0.0467 * T - 0.00006 * T2)
    R = R + sm2 * (-0.2162 - 0.0198 * T - 0.00001 * T2) + cm2 * (-4.5028 - 0.0019 * T + 0.00007 * T2)
    R = R + sm3 * (0.8987 + 0.0058 * T - 0.00002 * T2) + cm3 * (0.7666 - 0.005 * T - 0.00003 * T2)
    R = R + sm4 * (-0.3636 - 0.0001 * T + 0.00002 * T2) + cm4 * (0.0402 + 0.0032 * T)
    R = R + sm5 * (0.0737 - 0.0008 * T) + cm5 * (-0.098 - 0.0011 * T)
  ElseIf P = 3 And j = 1 Then
    R = 0.3102 - 0.0001 * T + 0.00001 * T2
    R = R + sm1 * (9.7273 - 0.0156 * T + 0.00001 * T) + cm1 * (-18.3195 - 0.0467 * T + 0.00009 * T2)
    R = R + sm2 * (-1.6488 - 0.0133 * T + 0.00001 * T2) + cm2 * (-2.6117 - 0.002 * T + 0.00004 * T2)
    R = R + sm3 * (-0.6827 - 0.0026 * T + 0.00001 * T2) + cm3 * (0.0281 + 0.0035 * T + 0.00001 * T2)
    R = R + sm4 * (-0.0823 + 0.0006 * T + 0.00001 * T2) + cm4 * (0.1584 + 0.0013 * T)
    R = R + sm5 * (0.027 + 0.0005 * T) + cm5 * 0.0433
  ElseIf P = 4 And j = 0 Then
    R = -0.1029 - 0.00009 * T2
    R = R + sm1 * (-1.9658 - 0.0056 * T + 0.00007 * T2) + cm1 * (6.1537 + 0.021 * T - 0.00006 * T2)
    R = R + sm2 * (-0.2081 - 0.0013 * T) + cm2 * (-0.1116 - 0.001 * T)
    R = R + sm3 * (0.0074 + 0.0001 * T) + cm3 * (-0.0097 - 0.0001 * T)
    R = R + Sind(a) * (0.0144 * T - 0.00008 * T2) + Cosd(a) * (0.3642 - 0.0019 * T - 0.00029 * T2)
  ElseIf P = 4 And j = 1 Then
    R = 0.1027 - 0.0002 * T - 0.00009 * T2
    R = R + sm1 * (-2.2637 + 0.0163 * T - 0.00003 * T2) + cm1 * (-6.154 - 0.021 * T - 0.00003 * T2)
    R = R + sm2 * (-0.2021 - 0.0017 * T + 0.00001 * T2) + cm2 * (0.131 - 0.0008 * T)
    R = R + sm3 * 0.0086 + cm3 * (0.0087 + 0.0002 * T)
    R = R + Sind(a) * (0.0144 * T - 0.00008 * T2) + Cosd(a) * (0.3642 - 0.0019 * T - 0.00029 * T2)
  ElseIf P = 5 And j = 0 Then
    R = -0.0209 + 0.0006 * T + 0.00023 * T2
    R = R + sm1 * (4.5795 - 0.0312 * T - 0.00017 * T2) + cm1 * (1.1462 - 0.0351 * T + 0.00011 * T2)
    R = R + sm2 * (0.0985 - 0.0015 * T) + cm2 * (0.0733 - 0.0031 * T + 0.00001 * T2)
    R = R + sm3 * (0.0025 - 0.0001 * T) + cm3 * (0.005 - 0.0002 * T)
    R = R + Sind(a) * (-0.0337 * T + 0.00018 * T2) + Cosd(a) * (-0.851 + 0.0044 * T + 0.00068 * T2)
    R = R + Sind(B) * (-0.0064 * T + 0.00004 * T2) + Cosd(B) * (0.2397 - 0.0012 * T - 0.00008 * T2)
    R = R + Sind(c) * (-0.001 * T) + Cosd(c) * (0.1245 + 0.0006 * T)
    R = R + Sind(D) * (0.0024 * T - 0.00003 * T2) + Cosd(D) * (0.0477 - 0.0005 * T - 0.00006 * T2)
  ElseIf P = 5 And j = 1 Then
    R = 0.0172 - 0.0006 * T + 0.00023 * T2
    R = R + sm1 * (-8.5885 + 0.0411 * T + 0.0002 * T2) + cm1 * (-1.147 + 0.0352 * T - 0.00011 * T2)
    R = R + sm2 * (0.3331 - 0.0034 * T - 0.00001 * T2) + cm2 * (0.1145 - 0.0045 * T + 0.00002 * T2)
    R = R + sm3 * (-0.0169 + 0.0002 * T) + cm3 * (-0.0109 + 0.0004 * T)
    R = R + Sind(a) * (-0.0337 * T + 0.00018 * T2) + Cosd(a) * (-0.851 + 0.0044 * T + 0.00068 * T2)
    R = R + Sind(B) * (-0.0064 * T + 0.00004 * T2) + Cosd(B) * (0.2397 - 0.0012 * T - 0.00008 * T2)
    R = R + Sind(c) * (-0.001 * T) + Cosd(c) * (0.1245 + 0.0006 * T)
    R = R + Sind(D) * (0.0024 * T - 0.00003 * T2) + Cosd(D) * (0.0477 - 0.0005 * T - 0.00006 * T2)
  ElseIf P = 6 And j = 0 Then
    R = 0.0844 - 0.0006 * T
    R = R + sm1 * (-0.1048 + 0.0246 * T) + cm1 * (-5.1221 + 0.0104 * T + 0.00003 * T2)
    R = R + sm2 * (-0.1428 + 0.0005 * T) + cm2 * (-0.0148 - 0.0013 * T)
    R = R + cm3 * 0.0055
    R = R + Cosd(E) * 0.885 + Cosd(f) * 0.2153
  ElseIf P = 6 And j = 1 Then
    R = -0.0859 + 0.0003 * T
    R = R + sm1 * (-3.8179 - 0.0148 * T + 0.00003 * T2) + cm1 * (5.1228 - 0.0105 * T - 0.00002 * T2)
    R = R + sm2 * (-0.0803 + 0.0011 * T) + cm2 * (-0.1905 - 0.0006 * T)
    R = R + sm3 * (0.0088 + 0.0001 * T)
    R = R + Cosd(E) * 0.885 + Cosd(f) * 0.2153
  ElseIf P = 7 And j = 0 Then
    R = -0.014 + 0.00001 * T2
    R = R + sm1 * (-1.3486 + 0.001 * T + 0.00001 * T2) + cm1 * (0.8597 + 0.0037 * T)
    R = R + sm2 * (-0.0082 - 0.0002 * T + 0.00001 * T2) + cm2 * (0.0037 - 0.0003 * T)
    R = R + Cosd(E) * -0.5964 + Cosd(g) * 0.0728
  ElseIf P = 7 And j = 1 Then
    R = 0.0168
    R = R + sm1 * (-2.5606 + 0.0088 * T + 0.00002 * T2) + cm1 * (-0.8611 - 0.0037 * T + 0.00002 * T2)
    R = R + sm2 * (0.0118 - 0.0004 * T - 0.00001 * T2) + cm2 * (0.0307 - 0.0003 * T)
    R = R + Cosd(E) * -0.5964 + Cosd(g) * 0.0728
  End If
  
  CorTime = R
End Function

Function CorTime2(ByVal P As Integer, ByVal j As Integer, ByVal T As Double, ByVal M As Double) As Double
  Dim sm1 As Double, cm1 As Double, sm2 As Double, cm2 As Double, sm3 As Double, cm3 As Double, T2 As Double
  Dim sm4 As Double, cm4 As Double, sm5 As Double, cm5 As Double, R As Double
  
  T2 = T * T: sm1 = Sind(M): cm1 = Cosd(M): sm2 = Sind(2 * M): cm2 = Cosd(2 * M): sm3 = Sind(3 * M): cm3 = Cosd(3 * M)
  sm5 = Sind(5 * M): cm5 = Cosd(5 * M): sm4 = Sind(4 * M): cm4 = Cosd(4 * M): R = 0
  
  If P = 1 And j = 0 Then
    R = -21.6101 + 0.0002 * T
    R = R + sm1 * (-1.9803 - 0.006 * T + 0.00001 * T2) + cm1 * (1.4151 - 0.0072 * T - 0.00001 * T2)
    R = R + sm2 * (0.5528 - 0.0005 * T - 0.00001 * T2) + cm2 * (0.2905 + 0.0034 * T + 0.00001 * T2)
    R = R + sm3 * (-0.1121 - 0.0001 * T + 0.00001 * T2) + cm3 * (-0.0098 - 0.0015 * T)
    R = R + sm4 * 0.0192 + cm4 * (0.0111 + 0.0004 * T)
    R = R + sm5 * -0.0061 + cm5 * (-0.0032 - 0.0001 * T)
  ElseIf P = 1 And j = 1 Then
    R = 21.6249 - 0.0002 * T
    R = R + sm1 * (0.1306 + 0.0065 * T) + cm1 * (-2.7661 - 0.0011 * T + 0.00001 * T2)
    R = R + sm2 * (0.2438 - 0.0024 * T - 0.00001 * T2) + cm2 * (0.5767 + 0.0023 * T)
    R = R + sm3 * 0.1041 + cm3 * (-0.0184 + 0.0007 * T)
    R = R + sm4 * (-0.0051 - 0.0001 * T) + cm4 * (0.0048 + 0.0001 * T)
    R = R + sm5 * 0.0026 + cm5 * 0.0037
  ElseIf P = 2 And j = 0 Then
    R = -70.76 + 0.0002 * T - 0.00001 * T2
    R = R + sm1 * (1.0282 - 0.001 * T - 0.00001 * T2) + cm1 * (0.2761 - 0.006 * T)
    R = R + sm2 * (-0.0438 - 0.0023 * T + 0.00002 * T2) + cm2 * (0.166 - 0.0037 * T - 0.00004 * T2)
    R = R + sm3 * (0.0036 + 0.0001 * T) + cm3 * (-0.0011 + 0.00001 * T2)
  ElseIf P = 2 And j = 1 Then
    R = 70.7462 - 0.00001 * T2
    R = R + sm1 * (1.1218 - 0.0025 * T - 0.00001 * T2) + cm1 * (0.4538 - 0.0066 * T)
    R = R + sm2 * (0.132 + 0.002 * T - 0.00003 * T2) + cm2 * (-0.0702 + 0.0022 * T + 0.00004 * T2)
    R = R + sm3 * (0.0062 - 0.0001 * T) + cm3 * (0.0015 - 0.00001 * T2)
  End If
  
  CorTime2 = R
End Function

Function CorTime3(ByVal P As Integer, ByVal j As Integer, ByVal T As Double, ByVal M As Double, ByVal a As Double, ByVal B As Double, _
                           ByVal c As Double, ByVal D As Double) As Double
  Dim sm1 As Double, cm1 As Double, sm2 As Double, cm2 As Double, sm3 As Double, cm3 As Double, T2 As Double
  Dim sm4 As Double, cm4 As Double, sm5 As Double, cm5 As Double, R As Double
  
  T2 = T * T: sm1 = Sind(M): cm1 = Cosd(M): sm2 = Sind(2 * M): cm2 = Cosd(2 * M): sm3 = Sind(3 * M): cm3 = Cosd(3 * M)
  sm5 = Sind(5 * M): cm5 = Cosd(5 * M): sm4 = Sind(4 * M): cm4 = Cosd(4 * M): R = 0
  
  If P = 1 And j = 0 Then
    R = -11.0761 + 0.0003 * T
    R = R + sm1 * (-4.7321 + 0.0023 * T + 0.00002 * T2) + cm1 * (-1.323 - 0.0156 * T2)
    R = R + sm2 * (0.227 - 0.0046 * T) + cm2 * (0.7184 + 0.0013 * T - 0.00002 * T2)
    R = R + sm3 * (0.0638 + 0.0016 * T) + cm3 * (-0.1655 + 0.0007 * T)
    R = R + sm4 * (-0.0395 - 0.0003 * T) + cm4 * (0.0247 - 0.0006 * T)
    R = R + sm5 * 0.0131 + cm3 * (0.0008 + 0.0002 * T)
  ElseIf P = 1 And j = 1 Then
    R = 11.1343 - 0.0001 * T
    R = R + sm1 * (-3.9137 + 0.0073 * T + 0.00002 * T2) + cm1 * (-3.3861 - 0.0128 * T + 0.00001 * T2)
    R = R + sm2 * (0.5222 - 0.004 * T - 0.00002 * T2) + cm2 * (0.5929 + 0.0039 * T - 0.00002 * T2)
    R = R + sm3 * (-0.0593 + 0.0018 * T) + cm3 * (-0.1733 - 0.0007 * T + 0.00001 * T2)
    R = R + sm4 * (-0.0053 - 0.0006 * T) + cm4 * (0.0476 - 0.0001 * T)
    R = R + sm5 * (0.007 + 0.0002 * T) + cm5 * (-0.0115 + 0.0001 * T)
  ElseIf P = 2 And j = 0 Then
    R = -21.0672 + 0.0002 * T - 0.00001 * T2
    R = R + sm1 * (1.9396 - 0.0029 * T - 0.00001 * T2) + cm1 * (1.0727 - 0.0102 * T)
    R = R + sm2 * (0.0404 - 0.0023 * T - 0.00001 * T2) + cm2 * (0.1305 - 0.0004 * T - 0.00003 * T2)
    R = R + sm3 * (-0.0007 - 0.0002 * T) + cm3 * 0.0098
  ElseIf P = 2 And j = 1 Then
    R = 21.0623 - 0.00001 * T2
    R = R + sm1 * (1.9913 - 0.004 * T - 0.00001 * T2) + cm1 * (-0.0407 - 0.0077 * T)
    R = R + sm2 * (0.1351 - 0.0009 * T - 0.00004 * T2) + cm2 * (0.0303 + 0.0019 * T)
    R = R + sm3 * (0.0089 - 0.0002 * T) + cm3 * (0.0043 + 0.0001 * T)
  ElseIf P = 3 And j = 0 Then
    R = -37.079 - 0.0009 * T + 0.00002 * T2
    R = R + sm1 * (-20.0651 + 0.0228 * T + 0.00004 * T2) + cm1 * (14.5205 + 0.0504 * T - 0.00001 * T2)
    R = R + sm2 * (1.1737 - 0.0169 * T) + cm2 * (-4.255 - 0.0075 * T + 0.00008 * T2)
    R = R + sm3 * (0.4897 + 0.0074 * T - 0.00001 * T2) + cm3 * (1.1151 - 0.0021 * T - 0.00005 * T2)
    R = R + sm4 * (-0.3636 - 0.002 * T + 0.00001 * T2) + cm4 * (-0.1769 + 0.0028 * T + 0.00002 * T2)
    R = R + sm5 * (0.1437 - 0.0004 * T) + cm5 * (-0.0383 - 0.0016 * T)
  ElseIf P = 3 And j = 1 Then
    R = 36.7191 + 0.0016 * T + 0.00003 * T2
    R = R + sm1 * (-12.6163 + 0.0417 * T - 0.00001 * T2) + cm1 * (20.1218 + 0.0379 * T - 0.00006 * T2)
    R = R + sm2 * (-1.636 - 0.019 * T) + cm2 * (-3.9657 + 0.0045 * T + 0.00007 * T2)
    R = R + sm3 * (1.1546 + 0.0029 * T - 0.00003 * T2) + cm3 * (0.2888 - 0.0073 * T - 0.00002 * T2)
    R = R + sm4 * (-0.3128 + 0.0017 * T + 0.00002 * T2) + cm4 * (0.2513 + 0.0026 * T - 0.00002 * T2)
    R = R + sm5 * (-0.0021 - 0.0016 * T) + cm5 * (-0.1497 - 0.0006 * T)
  ElseIf P = 4 And j = 0 Then
    R = -60.367 - 0.0001 * T - 0.00009 * T2
    R = R + sm1 * (-2.3144 - 0.0124 * T + 0.00007 * T2) + cm1 * (6.7439 + 0.0166 * T - 0.00006 * T2)
    R = R + sm2 * (-0.2259 - 0.001 * T) + cm2 * (-0.1497 - 0.0014 * T)
    R = R + sm3 * (0.0105 + 0.0001 * T) + cm3 * -0.0098
    R = R + Sind(a) * (0.0144 * T - 0.00008 * T2) + Cosd(a) * (0.3642 - 0.0019 * T - 0.00029 * T2)
  ElseIf P = 4 And j = 1 Then
    R = 60.3023 + 0.0002 * T - 0.00009 * T2
    R = R + sm1 * (0.3506 - 0.0034 * T + 0.00004 * T2) + cm1 * (5.3635 + 0.0274 * T - 0.00007 * T2)
    R = R + sm2 * (-0.1872 - 0.0016 * T) + cm2 * (-0.0037 - 0.0005 * T)
    R = R + sm3 * (0.0012 + 0.0001 * T) + cm3 * (-0.0096 - 0.0001 * T)
    R = R + Sind(a) * (0.0144 * T - 0.00008 * T2) + Cosd(a) * (0.3642 - 0.0019 * T - 0.00029 * T2)
  ElseIf P = 5 And j = 0 Then
    R = -68.884 + 0.0009 * T + 0.00023 * T2
    R = R + sm1 * (5.5452 - 0.0279 * T - 0.0002 * T2) + cm1 * (3.0727 - 0.043 * T + 0.00007 * T2)
    R = R + sm2 * (0.1101 - 0.0006 * T - 0.00001 * T2) + cm2 * (0.1654 - 0.0043 * T + 0.00001 * T2)
    R = R + sm3 * (0.001 + 0.0001 * T) + cm3 * (0.0095 - 0.0003 * T)
    R = R + Sind(a) * (-0.0337 * T + 0.00018 * T2) + Cosd(a) * (-0.851 + 0.0044 * T + 0.00068 * T2)
    R = R + Sind(B) * (-0.0064 * T + 0.00004 * T2) + Cosd(B) * (0.2397 - 0.0012 * T - 0.00008 * T2)
    R = R + Sind(c) * (-0.001 * T) + Cosd(c) * (0.1245 + 0.0006 * T)
    R = R + Sind(D) * (0.0024 * T - 0.00003 * T2) + Cosd(D) * (0.0477 - 0.0005 * T - 0.00006 * T2)
  ElseIf P = 5 And j = 1 Then
    R = 68.872 - 0.0007 * T + 0.00023 * T2
    R = R + sm1 * (5.9399 - 0.04 * T - 0.00015 * T2) + cm1 * (-0.7998 - 0.0266 * T + 0.00014 * T2)
    R = R + sm2 * (0.1738 - 0.0032 * T) + cm2 * (-0.0039 - 0.0024 * T + 0.00001 * T2)
    R = R + sm3 * (0.0073 - 0.0002 * T) + cm3 * (0.002 - 0.0002 * T)
    R = R + Sind(a) * (-0.0337 * T + 0.00018 * T2) + Cosd(a) * (-0.851 + 0.0044 * T + 0.00068 * T2)
    R = R + Sind(B) * (-0.0064 * T + 0.00004 * T2) + Cosd(B) * (0.2397 - 0.0012 * T - 0.00008 * T2)
    R = R + Sind(c) * (-0.001 * T) + Cosd(c) * (0.1245 + 0.0006 * T)
    R = R + Sind(D) * (0.0024 * T - 0.00003 * T2) + Cosd(D) * (0.0477 - 0.0005 * T - 0.00006 * T2)
  End If
  
  CorTime3 = R
End Function

Sub SetConst(ByVal P As Integer, ByVal j As Integer, a As Double, B As Double, M0 As Double, M1 As Double)
  If P = 1 Then
    B = 115.8774771: M1 = 114.2088742
    If j = 0 Then a = 2451612.023: M0 = 63.5867
    If j = 1 Then a = 2451554.084: M0 = 6.4822
  End If
  
  If P = 2 Then
    B = 583.921361: M1 = 215.513058
    If j = 0 Then a = 2451996.706: M0 = 82.7311
    If j = 1 Then a = 2451704.746: M0 = 154.9745
  End If
  
  If P = 3 Then
    B = 779.936104: M1 = 48.705244
    If j = 0 Then a = 2452097.382: M0 = 181.9573
    If j = 1 Then a = 2451707.414: M0 = 157.6047
  End If
  
  If P = 4 Then
    B = 398.884046: M1 = 33.140229
    If j = 0 Then a = 2451870.628: M0 = 318.4681
    If j = 1 Then a = 2451671.186: M0 = 121.898
  End If
  
  If P = 5 Then
    B = 378.091904: M1 = 12.647487
    If j = 0 Then a = 2451870.17: M0 = 318.0172
    If j = 1 Then a = 2451681.124: M0 = 131.6934
  End If
  
  If P = 6 Then
    B = 369.656035: M1 = 4.333093
    If j = 0 Then a = 2451764.317: M0 = 213.6884
    If j = 1 Then a = 2451579.489: M0 = 31.5219
  End If
  
  If P = 7 Then
    B = 367.486703: M1 = 2.194998
    If j = 0 Then a = 2451753.122: M0 = 202.6544
    If j = 1 Then a = 2451569.379: M0 = 21.5569
  End If
End Sub
