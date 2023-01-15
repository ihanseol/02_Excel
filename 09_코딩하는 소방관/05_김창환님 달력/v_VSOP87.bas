Attribute VB_Name = "v_VSOP87"
Option Explicit

Sub VSOP87_FK5(ByVal JDE As Double, l As Double, B As Double)
  Dim T As Double, ll As Double, cLL As Double, sLL As Double
  
  T = (JDE - 2451545#) / 36525#
  ll = l - 1.397 * T - 0.00031 * T * T
  cLL = Cosd(ll): sLL = Sind(ll)
  
  l = l - 2.50916666666667E-05 + 1.08777777777778E-05 * (cLL + sLL) * Tand(B)
  B = B + 1.08777777777778E-05 * (cLL - sLL)
End Sub

'입력: JED
'출력: L(deg), B(deg), R(AU)
'설명: 출력은 입력한 날의 분점으로 계산된 결과임.
Sub GetLBR(ByVal JDE As Double, ByVal Planet As Byte, ByVal FK5 As Boolean, l As Double, B As Double, R As Double)
  Dim T As Double, L0 As Double, B0 As Double, R0 As Double
  
  T = (JDE - 2451545#) / 365250
  
  Select Case Planet
    Case 2 '지구, -2000~+6000 범위에서 1초 이하
      L0 = Earth_L0(T) + Earth_L1(T) + Earth_L2(T) + Earth_L3(T) + Earth_L4(T) + Earth_L5(T)
      B0 = Earth_B0(T) + Earth_B1(T)
      R0 = Earth_R0(T) + Earth_R1(T) + Earth_R2(T) + Earth_R3(T) + Earth_R4(T)
'      LBR_LOW 0, JDE, l0, b0, r0: l0 = (l0 - 180) * DegtoRad
     Case 9 '달, 경도 10초, 위도 4초 이하
      Call MoonLBR(JDE, L0, B0, R0)  '이미 광차가 보정된 결과
'      LBR_LOW 1, JDE, l0, b0, r0: l0 = l0 * DegtoRad: b0 = b0 * DegtoRad
      R0 = R0 / 149597870
   End Select
   
   L0 = Rev(L0 * RadtoDeg)
   B0 = B0 * RadtoDeg
   
   If FK5 And Planet < 9 Then Call VSOP87_FK5(JDE, L0, B0)
   
   l = Rev(L0)
   B = B0
   R = R0
End Sub

 Sub GetLBR2000(ByVal JDE As Double, ByVal Planet As Byte, ByVal FK5 As Boolean, l As Double, B As Double, R As Double)
     Dim L0 As Double, B0 As Double, L1 As Double, B1 As Double
     
     Call GetLBR(JDE, Planet, False, L0, B0, R)
     PrecessionEcl JDE, 2451545, L0, B0, L1, B1
     If FK5 And Planet < 9 Then Call VSOP87_FK5(2451545, L1, B1)
     
     l = Rev(L1): B = B1
 End Sub

Sub LBR_LOW(ByVal P As Byte, ByVal JD As Double, Lamda As Double, Beta As Double, R As Double)
  Dim N As Double, l As Double, g As Double, T As Double, E As Double
  
  N = JD - 2451545#
  If P = 0 Then  'sun
    l = 280.46 + 0.9856474 * N
    g = 357.528 + 0.9856003 * N
    Lamda = l + 1.915 * Sind(g) + 0.02 * Sind(2 * g)
    Beta = N / 36525
    E = 0.016708634 - 0.000042037 * Beta - 0.0000001267 * Beta * Beta
    Beta = 0
    R = 1.000001018 * (1 - E * E) / (1 + E * Cosd(g + Lamda - l))
    
  Else  'moon
    T = N / 36525
    l = 218.32 + 481267.883 * T + 6.29 * Sind(134.9 + 477198.85 * T) - 1.27 * Sind(259.2 - 413335.38 * T) _
      + 0.66 * Sind(235.7 + 890534.23 * T) + 0.21 * Sind(269.9 + 954397.7 * T) - 0.19 * Sind(357.5 + 35999.05) _
      - 0.11 * Sind(186.6 + 966404.05 * T)
    g = 5.13 * Sind(93.3 + 483202.03 * T) + 0.28 * Sind(228.2 + 960400.87 * T) - 0.28 * Sind(318.3 + 6003.18 * T) - 0.17 * Sind(217.6 - 407332.2 * T)
    N = 0.9508 + 0.0518 * Cosd(134.9 + 477198.85 * T) + 0.0095 * Cosd(259.2 - 413335.38 * T) _
      + 0.0078 * Cosd(235.7 + 890534.23 * T) + 0.0028 * Cosd(269.9 + 954397.7 * T)
    N = 1 / Sind(N)
    R = N * 6378.14
    Beta = g
    Lamda = l
  End If
End Sub
