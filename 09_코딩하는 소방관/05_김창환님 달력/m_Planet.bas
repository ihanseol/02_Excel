Attribute VB_Name = "m_Planet"
Option Explicit

Type TPlanetData  'DE404를 위한 자료형
   JD As Double
   l As Double
   B As Double
   R As Double
   ipla As Long
End Type

Dim oblecl As Double   '황도의 기울기

Public Function ecl(ByVal julday As Double) As Double '황도 경사각
  Dim T As Double, ecl2 As Double
  T = (julday - 2451545#) / 3652500
  ecl2 = -249.67 + (-39.05 + (7.12 + (27.87 + (5.79 + 2.45 * T) * T) * T) * T) * T
  ecl = 23.439291111 + (-4680.93 + (-1.55 + (1999.25 + (-51.38 + ecl2 * T) * T) * T) * T) * T / 3600
End Function

Public Function JDtoTDT(ByVal julday As Double) As Double '지구시 계산(JD 입력)
  Dim y As Long, T As Double, dt As Double

  y = CInt(InvJDYear(julday))
  Select Case y
   Case Is < 949
     T = (y - 2000) / 100
     dt = (2715.6 + 573.36 * T + 46.5 * T * T) / 3600
   Case 949 To 1619
     T = (y - 1850) / 100
     dt = (22.5 * T * T) / 3600
   Case 1620 To 1621: dt = 124 / 3600
   Case 1622 To 1623: dt = 115 / 3600
   Case 1624 To 1625: dt = 106 / 3600
   Case 1626 To 1627: dt = 98 / 3600
   Case 1628 To 1629: dt = 91 / 3600
   Case 1630 To 1631: dt = 85 / 3600
   Case 1632 To 1633: dt = 79 / 3600
   Case 1634 To 1635: dt = 74 / 3600
   Case 1636 To 1637: dt = 70 / 3600
   Case 1638 To 1639: dt = 65 / 3600
   Case 1640 To 1645: dt = 60 / 3600
   Case 1646 To 1653: dt = 50 / 3600
   Case 1654 To 1661: dt = 40 / 3600
   Case 1662 To 1671: dt = 30 / 3600
   Case 1672 To 1681: dt = 20 / 3600
   Case 1682 To 1691: dt = 10 / 3600
   Case 1692 To 1707: dt = 9 / 3600
   Case 1708 To 1717: dt = 10 / 3600
   Case 1718 To 1733: dt = 11 / 3600
   Case 1734 To 1743: dt = 12 / 3600
   Case 1744 To 1751: dt = 13 / 3600
   Case 1752 To 1757: dt = 14 / 3600
   Case 1758 To 1765: dt = 15 / 3600
   Case 1766 To 1775: dt = 16 / 3600
   Case 1776 To 1791: dt = 17 / 3600
   Case 1792 To 1795: dt = 16 / 3600
   Case 1796 To 1797: dt = 15 / 3600
   Case 1798 To 1799: dt = 14 / 3600
   Case 1800 To 1899
     T = (y - 1900) / 100
     dt = (727058.63 + T * 123563.95)
     dt = (2513807.78 + T * (1818961.41 + T * dt))
     dt = (1061660.75 + T * (2087298.89 + T * dt))
     dt = (56282.84 + T * (324011.78 + T * dt))
     dt = -2.5 + T * (228.95 + T * (5218.61 + T * dt))
     dt = dt / 3600
   Case 1900 To 1987
     T = (y - 1900) / 100
     dt = (-0.861938 + T * (0.677066 + T * -0.212591))
     dt = (0.025184 + T * (-0.181133 + T * (0.55304 + T * dt)))
     dt = (-0.00002 + T * (0.000297 + T * dt))
     dt = dt * 24
   Case 1988 To 1996
     T = (y - 2000) / 100
     dt = (67 + 123.5 * T + 32.5 * T * T) / 3600
   Case 1997: dt = 62 / 3600
   Case 1998 To 1999: dt = 63 / 3600
   Case 2000 To 2001: dt = 64 / 3600
   Case 2002 To 2020
     T = (y - 2000) / 100
     dt = (63 + 123.5 * T + 32.5 * T * T) / 3600
   Case Is > 2020
     T = (y - 1875.1) / 100
     dt = 45.39 * T * T / 3600
   Case Else
     dt = 0
  End Select

  JDtoTDT = dt / 24 + julday
End Function

'2항 보간
Function Int2(ByVal X1 As Double, ByVal x2 As Double, ByVal Y1 As Double, ByVal Y2 As Double, ByVal P As Double) As Double
  Dim DX As Double, dy As Double, R As Double
  
  DX = x2 - X1
  dy = Y2 - Y1
  R = (P - X1) / DX
  
  Int2 = Y1 + R * dy
End Function
'============================================여기까지

'이름만 DE404, DLL 사용을 없애기위해 VSOP87+EPL2000/82로 대체
Public Function Plan404(Pla As TPlanetData) As Long
  Dim i As Integer, l As Double, B As Double, R As Double
  
  i = Pla.ipla
  i = i - 1
  If i = 10 Then i = 9
  Pla.ipla = i
  
  GetLBR Pla.JD, CByte(Pla.ipla), True, l, B, R
  l = l * DegtoRad: B = B * DegtoRad
  Pla.l = l: Pla.B = B: Pla.R = R
  If i <> 8 Then Plan404 = 1 Else Plan404 = 0  '명왕성은 계산 안함
End Function

Public Sub PlanetPosB(ByVal pIndex As Integer, ByVal julday As Double, ByVal TZ As Double, ByVal TimeAbbr As Boolean, pRA As Double, pDE As Double, PH As Single)        '행성 궤도 계산
    Dim i As Integer, j As Long, tpla As TPlanetData
    Dim xm As Double, ym As Double, zm As Double, DX As Double, dy As Double, dz As Double
    Dim xe As Double, ye As Double, ze As Double, Rm As Double, MoonDist As Double
    Dim xh As Double, yh As Double, zh As Double
    Dim xg As Double, yg As Double, zg As Double
    Dim FV As Double, elon As Double
    Dim B As Double, ir As Double, nR As Double, D As Double, TDT As Double

    Dim SunLon As Double, SunLat As Double, SunDist As Double, MoonLon As Double, MoonLat As Double
    Dim eln As Double, elt As Double, bc As Double, c As Double, Tc As Double, CorTimeA As Boolean
    Dim gDist As Double, hDist As Double, gLAT As Double, glon As Double
    
    CorTimeA = False: c = 173.144632684657 '=(299792.458 / 149597870.691) * 86400 'speed of light(AU/Day)

    TDT = JDtoTDT(julday - TZ / 24)
    oblecl = ecl(julday - TZ / 24) * DegtoRad
    D = (julday - TZ / 24) - 2451543.5

    '행성 위치 계산(일심황도좌표)
    tpla.JD = TDT   'UT
    i = pIndex

    '지구의 좌표 얻기==============
    tpla.ipla = 3
    j = Plan404(tpla)
    SunLon = tpla.l
    SunLat = tpla.B
    SunDist = tpla.R

    xe = SunDist * Cos(SunLon) * Cos(SunLat)   '지구의 태양중심 황도 좌표
    ye = SunDist * Sin(SunLon) * Cos(SunLat)
    ze = SunDist * Sin(SunLat)
    '============================
    
    '달의 좌표 얻기==============달은 이미 광차가 적용된 결과이므로 광차보정 불필요
    tpla.ipla = 11
    j = Plan404(tpla)
    MoonLon = tpla.l
    MoonLat = tpla.B
    MoonDist = tpla.R
    
    DX = MoonDist * Cos(MoonLon) * Cos(MoonLat)
    dy = MoonDist * Sin(MoonLon) * Cos(MoonLat)
    dz = MoonDist * Sin(MoonLat)
     '============================
     
     '달의 태양중심 황도 좌표
    xm = xe + DX
    ym = ye + dy
    zm = ze + dz
    
    SunLon = Rev(Arctan2d(ye, xe))    '지구의 태양중심 황도 좌표
    SunLat = Arctan2d(zg, Sqr(xe * xe + ye * ye))
    '===================================

CorTimeAbbr:  '광차 보정에는 해와 달 제외
    Select Case i
      Case 0  '해
          eln = 0: elt = 0: hDist = 0
          gDist = Sqr(xe * xe + ye * ye + ze + ze)

      Case 8 '달
          hDist = Sqr(xm * xm + ym * ym + zm + zm) 'tpla.R '* 23454.7800285036
    End Select

    '천체의 태양중심 황도 좌표(직교좌표):xh, yh, zh, (구면좌표): eln, elt
      xh = hDist * Cos(eln) * Cos(elt)
      yh = hDist * Sin(eln) * Cos(elt)
      zh = hDist * Sin(elt)

      '지심황도좌표(직교좌표):xg, yg, zg
      If i = 0 Then
        xg = xh - xe: yg = yh - ye: zg = zh - ze
      ElseIf i = 8 Then
        xg = xm - xe: yg = ym - ye: zg = zm - ze
      End If
      gDist = Sqr(xg * xg + yg * yg + zg * zg)  '지심 거리(AU)
      
      If TimeAbbr And Not CorTimeA Then  '시간에 따른 오차 보정하기
        Tc = gDist / c: julday = julday - Tc: CorTimeA = True
        GoTo CorTimeAbbr
      End If

      '지심황도좌표(구면좌표):glon, glat
      glon = Rev(Arctan2d(yg, xg))
      gLAT = Arctan2d(zg, Sqr(xg * xg + yg * yg))
      
      '지심적도좌표(직교좌표):xe, ye, ze
      xe = xg
      ye = yg * Cos(oblecl) - zg * Sin(oblecl)
      ze = yg * Sin(oblecl) + zg * Cos(oblecl)
      
      '지심적도좌표(구면좌표):eln, elt
      eln = Rev(Arctan2d(ye, xe))
      elt = Arctan2d(ze, Sqr(xe * xe + ye * ye))

    '물리적 특성
    elon = Arccosd((SunDist ^ 2 + gDist ^ 2 - hDist ^ 2) / (2 * SunDist * gDist))
    If i > 0 Then FV = Arccosd((gDist ^ 2 + hDist ^ 2 - SunDist ^ 2) / (2 * hDist * gDist))
    Select Case i
      Case 0
        PH = 1
      Case 8
        elon = Rev(Arccosd(Cosd((180 + SunLon) - glon) * Cosd(gLAT)))
        FV = 180 - elon
        PH = CSng((1 + Cosd(FV)) / 2)
        'fv: 이각,  위상(월광,%)= (1 + Cosd(FV)) / 2
    End Select

    pRA = eln
    pDE = elt
End Sub
