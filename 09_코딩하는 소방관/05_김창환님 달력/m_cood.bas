Attribute VB_Name = "m_cood"
Option Explicit

Public Function AngDistLon(ByVal RA1 As Double, ByVal RA2 As Double) As Double
  If RA1 = RA2 Then
    AngDistLon = 0
  Else
    AngDistLon = Arccosd(Cosd(RA1 - RA2))
  End If
End Function

Public Sub EquToAltAz(ByVal RA As Double, ByVal DEC As Double, ByVal LST As Double, ByVal Latitude As Double, az As Double, ALT As Double)
   Dim X As Double, y As Double, Z As Double
   X = Cosd(RA) * Cosd(DEC)
   y = Sind(RA) * Cosd(DEC)
   Z = Sind(DEC)
   
   RotZ X, y, Z, LST * 15 + 180
   RotY X, y, Z, Latitude
   
   az = Rev(-Arctan2d(y, Z))
   ALT = Arcsind(-X)
End Sub

Public Sub Nutation(ByVal JDE As Double, dPsi As Double, dEps As Double)
  Dim eps0 As Double
  Dim OM As Double, T As Double, L1 As Double, L2 As Double, T2 As Double
  
  T = (JDE - 2451545) / 36525
  T2 = T * T
  L1 = Rev(280.466457 + 36000.7698278 * T + 0.00030322 * T2 + 0.00000002 * T2 * T)  '태양의 평균 황경
  L2 = Rev(218.3164477 + 481267.88123421 * T - 0.0015786 * T2 + T2 * T / 538841 - T2 * T2 / 65194000) '달의 평균 황경
  OM = Rev(125.04452 - 1934.136261 * T + 0.0020708 * T2 + (T2 * T) / 450000)  '달의 승교점 경도
  
  dPsi = -17.2 * Sind(OM) - 1.32 * Sind(2 * L1) - 0.23 * Sind(2 * L2) + 0.21 * Sind(2 * OM)
  dEps = 9.2 * Cosd(OM) + 0.57 * Cosd(2 * L1) + 0.1 * Cosd(2 * L2) - 0.09 * Cosd(2 * OM)
End Sub
