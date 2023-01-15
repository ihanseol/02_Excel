Attribute VB_Name = "m_Math"
Option Explicit

'상수
Public Const pi As Double = 3.14159265358979
Public Const hpi As Double = 1.5707963267949
Public Const RadtoDeg As Double = 180 / pi
Public Const DegtoRad As Double = pi / 180

'삼각 함수
Public Function Arcsin(ByVal X As Double) As Double
    Arcsin = Atn(X / Sqr(1 - X * X))
End Function
Public Function Arccos(ByVal X As Double) As Double
  If X <= -1 Then
    Arccos = pi
  ElseIf X < 1 And X > -1 Then
    Arccos = hpi - Atn(X / Sqr(1 - X * X))
  Else
    Arccos = 0
  End If
End Function
Public Function Sind(ByVal X As Double) As Double
    Sind = Sin(X * DegtoRad)
End Function
Public Function Cosd(ByVal X As Double) As Double
    Cosd = Cos(X * DegtoRad)
End Function
Public Function Tand(ByVal X As Double) As Double
    Tand = Tan(X * DegtoRad)
End Function
Public Function Arcsind(ByVal X As Double) As Double
    If X >= 1 Then
      Arcsind = 90
    ElseIf X < 1 Then
      Arcsind = RadtoDeg * Atn(X / Sqr(1 - X * X))
    End If
End Function
Public Function Arccosd(ByVal X As Double) As Double
    If X <= -1 Then
      Arccosd = 180
    ElseIf X < 1 And X > -1 Then
      Arccosd = 90 - RadtoDeg * Atn(X / Sqr(1 - X * X))
    Else
      Arccosd = 0
    End If
End Function
Public Function Arctand(ByVal X As Double) As Double
    Arctand = RadtoDeg * Atn(X)
End Function
Public Function Rev(ByVal X As Double) As Double
    Rev = X - Int(X / 360) * 360
End Function
Public Function Arctan2(ByVal y As Double, ByVal X As Double) As Double
    If X = 0 Then
      If y = 0 Then
        'error
      ElseIf y > 0 Then
        X = hpi
      Else
        X = -hpi
      End If
    Else
      If X > 0 Then
        X = Atn(y / X)
      ElseIf X < 0 Then
        If y >= 0 Then
          X = Atn(y / X) + pi
        Else
          X = Atn(y / X) - pi
        End If
      End If
    End If
    'Arctan2 = Atn(y / x) - pi * (x < 0)
    Arctan2 = X
End Function
Public Function Arctan2d(ByVal y As Double, ByVal X As Double) As Double
    If X = 0 Then
      If y = 0 Then
        'error
      ElseIf y > 0 Then
        X = hpi
      Else
        X = -hpi
      End If
    Else
      If X > 0 Then
        X = Atn(y / X)
      ElseIf X < 0 Then
        If y >= 0 Then
          X = Atn(y / X) + pi
        Else
          X = Atn(y / X) - pi
        End If
      End If
    End If
    Arctan2d = X * RadtoDeg
End Function


Public Function Log10(ByVal X As Double) As Double
    If X < 0 Then X = -X
    Log10 = Log(X) / 2.30258509299405 ' Log(10)=2.30258509299405
End Function



'Public Function Rmod(X As Double, Range As Double) As Double
'    Rmod = X - Int(X / Range) * Range
'End Function

'3개의 항을 이용한 보간법
'J. Meeus, Astronomical Algorithms(2nd edition), 1998 23~24쪽에서
Public Function Inter3(ByVal Y1 As Double, ByVal Y2 As Double, ByVal y3 As Double, ByVal N As Double) As Double
  Dim a As Double, B As Double, c As Double
  
  a = Y2 - Y1
  B = y3 - Y2
  c = Y1 + y3 - 2 * Y2
  
  Inter3 = Y2 + (N / 2) * (a + B + N * c)
End Function

Public Sub Inter3Sph(ByVal A1 As Double, ByVal B1 As Double, ByVal A2 As Double, ByVal B2 As Double, _
                                               ByVal A3 As Double, ByVal b3 As Double, ByVal N As Double, resLon As Double, resLat As Double)
  Dim X1 As Double, x2 As Double, x3 As Double
  Dim Y1 As Double, Y2 As Double, y3 As Double
  Dim Z1 As Double, Z2 As Double, z3 As Double
  Dim rx As Double, ry As Double, rz As Double, RA As Double, rb As Double
  
  SphToRect A1, B1, X1, Y1, Z1
  SphToRect A2, B2, x2, Y2, Z2
  SphToRect A3, b3, x3, y3, z3
  
  rx = Inter3(X1, x2, x3, N)
  ry = Inter3(Y1, Y2, y3, N)
  rz = Inter3(Z1, Z2, z3, N)
  
  RTS_Real rx, ry, rz, RA, rb
  resLon = RA: resLat = rb
End Sub

'직교좌표에서 구면좌표로(degree)
Public Sub RectToSph(ByVal X As Double, ByVal y As Double, ByVal Z As Double, a As Double, B As Double)
  a = Rev(Arctan2d(y, X))
  B = Arcsind(Z / 1000#)
  '또는 b = Arcsind(Z / Sqr(X * X + y * y + Z * Z))
End Sub

'구면좌표에서 직교좌표로(degree)
Public Sub SphToRect(ByVal a As Double, ByVal B As Double, X As Double, y As Double, Z As Double)
  a = a * DegtoRad: B = B * DegtoRad
  X = 1000# * Cos(a) * Cos(B)
  y = 1000# * Sin(a) * Cos(B)
  Z = 1000# * Sin(B)
End Sub

Public Sub RTS_Real(ByVal X As Double, ByVal y As Double, ByVal Z As Double, a As Double, B As Double)
  a = Rev(Arctan2d(y, X))
  B = Arcsind(Z / Sqr(X * X + y * y + Z * Z))
End Sub

Public Function RotX(X As Double, y As Double, Z As Double, ByVal Rot As Double)
    Dim Y1 As Double, Z1 As Double
    Y1 = y: Z1 = Z
    Rot = Rot * DegtoRad
    y = Y1 * Cos(Rot) + Z1 * Sin(Rot)
    Z = -Y1 * Sin(Rot) + Z1 * Cos(Rot)
End Function

Public Function RotY(X As Double, y As Double, Z As Double, ByVal Rot As Double)
    Dim X1 As Double, Z1 As Double
    X1 = X: Z1 = Z
    Rot = Rot * DegtoRad
    X = X1 * Cos(Rot) - Z1 * Sin(Rot)
    Z = X1 * Sin(Rot) + Z1 * Cos(Rot)
End Function

Public Function RotZ(X As Double, y As Double, Z As Double, ByVal Rot As Double)
    Dim X1 As Double, Y1 As Double
    X1 = X: Y1 = y
    Rot = Rot * DegtoRad
    X = X1 * Cos(Rot) + Y1 * Sin(Rot)
    y = -X1 * Sin(Rot) + Y1 * Cos(Rot)
End Function
