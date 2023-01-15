Attribute VB_Name = "modConvertLunar"
'Attribute VB_Name = "Lunisolar"
Option Explicit

'====================================================
'
'      음력-양력 상호 변환 함수
'
' *배포 버전: 1.01( VisualBasic 6, VBA)
' *제작자: 김창환(blueedu@hanmail.net)
' *홈페이지: http://blueedu.dothome.co.kr
' *제작일: 2009. 2. 19.
' *저작권: 이 프로그램의 저작권은 제작자에게 있습니다.
'
' 이 프로그램은 음력<->양력의 상호 변환을 위해 만들었습니다.
' 이 프로그램의 대부분은 달력 1.5 프로그램에서 쓰는 음양력 변환 코드와
' 같은 코드를 쓰고 있지만, 계산 속도의 향상을 위해 약간 다른 코드를
' 사용한 부분도 있습니다.
'
' 음양력 변환은 서기 1900년부터 서기 2100년까지의 기간동안 정확성을
' 보장할 수 있으며 확인해보지는 않았지만, 2000년 기준으로 전후 500년
' 정도의 기간 동안에는 별 문제없이 계산이 가능할 것으로 추정됩니다.
'
' 이 프로그램은 VisualBasic 6과 마이크로소프트 오피스 2000 이상에
' 포함되어 있는 VBA에 쓸 수 있습니다.
'
' <사용 방법>
' 이 프로그램에서는 다음의 두 함수를 써서 음양력 변환을 할 수 있습니다.
'
' 1) 양력->음력
'   양력에서 음력으로 변환할 때에는 Sol2Lun() 함수를 사용합니다.
'   Sol2Lun(년,월,일) : 매개변수는 모두 정수형, 양력 날짜를 입력
'   예] print Sol2Lun(1982,5,24)  '양력 1982년 5월 24일을 음력으로 바꿀 때
'   결과] 1982-4-2(윤)
'
' 2) 음력->양력
'   음력에서 양력으로 변환할 때에는 Lun2Sol() 함수를 사용합니다.
'   잘못된 값을 입력하면 빈 문자열을 반환합니다.
'   Lun2Sol(년,월,일,윤달 여부) : 날짜 매개변수는 모두 정수형, 음력 날짜를 입력.
'                                 윤달 여부는 논리형 변수, 윤달이면 참, 평달이면 거짓.
'   예] print Lun2Sol(1982,4,2,true)  '음력 1982년 윤 4월 2일을 양력으로 바꿀 때..
'   결과] 1982-5-24
'
'============================================================


Type Julgi12
  MonNumber As Byte
  Ref_Day As Double
  Longitude As Double
  RealDay As Double
End Type

Type LunarDay
  StartDay As Double
  MonLength As Integer
  Junggi As Boolean
  MonName As Byte
  LYear As Integer
End Type

Const pi As Double = 3.14159265358979
Const hpi As Double = 1.5707963267949
Const RadtoDeg As Double = 180 / pi
Const DegtoRad As Double = pi / 180
Const MoonMonth As Double = 29.5305882
Const MoonDay As Double = 12.190749387105
Const OneYear As Double = 365.24219
Const Oneday As Double = 360 / OneYear

Dim Junggi(15) As Julgi12, timezone As Double

'음양력 변환 입출력 함수입니다. 음양력의 출력 형태를 바꾸려면 아래에 있는 두 함수를 바꾸어주면 됩니다.
Public Function Sol2Lun(ByVal y As Integer, ByVal m As Integer, ByVal d As Integer) As String
  Dim ly As Integer, lm As Byte, LD As Byte, ll As Boolean, jd As Double
  
  jd = JULIANDAY(CDbl(y), CDbl(m), CDbl(d))
  If y > 1910 Then timezone = 9 Else timezone = 8
  Call LuniSolarCal(jd, ly, lm, LD, ll)
  Sol2Lun = Trim$(Str(lm)) & "." & Trim$(Str(LD)) & IIf(ll, "(윤)", "")
  'Sol2Lun = Trim$(Str(lm)) & "." & Trim$(Str(LD))

End Function

Public Function Lun2Sol(ByVal y As Integer, ByVal m As Integer, ByVal d As Integer, ByVal Leap As Boolean) As String
  Dim jd As Double, x As Boolean, a As Double, b As Double, c As Double
  
  If y > 1910 Then timezone = 9 Else timezone = 8
  x = InvLuniSolarCal(y, m, d, Leap, jd)
  If x Then
    InvJD jd, a, b, c
    'Lun2Sol = Trim$(Str(CInt(c)))
    Lun2Sol = Trim$(Str(CInt(a))) & "." & Trim$(Str(CInt(b))) & "." & Trim$(Str(CInt(c))) & "."
    'Lun2Sol = Trim$(Str(lm)) & "." & Trim$(Str(LD))
  Else
    Lun2Sol = ""
  End If
End Function

'아래부터는 음양력 변환을 위한 함수와 보조함수입니다.
Private Function InvLuniSolarCal(ByVal LunarYear As Integer, ByVal LunarMon As Byte, ByVal LunarDay As Byte, ByVal IsLeap As Boolean, jd As Double) As Boolean
  Dim iJD As Double, iLY As Integer, iLM As Byte, iLD As Byte, iLM2 As Single, iLeap As Boolean, i As Integer, D1 As Double, lm2 As Single, IsValid As Boolean
  Dim i2 As Integer
  
  IsValid = False
  
  iJD = JULIANDAY(CDbl(LunarYear), CDbl(LunarMon), CDbl(LunarDay))
  iJD = iJD + 25: lm2 = LunarMon
  If IsLeap Then iJD = iJD + 25: lm2 = lm2 + 0.5
  
  If LunarYear < 1582 Then iJD = iJD + Int(0.0078 * (1582 - LunarYear))

  i = 0
  Do
    LuniSolarCal iJD, iLY, iLM, iLD, iLeap
    
    If iLY > LunarYear Then
      iJD = iJD - 30
    ElseIf iLY < LunarYear Then
      iJD = iJD + 30
    ElseIf iLY = LunarYear Then
      i2 = 0
      Do
        LuniSolarCal iJD, iLY, iLM, iLD, iLeap
        iLM2 = iLM + IIf(iLeap, 0.5, 0)
        
        If iLM2 > lm2 Or iLY > LunarYear Then
          iJD = iJD - 10
        ElseIf iLM2 < lm2 Or iLY < LunarYear Then
          iJD = iJD + 10
        ElseIf iLM2 = lm2 Then
          D1 = iJD - iLD
          IsValid = True
        End If
        i2 = i2 + 1
      Loop Until iLM2 = lm2 Or i2 > 15
    End If
    i = i + 1
  Loop Until iLY = LunarYear Or i > 10
  
  jd = D1 + LunarDay
  InvLuniSolarCal = IsValid
End Function

Private Sub LuniSolarCal(ByVal jd As Double, LunarYear As Integer, LunarMon As Byte, LunarDay As Byte, IsLeap As Boolean)
  Dim jd0 As Double, dYear As Double, yJD0 As Double, bf As Double, b As Integer, m As Integer
  Dim i As Integer, j As Integer, k As Integer, LD(25) As LunarDay, SD(25) As Double
  Dim PreWinter As Double, ThisWinter As Double, Count1 As Integer, idx1 As Integer, idx2 As Integer
  Dim LeapType As Byte, Leap13 As Boolean, fMON As Integer, a As Integer, LCount As Integer
  
  jd0 = GetJD0(jd) + 0.5
  dYear = InvJDYear(jd0)
  
  If jd0 < cJunggi(dYear, 270, -13) Then dYear = dYear - 1
  If jd0 > cJunggi(dYear, 270, 355) Then dYear = dYear + 1
  
  yJD0 = JULIANDAY(dYear, 1, 1)
  Call Set24Julgi: CalcJulGi yJD0
  
  j = 0: SD(0) = 0: k = 0
  Do
    bf = GetJD0(NewMoon(yJD0 - 96 + j * 28)) + 0.5
    If bf >= Junggi(0).RealDay Then
      If k > 0 Then
        If bf > SD(k - 1) Then SD(k) = bf: k = k + 1
      Else
        SD(0) = bf: k = 1
      End If
    End If
    j = j + 1
  Loop Until bf > yJD0 + 427
  j = k: k = j - 1
  
  For i = 0 To 25
    LD(i).StartDay = 0
    LD(i).StartDay = SD(i)
    LD(i).MonName = 100
  Next i
  PreWinter = Junggi(1).RealDay
  ThisWinter = Junggi(13).RealDay
  
  idx1 = 0: idx2 = 0
  For i = 0 To 24
    LD(i).Junggi = False
    For j = 0 To 15
      If LD(i + 1).StartDay > Junggi(j).RealDay And Junggi(j).RealDay >= LD(i).StartDay Then
        LD(i).Junggi = True
        If LD(i).MonName = 100 Then
          LD(i).MonName = Junggi(j).MonNumber
        ElseIf Junggi(j).MonNumber = 11 Or LD(i).MonName = 11 Then
          LD(i).MonName = 11
        End If
      End If
    Next j
  Next i
  
  Count1 = 0
  For i = 0 To k - 1
    If PreWinter < LD(i).StartDay And LD(i).StartDay <= ThisWinter Then Count1 = Count1 + 1
    If PreWinter < LD(i + 1).StartDay And LD(i).StartDay <= PreWinter Then idx1 = i
    If ThisWinter < LD(i + 1).StartDay And LD(i).StartDay <= ThisWinter Then idx2 = i
  Next i

  If Count1 = 12 Then
    LeapType = 4
  Else
    If LD(idx1 + 1).Junggi = True And LD(idx1 + 2).Junggi = True Then LeapType = 1
    If LD(idx1 + 1).Junggi = False And LD(idx1 + 2).Junggi = True Then LeapType = 2
    If LD(idx1 + 1).Junggi = True And LD(idx1 + 2).Junggi = False Then LeapType = 3
    
    Leap13 = False
    For i = (idx1 + 3) To (idx2 - 1)
      If LD(i).Junggi = False Then Leap13 = True
    Next i
    If LeapType = 2 Or LeapType = 3 Then Leap13 = False
  End If

  
  LCount = 0
  Select Case LeapType
   Case 1, 4
     LD(idx1 + 1).MonName = 12: LD(idx1 + 1).LYear = dYear - 1: LD(idx1 + 1).Junggi = True
     LD(idx1 + 2).MonName = 1: LD(idx1 + 2).LYear = dYear: LD(idx1 + 2).Junggi = True
     
   Case 2
      LD(idx1 + 1).MonName = 11: LD(idx1 + 1).LYear = dYear - 1: LD(idx1 + 1).Junggi = False
      LD(idx1 + 2).MonName = 12: LD(idx1 + 2).LYear = dYear - 1
     
   Case 3
     LD(idx1 + 1).MonName = 12: LD(idx1 + 1).LYear = dYear - 1
     LD(idx1 + 2).MonName = 12: LD(idx1 + 2).LYear = dYear - 1: LD(idx1 + 1).Junggi = False
  
  End Select
  LD(idx1).MonName = 11: LD(idx1).LYear = dYear - 1
  
  fMON = 1: a = 0
  If LeapType = 4 Then a = 1
  For i = idx1 + 3 To idx2
    LD(i).LYear = dYear
    
    If LeapType = 1 Then
      If LD(i).Junggi = True Or LCount > 0 Then
        LD(i).Junggi = True
        a = a + 1
      Else
        LCount = 1
      End If
      LD(i).MonName = fMON + a
    
    Else
      LD(i).MonName = fMON + a
      LD(i).Junggi = True
      a = a + 1
    End If
  Next i
  
  For i = 0 To 24
    If Abs(LD(i + 1).StartDay - LD(i).StartDay) < 31 Then
      LD(i).MonLength = LD(i + 1).StartDay - LD(i).StartDay
    End If
  Next i

  For i = 0 To k - 1
    If jd0 >= LD(i).StartDay And jd0 < LD(i + 1).StartDay Then
      LunarYear = LD(i).LYear
      LunarMon = LD(i).MonName
      LunarDay = jd0 - LD(i).StartDay + 1
      IsLeap = Not LD(i).Junggi
      Exit For
    End If
  Next i
End Sub

Private Sub CalcJulGi(ByVal JDt As Double)
  Dim i As Integer, nYear As Double
  
  nYear = InvJDYear(JDt)
  
  For i = 0 To 15
    Junggi(i).RealDay = GetJD0(cJunggi(nYear, Junggi(i).Longitude, Junggi(i).Ref_Day)) + 0.5
  Next i
End Sub

Private Function cJunggi(ByVal cYear As Double, ByVal LonSun As Double, ByVal RefDay As Double) As Double
  Dim JDyear As Double, aDay As Double
  Dim LamSun As Double, dt As Double, dLam As Double, tJD As Double, i As Long
  Dim dl As Double, SL As Double, SB As Double, SR As Double, JDE As Double
  
  dt = 0: i = 0
  JDyear = JULIANDAY(cYear, 1, 0) - 0.5
  
  tJD = JDyear + RefDay
  JDE = JDtoTDT(tJD)
  SUN JDE, SL, SB, SR
  LamSun = Rev(SL - 0.005691611 / SR)
  dLam = AngDistLon(LamSun, LonSun)
  
  Do
    dt = dLam / Oneday
    
    If LonSun > 357 Or LonSun < 3 Then
      If LamSun > 180 Then LamSun = LamSun - 360
    End If
    
    If LonSun > LamSun Then
      tJD = tJD + dt
    Else
      tJD = tJD - dt
    End If
    
    JDE = JDtoTDT(tJD)
    SUN JDE, SL, SB, SR
    Nutation JDE, dl
    LamSun = Rev(SL + dl / 3600 - 0.005691611 / SR)
    
    dLam = AngDistLon(LamSun, LonSun)
    i = i + 1
  Loop Until (dLam / Oneday * 86400) < 1 Or i > 50
  cJunggi = tJD + timezone / 24
End Function

Private Sub Set24Julgi()
  Junggi(0).MonNumber = 10: Junggi(0).Longitude = 240: Junggi(0).Ref_Day = -43
  Junggi(1).MonNumber = 11: Junggi(1).Longitude = 270: Junggi(1).Ref_Day = -13
  Junggi(2).MonNumber = 12: Junggi(2).Longitude = 300: Junggi(2).Ref_Day = 20
  Junggi(3).MonNumber = 1: Junggi(3).Longitude = 330: Junggi(3).Ref_Day = 50
  Junggi(4).MonNumber = 2: Junggi(4).Longitude = 0: Junggi(4).Ref_Day = 80
  Junggi(5).MonNumber = 3: Junggi(5).Longitude = 30: Junggi(5).Ref_Day = 110
  Junggi(6).MonNumber = 4: Junggi(6).Longitude = 60: Junggi(6).Ref_Day = 140
  Junggi(7).MonNumber = 5: Junggi(7).Longitude = 90: Junggi(7).Ref_Day = 170
  Junggi(8).MonNumber = 6: Junggi(8).Longitude = 120: Junggi(8).Ref_Day = 200
  Junggi(9).MonNumber = 7: Junggi(9).Longitude = 150: Junggi(9).Ref_Day = 230
  Junggi(10).MonNumber = 8: Junggi(10).Longitude = 180: Junggi(10).Ref_Day = 260
  Junggi(11).MonNumber = 9: Junggi(11).Longitude = 210: Junggi(11).Ref_Day = 290
  Junggi(12).MonNumber = 10: Junggi(12).Longitude = 240: Junggi(12).Ref_Day = 320
  Junggi(13).MonNumber = 11: Junggi(13).Longitude = 270: Junggi(13).Ref_Day = 350
  Junggi(14).MonNumber = 12: Junggi(14).Longitude = 300: Junggi(14).Ref_Day = 385
  Junggi(15).MonNumber = 1: Junggi(15).Longitude = 330: Junggi(15).Ref_Day = 415
End Sub

Private Function Sind(ByVal x As Double) As Double
    Sind = Sin(x * DegtoRad)
End Function

Private Function Cosd(ByVal x As Double) As Double
    Cosd = Cos(x * DegtoRad)
End Function

Private Function Tand(ByVal x As Double) As Double
    Tand = Tan(x * DegtoRad)
End Function

Private Function Arccosd(ByVal x As Double) As Double
    If x <= -1 Then
      Arccosd = 180
    ElseIf x < 1 And x > -1 Then
      Arccosd = 90 - RadtoDeg * Atn(x / Sqr(1 - x * x))
    Else
      Arccosd = 0
    End If
End Function

Private Function Rev(ByVal x As Double) As Double
    Rev = x - Int(x / 360) * 360
End Function

Private Function AngDistLon(ByVal RA1 As Double, ByVal RA2 As Double) As Double
  If RA1 = RA2 Then
    AngDistLon = 0
  Else
    AngDistLon = Arccosd(Cosd(RA1 - RA2))
  End If
End Function

Private Sub Nutation(ByVal JDE As Double, dPsi As Double)
  Dim eps0 As Double
  Dim OM As Double, t As Double, L1 As Double, L2 As Double, T2 As Double
  
  t = (JDE - 2451545) / 36525
  T2 = t * t
  L1 = Rev(280.466457 + 36000.7698278 * t + 0.00030322 * T2 + 0.00000002 * T2 * t)
  L2 = Rev(218.3164477 + 481267.88123421 * t - 0.0015786 * T2 + T2 * t / 538841 - T2 * T2 / 65194000)
  OM = Rev(125.04452 - 1934.136261 * t + 0.0020708 * T2 + (T2 * t) / 450000)
  
  dPsi = -17.2 * Sind(OM) - 1.32 * Sind(2 * L1) - 0.23 * Sind(2 * L2) + 0.21 * Sind(2 * OM)
End Sub

Private Function JULIANDAY(ByVal year As Double, ByVal Month As Double, ByVal Day As Double) As Double
    Dim ggg As Double, S As Double, a As Double, J1 As Double, tJD As Double
    
    ggg = 1
    If year < 1582 Then ggg = 0
    If year <= 1582 And Month < 10 Then ggg = 0
    If year <= 1582 And Month = 10 And Day < 5 Then ggg = 0
    tJD = -1 * Int(7 * (Int((Month + 9) / 12) + year) / 4)
    S = 1
    If (Month - 9) < 0 Then S = -1
    a = Abs(Month - 9)
    J1 = Int(year + S * Int(a / 7))
    J1 = -1 * Int((Int(J1 / 100) + 1) * 3 / 4)
    tJD = tJD + Int(275 * Month / 9) + Day + (ggg * J1)
    tJD = tJD + 1721027 + 2 * ggg + 367 * year
    tJD = tJD
    JULIANDAY = tJD
End Function

Private Sub InvJD(ByVal julday As Double, year As Double, Month As Double, Day As Double)
    Dim Z As Double, F As Double, a As Double, i As Double, b As Double
    Dim c As Double, d As Double, t As Double, rj As Double, jj As Double, RH As Double
    Dim mMe As Double, aae As Double
    
    Z = Int(julday + 0.5)
    F = julday + 0.5 - Z
    If Z < 2299161 Then
       a = Z
    Else
       i = Int((Z - 1867216.25) / 36524.25)
       a = Z + 1 + i - Int(i / 4)
    End If
    
    b = a + 1524
    c = Int((b - 122.1) / 365.25)
    d = Int(365.25 * c)
    t = Int((b - d) / 30.6)
    rj = b - d - Int(30.6001 * t) + F
    jj = Int(rj)
    
    If t < 14 Then
       mMe = t - 1
    Else
       If t = 14 Or t = 15 Then mMe = t - 13
    End If
    
    If mMe > 2 Then
       aae = c - 4716
    ElseIf mMe = 1 Or mMe = 2 Then
       aae = c - 4715
    End If
    
    year = aae: Month = mMe: Day = jj
End Sub

Private Function InvJDYear(ByVal julday As Double) As Double
    Dim Z As Double, a As Double, i As Double, b As Double
    Dim c As Double, d As Double, t As Double, mMe As Double
    
    Z = Int(julday + 0.5)
    
    If Z < 2299161 Then
       a = Z
    Else
       i = Int((Z - 1867216.25) / 36524.25)
       a = Z + 1 + i - Int(i / 4)
    End If
    
    b = a + 1524
    c = Int((b - 122.1) / 365.25)
    d = Int(365.25 * c)
    t = Int((b - d) / 30.6)
    
    If t < 14 Then
       mMe = t - 1
    Else
       If t = 14 Or t = 15 Then mMe = t - 13
    End If
    
    If mMe > 2 Then
       InvJDYear = c - 4716
    ElseIf mMe = 1 Or mMe = 2 Then
       InvJDYear = c - 4715
    End If
End Function

Private Function GetJD0(ByVal DateJD As Double) As Double
  If DateJD - Int(DateJD) >= 0.5 Then
    GetJD0 = Int(DateJD) + 0.5
  Else
    GetJD0 = Int(DateJD) - 0.5
  End If
End Function

Private Function JDtoTDT(ByVal julday As Double) As Double
  Dim y As Long, t As Double, dt As Double
  
  y = CInt(InvJDYear(julday))
  Select Case y
   Case Is < 949
     t = (y - 2000) / 100
     dt = (2715.6 + 573.36 * t + 46.5 * t * t) / 3600
   Case 949 To 1619
     t = (y - 1850) / 100
     dt = (22.5 * t * t) / 3600
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
     t = (y - 1900) / 100
     dt = (-0.067471 + t * (-0.058091))
     dt = (0.161416 + t * (0.145932 + t * dt))
     dt = (-0.14696 + t * (-0.146279 + t * dt))
     dt = (0.062971 + t * (0.079441 + t * dt))
     dt = (-0.012462 + t * (-0.022542 + t * dt))
     dt = (-0.000014 + t * (0.001148 + t * (0.003357 + t * dt)))
     dt = dt * 24
   Case 1900 To 1987
     t = (y - 1900) / 100
     dt = (-0.861938 + t * (0.677066 + t * -0.212591))
     dt = (0.025184 + t * (-0.181133 + t * (0.55304 + t * dt)))
     dt = (-0.00002 + t * (0.000297 + t * dt))
     dt = dt * 24
   Case 1988 To 1996
     t = (y - 2000) / 100
     dt = (67 + 123.5 * t + 32.5 * t * t) / 3600
   Case 1997: dt = 62 / 3600
   Case 1998 To 1999: dt = 63 / 3600
   Case 2000 To 2001: dt = 64 / 3600
   Case 2002 To 2020
     t = (y - 2000) / 100
     dt = (63 + 123.5 * t + 32.5 * t * t) / 3600
   Case Is > 2020
     t = (y - 1875.1) / 100
     dt = 45.39 * t * t / 3600
   Case Else
     dt = 0
  End Select
  
  JDtoTDT = dt / 24 + julday
End Function

Private Sub ANOMALY(ByVal AM As Double, ByVal EC As Double, AT As Double, AE As Double)
  Dim a As Double, E1 As Double, E0 As Double
  
  AM = Rev(AM) * DegtoRad: E1 = AM
  
  Do
    E0 = E1
    E1 = E0 + (AM + EC * Sin(E0) - E0) / (1 - EC * Cos(E0))
  Loop Until Abs(E1 - E0) < 0.000001
  AE = E1
  
  a = Sqr((1 + EC) / (1 - EC)) * Tan(AE / 2)
  AT = Rev(2 * Atn(a) * RadtoDeg)
  AE = Rev(AE * RadtoDeg)
End Sub

Private Sub SUN(ByVal jd As Double, SL As Double, SB As Double, SR As Double)
  Dim t As Double, T2 As Double, a As Double, b As Double, L As Double, M1 As Double, EC As Double
  Dim A1 As Double, B1 As Double, C1 As Double, D1 As Double, E1 As Double, H1 As Double
  Dim D2 As Double, D3 As Double, AT As Double, AE As Double
  
  t = (jd - 2415020) / 36525: T2 = t * t
 
  a = 100.0021359 * t: b = 360 * (a - Int(a))
  L = 279.69668 + 0.0003025 * T2 + b
  a = 99.99736042 * t: b = 360 * (a - Int(a))
  M1 = 358.47583 - (0.00015 + 0.0000033 * t) * T2 + b
  EC = 0.01675104 - 0.0000418 * t - 0.000000126 * T2
  
  Call ANOMALY(M1, EC, AT, AE)
  
  a = 62.55209472 * t: b = 360 * (a - Int(a))
  A1 = 153.23 + b
  a = 125.1041894 * t: b = 360 * (a - Int(a))
  B1 = 216.57 + b
  a = 91.56766028 * t: b = 360 * (a - Int(a))
  C1 = 312.69 + b
  a = 1236.853095 * t: b = 360 * (a - Int(a))
  D1 = 350.74 - 0.00144 * T2 + b
  E1 = 231.19 + 20.2 * t
  a = 183.1353208 * t: b = 360 * (a - Int(a))
  H1 = 353.4 + b
  
  D2 = 0.00134 * Cosd(A1) + 0.00154 * Cosd(B1) + 0.002 * Cosd(C1)
  D2 = D2 + 0.00179 * Sind(D1) + 0.00178 * Sind(E1)
  
  D3 = 0.00000543 * Sind(A1) + 0.00001575 * Sind(B1)
  D3 = D3 + 0.00001627 * Sind(C1) + 0.00003076 * Cosd(D1)
  D3 = D3 + 0.00000927 * Sind(H1)
  
  SL = Rev(AT + L - M1 + D2)
  SR = 1.0000002 * (1 - EC * Cosd(AE)) + D3
  SB = 0
End Sub

Private Function NewMoon(ByVal JDN As Double) As Double
  Dim jd As Double, cor As Double, m As Double, Mp As Double, F As Double, t As Double, OMG As Double
  Dim YR As Double, k As Double, nph As Double, dt As Double, T2 As Double, T3 As Double, T4 As Double
  Dim A1 As Double, A2 As Double, A3 As Double, A4 As Double, A5 As Double, A6 As Double, A7 As Double
  Dim A8 As Double, A9 As Double, A10 As Double, E As Double, E2 As Double, TX As Double
  Dim A11 As Double, A12 As Double, A13 As Double, A14 As Double, cor2 As Double

  YR = InvJDYear(JDN)
  k = JDN - JULIANDAY(YR, 1, 1)
  YR = YR + k / 365.25
  nph = (YR - 2000) * 12.3685
  nph = Int(nph)
  k = nph: t = k / 1236.85
  T2 = t * t: T3 = T2 * t: T4 = T2 * T2

  jd = 2451550.09766 + 29.530588861 * k + 0.00015437 * T2 - 0.00000015 * T3 + 0.00033 * 0.00000000073 * T4
       
  TX = (jd - 2451545) / 36525
  E = 1 - 0.002516 * TX - 0.0000074 * TX * TX: E2 = E * E
  m = Rev(2.5534 + 29.1053567 * k - 0.0000014 * T2 - 0.00000011 * T3)
  Mp = Rev(201.5643 + 385.81693528 * k + 0.0107582 * T2 + 0.00001238 * T3 - 0.000000058 * T4)
  F = Rev(160.7108 + 390.67050284 * k - 0.0016118 * T2 - 0.00000227 * T3 + 0.000000011 * T4)
  OMG = Rev(124.7746 - 1.56375588 * k + 0.0020672 * T2 + 0.00000215 * T3)
  
  A1 = 299.77 + 0.107408 * k - 0.009173 * T2
  A2 = 251.88 + 0.016321 * k
  A3 = 251.83 + 26.651886 * k
  A4 = 349.42 + 36.412478 * k
  A5 = 84.66 + 18.206239 * k
  A6 = 141.74 + 53.303771 * k
  A7 = 207.14 + 2.453732 * k
  A8 = 154.84 + 7.30686 * k
  A9 = 34.52 + 27.261239 * k
  A10 = 207.19 + 0.121824 * k
  A11 = 291.34 + 1.844379 * k
  A12 = 161.72 + 24.198154 * k
  A13 = 239.56 + 25.513099 * k
  A14 = 331.55 + 3.592518 * k
  
  cor = -0.4072 * Sind(Mp) + 0.17241 * E * Sind(m) + 0.01608 * Sind(2 * Mp) _
        + 0.01039 * Sind(2 * F) + 0.00739 * E * Sind(Mp - m) - 0.00514 * E * Sind(Mp + m) _
        + 0.00208 * E2 * Sind(2 * m) - 0.00111 * Sind(Mp - 2 * F) - 0.00057 * Sind(Mp + 2 * F) _
        + 0.00056 * E * Sind(2 * Mp + m) - 0.00042 * Sind(3 * Mp) + 0.00042 * E * Sind(m + 2 * F) _
        + 0.00038 * E * Sind(m - 2 * F) - 0.00024 * E * Sind(2 * Mp - m) - 0.00017 * Sind(OMG) _
        - 0.00007 * Sind(Mp + 2 * m) + 0.00004 * Sind(2 * Mp - 2 * F) + 0.00004 * Sind(3 * m) _
        + 0.00003 * Sind(Mp + m - 2 * F) + 0.00003 * Sind(2 * Mp + 2 * F) - 0.00003 * Sind(Mp + m + 2 * F) _
        + 0.00003 * Sind(Mp - m + 2 * F) - 0.00002 * Sind(Mp - m - 2 * F) - 0.00002 * Sind(3 * Mp + m) + 0.00002 * Sind(4 * Mp)
  
  cor2 = 0.000325 * Sind(A1) + 0.000165 * Sind(A2) + 0.000164 * Sind(A3) + 0.000126 * Sind(A4) _
       + 0.00011 * Sind(A5) + 0.000062 * Sind(A6) + 0.00006 * Sind(A7) + 0.000056 * Sind(A8) _
       + 0.000047 * Sind(A9) + 0.000042 * Sind(A10) + 0.00004 * Sind(A11) + 0.000037 * Sind(A12) _
       + 0.000035 * Sind(A13) + 0.000023 * Sind(A14)
      
  jd = jd + cor + cor2
  dt = JDtoTDT(jd) - jd
  NewMoon = jd - dt + timezone / 24
End Function












