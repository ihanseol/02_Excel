Attribute VB_Name = "m_MeanPos"
Option Explicit

Public Function SunML(ByVal JDE As Double) As Double
  Dim L0 As Double, T As Double, T2 As Double, T3 As Double
  
  T = (JDE - 2451545) / 365250: T2 = T * T: T3 = T2 * T
  L0 = 280.4664567 + 360007.6982779 * T + 0.03032028 * T2 + T3 / 49931 - (T2 * T2) / 15300 - (T3 * T2) / 2000000
  SunML = Rev(L0)
End Function

Public Function MoonML(ByVal JDE As Double) As Double
  Dim L0 As Double, T As Double, T2 As Double, T3 As Double
  
  T = (JDE - 2451545) / 36525: T2 = T * T: T3 = T2 * T
  L0 = 218.3164477 + 481267.88123421 * T - 0.0015786 * T2 + T3 / 538841 - (T2 * T2) / 65194000
  MoonML = Rev(L0)
End Function

'평기법
Public Function Pyunggi(ByVal cYear As Double, ByVal LonSun As Double, ByVal RefDay As Double, ByVal TZone As Double) As Double
  Dim JDyear As Double, TDT As Double, LamSun As Double, dt As Double, dLam As Double, tJD As Double, i As Long
  
  dt = 0: i = 0
  If cYear < 1582 Then RefDay = RefDay + Int(0.0078 * (1582 - cYear))
  JDyear = JULIANDAY(cYear, 1, 0, 0, 0, 0)
  tJD = JDyear + RefDay
  
  Do
    TDT = JDtoTDT(tJD)
    LamSun = SunML(TDT) - 0.0057183
    
    dLam = LonSun - LamSun
    If dLam < 0 Then dLam = dLam + 360
    If dLam > 180 Then dLam = dLam - 360
    
    dt = dLam / 0.985647359085214
    tJD = tJD + dt
    
    i = i + 1
  Loop Until (Abs(dLam) / 0.985647359085214 * 86400) < 1 Or i > 10  '시간 오차: 1초 이내
  Pyunggi = tJD + TZone / 24
End Function

'평삭법
Public Function GetMeanMoon(ByVal cJD As Double, ByVal TZone As Double) As Double
  Dim LamSun As Double, dt As Double, dLam As Double, tJD As Double
  Dim LamMoon As Double, i As Long, TDT As Double
  
  dt = 0: i = 0: tJD = cJD

  Do
    TDT = JDtoTDT(tJD)
    LamSun = SunML(TDT) - 0.0057183
    LamMoon = MoonML(TDT)
    
    dLam = LamSun - LamMoon
    If dLam < 0 Then dLam = dLam + 360
    If dLam > 180 Then dLam = dLam - 360
    
    dt = dLam / 12.190749387105
    tJD = tJD + dt
  
    i = i + 1
  Loop Until (Abs(dLam) / 12.190749387105 * 86400) < 1 Or i > 10  '시간 오차: 1초 이내
  
  GetMeanMoon = tJD + TZone / 24
End Function

Public Sub AutoChoose(ByVal LYear As Integer)
  Select Case LYear
   Case Is < 619
     UseMeanSun = True: UseMeanMoon = True: UseJinsak = False
   Case 619 To 664
     UseMeanSun = True: UseMeanMoon = False: UseJinsak = False
   Case 665 To 1280
     UseMeanSun = True: UseMeanMoon = False: UseJinsak = True
   Case 1281 To 1644
     UseMeanSun = True: UseMeanMoon = False: UseJinsak = False
   Case Is > 1644
     UseMeanSun = False: UseMeanMoon = False: UseJinsak = False
  End Select
End Sub

'Public Sub AutoChooseKR(ByVal LYear As Integer)
'  Select Case LYear
'   Case Is < 624
'     UseMeanSun = True: UseMeanMoon = True: UseJinsak = False
'   Case 625 To 673
'     UseMeanSun = True: UseMeanMoon = False: UseJinsak = False
'   Case 674 To 1309
'     UseMeanSun = True: UseMeanMoon = False: UseJinsak = True
'   Case 1310 To 1652
'     UseMeanSun = True: UseMeanMoon = False: UseJinsak = False
'   Case Is > 1652
'     UseMeanSun = False: UseMeanMoon = False: UseJinsak = False
'  End Select
'End Sub

