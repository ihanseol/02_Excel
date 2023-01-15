Attribute VB_Name = "m_Time"
Option Explicit

'2000.0년 기준


Public Function JULIANDAY(ByVal Year As Double, ByVal Month As Double, ByVal Day As Double, ByVal hour As Double, ByVal min As Double, Optional ByVal Sec As Double = 0) As Double     '율리우스 적일
    Dim ggg As Double, S As Double, a As Double, J1 As Double, tJD As Double, thr As Double
    
    thr = hour + min / 60 + Sec / 3600
    ggg = 1
    If Year < 1582 Then ggg = 0
    If Year <= 1582 And Month < 10 Then ggg = 0
    If Year <= 1582 And Month = 10 And Day < 5 Then ggg = 0
    tJD = -1 * Int(7 * (Int((Month + 9) / 12) + Year) / 4)
    S = 1
    If (Month - 9) < 0 Then S = -1
    a = Abs(Month - 9)
    J1 = Int(Year + S * Int(a / 7))
    J1 = -1 * Int((Int(J1 / 100) + 1) * 3 / 4)
    tJD = tJD + Int(275 * Month / 9) + Day + (ggg * J1)
    tJD = tJD + 1721027 + 2 * ggg + 367 * Year - 0.5
    tJD = tJD + (thr / 24)
    JULIANDAY = tJD
End Function

Public Sub InvJD(ByVal julday As Double, Year As Double, Month As Double, Day As Double, hour As Double, min As Double, Optional Sec As Double)
    Dim Z As Double, f As Double, a As Double, i As Double, B As Double
    Dim c As Double, D As Double, T As Double, rj As Double, JJ As Double, RH As Double
    Dim hRe As Double, mNe As Double, sn As Double, mMe As Double, aae As Double
    
    Z = Int(julday + 0.5)
    f = julday + 0.5 - Z
    If Z < 2299161 Then
       a = Z
    Else
       i = Int((Z - 1867216.25) / 36524.25)
       a = Z + 1 + i - Int(i / 4)
    End If
    
    B = a + 1524
    c = Int((B - 122.1) / 365.25)
    D = Int(365.25 * c)
    T = Int((B - D) / 30.6)
    rj = B - D - Int(30.6001 * T) + f
    JJ = Int(rj)
    RH = (rj - Int(rj)) * 24
    hRe = Int(RH)
    mNe = Int((RH - hRe) * 60)
    sn = Int(((RH - hRe) * 60 - mNe) * 60)
    
    If T < 14 Then
       mMe = T - 1
    Else
       If T = 14 Or T = 15 Then mMe = T - 13
    End If
    
    If mMe > 2 Then
       aae = c - 4716
    ElseIf mMe = 1 Or mMe = 2 Then
       aae = c - 4715
    End If
    
    Year = aae: Month = mMe: Day = JJ
    hour = hRe: min = mNe: Sec = CInt(sn)
End Sub

Public Function InvJDYear(ByVal julday As Double) As Double
    Dim Z As Double, a As Double, i As Double, B As Double
    Dim c As Double, D As Double, T As Double, mMe As Double
    
    Z = Int(julday + 0.5)
    
    If Z < 2299161 Then
       a = Z
    Else
       i = Int((Z - 1867216.25) / 36524.25)
       a = Z + 1 + i - Int(i / 4)
    End If
    
    B = a + 1524
    c = Int((B - 122.1) / 365.25)
    D = Int(365.25 * c)
    T = Int((B - D) / 30.6)
    
    If T < 14 Then
       mMe = T - 1
    Else
       If T = 14 Or T = 15 Then mMe = T - 13
    End If
    
    If mMe > 2 Then
       InvJDYear = c - 4716
    ElseIf mMe = 1 Or mMe = 2 Then
       InvJDYear = c - 4715
    End If
    
    'InvJDYear = InvJDYear + (JulDay - JulianDay(InvJDYear, 1, 1, 0, 0)) / 365.25
End Function

Public Function chkJD(ByVal julday As Double, ByVal YY As Double, ByVal MM As Double, ByVal DD As Double, ByVal hr As Double, ByVal mn As Double, Optional ByVal ss As Double) As Boolean
    Dim a As Double, B As Double, c As Double, D As Double, E As Double, f As Double
    
    chkJD = False
    InvJD julday, a, B, c, D, E, f
    If f >= 30 Then E = E + 1
    If YY = a And MM = B And DD = c And hr = D And mn = E Then
      chkJD = True
    End If
End Function

Public Function JDtoYear(ByVal julday As Double) As Double
  Dim SY As Double, eY As Double, CY As Double, fY As Double
  
  CY = Fix(InvJDYear(julday))
  SY = JULIANDAY(CY, 1, 1, 12, 0)
  eY = JULIANDAY(CY + 1, 1, 1, 12, 0)
  fY = (julday - SY) / (eY - SY)
  JDtoYear = CY + fY
End Function

Public Function MakeTimeString(ByVal julday As Double) As String
    Dim pY As Double, pm As Double, pd As Double, PH As Double, pmn As Double, ss As Double
    InvJD julday, pY, pm, pd, PH, pmn, ss
    pmn = Round((pmn + ss / 60) / 60)
    PH = PH + pmn
    MakeTimeString = IIf(PH < 10, "(0", "(") & Trim$(STR$(PH)) & "시)"        ' & "분"
End Function

Public Function MakeTimeString3(ByVal julday As Double) As String
    Dim pY As Double, pm As Double, pd As Double, PH As Double, pmn As Double, ss As Double, tYear As String
    InvJD julday, pY, pm, pd, PH, pmn, ss
    pmn = Round(pmn + ss / 60)
    If pmn = 60 Then PH = PH + 1: pmn = 0
    MakeTimeString3 = IIf(PH < 10, "0", "") & Trim$(STR$(PH)) & ":" & IIf(pmn < 10, "0", "") & Trim$(STR$(pmn))
End Function

Public Function MakeTimeString4(ByVal julday As Double) As String
    Dim pY As Double, pm As Double, pd As Double, PH As Double, pmn As Double, tYear As String, ss As Double
    InvJD julday, pY, pm, pd, PH, pmn, ss
    If pY <= 0 Then tYear = "기원전 " & Trim$(STR$(Abs(pY) + 1)) Else tYear = STR$(pY)
    MakeTimeString4 = tYear & "년" & STR$(pm) & "월" & STR$(pd) & "일"
End Function

Public Function GetJD0(ByVal DateJD As Double) As Double
  If DateJD - Int(DateJD) >= 0.5 Then
    GetJD0 = Int(DateJD) + 0.5
  Else
    GetJD0 = Int(DateJD) - 0.5
  End If
End Function
