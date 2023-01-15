Attribute VB_Name = "m_Moslem"
Option Explicit

'이슬람력 계산

Const AH1 As Double = 1948440  '이슬람력의 시작일(헤지라 날짜)

Public Function MonthName(ByVal mon As Integer) As String
  Dim MonthNm(11) As String  '달의 이름
  
  MonthNm(0) = "Muharram"
  MonthNm(1) = "Safar"
  MonthNm(2) = "Rabi'al-Awwal"
  MonthNm(3) = "Rabi'ath-Thani"
  MonthNm(4) = "Jumada l-Ula"
  MonthNm(5) = "Jumada t-Tania"
  MonthNm(6) = "Rajab"
  MonthNm(7) = "Sha'ban"
  MonthNm(8) = "Ramadan"
  MonthNm(9) = "Shawwal"
  MonthNm(10) = "Dhu l-Qa'da"
  MonthNm(11) = "Dhu l-Hijja"
  
  MonthName = MonthNm(mon - 1)
End Function

'이슬람력->JD
Public Function M2JD(ByVal mY As Integer, MM As Integer, MD As Integer) As Double
  Dim N As Double, Q As Double, R As Double, w As Double, Q1 As Double, Q2 As Double
  Dim g As Double, k As Double, E As Double, j As Double, X As Double, a As Double, JD As Double
  
  '입력 값이 이상하면 계산 안함
  If mY < 0 Or MM < 0 Or MD < 0 Then M2JD = 0: Exit Function
  
  '변환 계산 시작
  N = MD + Int(29.5001 * (MM - 1) + 0.99)
  Q = Int(mY / 30)
  R = mY Mod 30
  a = Int((11 * R + 3) / 30)
  w = 404 * Q + 354 * R + 208 + a
  Q1 = Int(w / 1461)
  Q2 = w Mod 1461
  g = 621 + 4 * Int(7 * Q + Q1)
  k = Int(Q2 / 365.2422)
  E = Int(365.2422 * k)
  j = Q2 - E + N - 1
  X = g + k
  
  If (j > 366) And (X Mod 4 = 0) Then j = j - 366: X = X + 1
  If (j > 365) And (X Mod 4 > 0) Then j = j - 365: X = X + 1
  
  JD = Int(365.25 * (X - 1)) + 1721423 + j
  M2JD = JD
End Function

'JD->이슬람력(문자 형식 출력)
Public Function JD2M2(ByVal JD As Double) As String
  Dim R As String, y As Integer, M As Integer, D As Integer, jd0 As Double

  jd0 = GetJD0(JD) + 0.5
  Call JD2M(jd0, y, M, D)
  If y = 0 Then
    R = ""
  Else
    R = M & "월 " & D & "일"
    If M = 9 And D = 1 Then R = "라마단 시작"
    If M = 9 And D = 30 Then R = "라마단 종료"
  End If
  JD2M2 = R
End Function

'JD->이슬람력
Public Sub JD2M(ByVal JD As Double, mY As Integer, MM As Integer, MD As Integer)
  Dim gY As Double, gM As Double, gD As Double, a As Double, B As Double, JD12 As Double
  Dim alp As Double, Bet As Double, c As Double, D As Double, E As Double
  Dim X1 As Double, M1 As Double, D1 As Double
  Dim N As Double, w As Double, C1 As Double, C2 As Double, Q As Double, R As Double, j As Double, k As Double, O As Double, h As Double
  Dim JJ As Double, cl As Double, dl As Double, S As Double
  
  JD12 = 0.5 + GetJD0(JD)
  '입력받은 날짜가 헤지라 이전이면 계산 안함
  If JD < AH1 Then mY = 0: MM = 0: MD = 0: Exit Sub
  
  Call InvJD(JD12, gY, gM, gD, a, B)
  
  '그레고리력이면 율리우스력으로 바꾸기
  If JD12 >= 2299161 Then
    If gM < 3 Then gY = gY - 1: gM = gM + 12
    alp = Int(gY / 100): Bet = 2 - alp + Int(alp / 4)
    B = Int(365.25 * gY) + Int(30.6001 * (gM + 1)) + gD + 1722519 + Bet
    c = Int((B - 122.1) / 365.25)
    D = Int(365.25 * c)
    E = Int((B - D) / 30.6001)
    D1 = B - D - Int(30.6001 * E)
    If E < 14 Then M1 = E - 1
    If E > 13 Then M1 = E - 13
    If M1 > 2 Then X1 = c - 4716
    If M1 < 3 Then X1 = c - 4715
  Else
    X1 = gY: M1 = gM: D1 = gD
  End If
  
  '이슬람력 계산
  If X1 Mod 4 = 0 Then w = 1 Else w = 2
  N = Int((275 * M1) / 9) - w * Int((M1 + 9) / 12) + D1 - 30
  a = X1 - 623
  B = Int(a / 4)
  c = a Mod 4
  C1 = 365.2501 * c: C2 = Int(C1)
  If C1 - C2 > 0.5 Then C2 = C2 + 1
  D = 1461 * B + 170 + C2
  Q = Int(D / 10631)
  R = D Mod 10631
  j = Int(R / 354)
  k = R Mod 354
  O = Int((11 * j + 14) / 30)
  h = 30 * Q + j + 1
  JJ = k - O + N - 1  'JJ는 이슬람력에서 당년의 날짜 수임
  
  If JJ > 354 Then
    cl = h Mod 30: dl = (11 * cl + 3) Mod 30
    If dl < 19 Then JJ = JJ - 354: h = h + 1
    If dl > 18 Then JJ = JJ - 355: h = h + 1
    If JJ = 0 Then JJ = 355: h = h - 1
  End If
  
  S = Int((JJ - 1) / 29.5)
  
  If JJ = 355 Then
    mY = h: MM = 12: MD = 30
  Else
    mY = h: MM = 1 + S: MD = Int(JJ - 29.5 * S)
  End If
End Sub
