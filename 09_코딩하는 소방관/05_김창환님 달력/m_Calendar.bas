Attribute VB_Name = "m_Calendar"
Option Explicit

Dim Lon As Double, Lat As Double, S1 As Double, S2 As Double, FH As Double, td As Boolean
Public TimeZone As Double, RefArea As String, PISLAM As Boolean, YR As Integer, RSP As Integer
Public ShowRSTime As Boolean, ShowAstro As Boolean, ShowSpDay As Boolean
Public UseMeanSun As Boolean, UseMeanMoon As Boolean, UseJinsak As Boolean, AutoConfig As Boolean

Public Sub ReadSet()
  Dim y As Integer, M As Integer
  
  y = YR
  M = Sheet1.Cells(1, 5).Value
  
  Lat = Sgn(Sheet3.Cells(2, 2).Value) * Abs(Sheet3.Cells(2, 2).Value) + Sheet3.Cells(2, 3).Value / 60 + Sheet3.Cells(2, 4).Value / 3600
  Lon = Sgn(Sheet3.Cells(3, 2).Value) * Abs(Sheet3.Cells(3, 2).Value) + Sheet3.Cells(3, 3).Value / 60 + Sheet3.Cells(3, 4).Value / 3600
  TimeZone = Sheet3.Cells(4, 2).Value
  FH = Sheet3.Cells(9, 2).Value
  PISLAM = IIf(Trim(Sheet3.Cells(14, 3).Value) <> "", True, False)
  
  S1 = JULIANDAY(CDbl(y), Sheet3.Cells(7, 2).Value, Sheet3.Cells(7, 3), 12, 0)  '서머타임 시작일
  S2 = JULIANDAY(CDbl(y), Sheet3.Cells(8, 2).Value, Sheet3.Cells(8, 3), 12, 0)  '서머타임 종료일
  
  'RS = Trim(Sheet3.Cells(13, 3).Value) <> ""
  ShowRSTime = Trim(Sheet3.Cells(17, 3).Value) <> ""
  ShowAstro = Trim(Sheet3.Cells(18, 3).Value) <> ""
  ShowSpDay = Trim(Sheet3.Cells(19, 3).Value) <> ""
  
  RefArea = Trim(Sheet3.Cells(11, 3).Value)
  td = Trim(Sheet3.Cells(15, 3).Value) <> ""
  RSP = CInt(Sheet3.Cells(16, 3).Value)
  If RSP < 1 Then RSP = 2: Sheet3.Cells(16, 3).Value = 2
  
  UseMeanSun = Trim(Sheet3.Cells(22, 3).Value) <> ""
  UseMeanMoon = Trim(Sheet3.Cells(23, 3).Value) <> ""
  UseJinsak = Trim(Sheet3.Cells(24, 3).Value) <> ""
  AutoConfig = Trim(Sheet3.Cells(25, 4).Value) <> ""
  
  If AutoConfig = True Then Call AutoChoose(y)
'  If RefArea <> "" And AutoConfig Then Call AutoChooseKR(y, UseMeanSun, UseMeanMoon, UseJinsak)
End Sub

Public Sub GenCal()
  Dim y As Integer, M As Integer
  
  ReadSet
  y = YR
  M = Sheet1.Cells(1, 5).Value
  
  Call MakeCal(y, M)
End Sub


Sub MakeCal(y As Integer, M As Integer)
  Dim JD1 As Double, JD2 As Double, ML As String, ML2 As Double
  Dim Week1 As Integer, DR1 As String, DR2 As String, DR3 As String, DR4 As String, DR5 As String
  Dim PosX As Long, PosY As Long, i As Long, h As Byte, T As String
  Dim B As Boolean, ly As Integer, lm As Byte, LD As Byte, ll As Boolean, TZZ As Double
  Dim TODAY1 As Double, AD As Double, si As String
  
  TZZ = TimeZone
  TODAY1 = JULIANDAY(Year(Now), Month(Now), Day(Now), 12, 0): AD = 0
  
  ML = "312931303130313130313031"
  ML2 = Val(Mid$(ML, 1 + (M - 1) * 2, 2))  '해당월의 길이
  
  JD1 = JULIANDAY(CDbl(y), CDbl(M), 1, 12, 0) '초일의 JD
  JD2 = JD1 + ML2 - 1 '말일의 JD
  If chkJD(JD2, CDbl(y), CDbl(M), ML2, 12, 0) = False Then JD2 = JD2 - 1  '2월 29일이 있는지 판단
  
  Week1 = CInt((JD1 + 1) Mod 7)  '첫날의 요일(0=Sunday)
  If Week1 < 0 Then Week1 = Week1 + 7
  
  For ML2 = 5 To 34   '시트 지우기
    Sheet1.Cells(ML2, 2).Value = ""
    Sheet1.Cells(ML2, 3).Value = ""
    Sheet1.Cells(ML2, 4).Value = ""
    Sheet1.Cells(ML2, 5).Value = ""
    Sheet1.Cells(ML2, 6).Value = ""
    Sheet1.Cells(ML2, 7).Value = ""
    If ML2 < 30 Then Sheet1.Cells(ML2, 8).Value = ""
    Sheet1.Cells(ML2, 2).Interior.Color = vbWhite
    Sheet1.Cells(ML2, 3).Interior.Color = vbWhite
    Sheet1.Cells(ML2, 4).Interior.Color = vbWhite
    Sheet1.Cells(ML2, 5).Interior.Color = vbWhite
    Sheet1.Cells(ML2, 6).Interior.Color = vbWhite
    Sheet1.Cells(ML2, 7).Interior.Color = vbWhite
    If ML2 < 30 Then Sheet1.Cells(ML2, 8).Interior.Color = vbWhite
  Next ML2
  If PISLAM = False Then si = "기념일" Else si = "이슬람력"
  Sheet1.Cells(6, 1).Value = si: Sheet1.Cells(11, 1).Value = si
  Sheet1.Cells(16, 1).Value = si: Sheet1.Cells(21, 1).Value = si
  Sheet1.Cells(26, 1).Value = si: Sheet1.Cells(31, 1).Value = si
  Sheet1.Cells(31, 8).Value = si
  
  If JD1 < 2299161 And JD2 > 2299161 Then AD = 10: JD2 = JD2 - 9
  PosX = 2: PosY = 5 '시트상의 달력 위치
  PosX = PosX + Week1 '첫날의 위치
  FindPPheno JD1 - 0.5, JD2 + 0.5   '행성 현상 찾기
  For ML2 = JD1 To JD2  '날짜별로 필요한 것 계산
    If PosX > 8 Then PosX = 2: PosY = PosY + 5  '한 주가 지나면 다음 줄로
    i = i + 1
    
    DR1 = "": DR2 = "": DR3 = "": DR4 = "": DR5 = ""
    '필요한 계산항==========
    DR1 = STR(ML2 - JD1 + 1 + IIf(ML2 >= 2299161, AD, 0))  '날짜
    If ShowSpDay Then DR2 = Trim$(FindSDay(ML2, h))
    T = JD2M2(ML2)
    If InStr(T, "라") > 0 Then
      If DR2 = "" Then
        DR2 = T
      Else
        DR2 = DR2 & ", " & T
      End If
    End If
    If PISLAM And DR2 = "" Then DR2 = T
    
    Call FindTBL(ML2, ly, lm, LD, ll)  '음력표에서 음력 찾기(고속 계산 가능)
    DR3 = DR3 & MakeDayS(CDbl(ly), CDbl(lm), CDbl(LD), ll) & "/" & Make60(ML2)
    If ShowSpDay Then DR2 = FindSDayL(DR2, lm, LD, ll, h)
    
    If S1 <= ML2 And ML2 <= S2 Then '서머타임 기간이면
      TZZ = TimeZone + FH
    Else
      TZZ = TimeZone
    End If
      
     If ShowAstro Then DR2 = FindSDayA(DR2, ML2, TZZ)
     If ShowRSTime Then DR4 = RSTime(Lon, Lat, ML2, TZZ, HorSun, SUN, RSP)
     If ShowRSTime Then DR5 = RSTime(Lon, Lat, ML2, TZZ, HorMoon, MOON, RSP)
    '여기까지===============
    
    Sheet1.Cells(PosY, PosX).Value = DR1 '내용 적기
    Sheet1.Cells(PosY + 1, PosX).Value = DR2
    Sheet1.Cells(PosY + 2, PosX).Value = DR3
    Sheet1.Cells(PosY + 3, PosX).Value = DR4
    Sheet1.Cells(PosY + 4, PosX).Value = DR5
    If PosX = 8 Then Sheet1.Cells(PosY, PosX).Font.Color = vbBlue
    If PosX = 2 Then Sheet1.Cells(PosY, PosX).Font.Color = vbRed
    If PosX > 2 And PosX < 8 Then Sheet1.Cells(PosY, PosX).Font.Color = vbBlack
    If ML2 = TODAY1 And td Then Sheet1.Cells(PosY, PosX).Interior.Color = vbYellow
    Sheet1.Cells(PosY + 1, PosX).Interior.Color = vbWhite
    If h = 1 Then '휴일이면
      Sheet1.Cells(PosY, PosX).Font.Color = vbRed
    ElseIf h = 2 Then
      Sheet1.Cells(PosY + 1, PosX).Interior.Color = vbCyan
    ElseIf h = 3 Then
      Sheet1.Cells(PosY, PosX).Font.Color = vbRed
      Sheet1.Cells(PosY + 1, PosX).Interior.Color = vbCyan
    End If
    
    PosX = PosX + 1
  Next ML2
End Sub


Function MakeDayS(y As Double, M As Double, D As Double, l As Boolean) As String
  MakeDayS = M & ". " & D & IIf(l, "(윤)", "")
End Function

Function Make60(JD As Double) As String
  Dim i As Double
  
  i = (JD + 49) Mod 60
  If i < 0 Then i = i + 60
  Make60 = Mid$("甲乙丙丁戊己庚辛壬癸", i Mod 10 + 1, 1) & Mid$("子丑寅卯辰巳午未申酉戌亥", i Mod 12 + 1, 1)
End Function

Sub ShowRST(ByVal ShowRS As Boolean)
    Sheet1.Range("A8:H9,A13:H14,A18:H19,A23:H24,A28:H29,A33:H34").Select
    Selection.EntireRow.Hidden = Not ShowRS
    
    If ShowRS = True Then DelLine Else DrawLine
    
    Sheet1.Range("A1").Select
End Sub

Sub DrawLine()
    Sheet1.Range("B32:H32").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub

Sub DelLine()
    Sheet1.Range("B32:H32").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub

