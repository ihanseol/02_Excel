VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub PrevMon()
  Cells(1, 5).Value = Cells(1, 5).Value - 1
  If Cells(1, 5).Value < 1 Then Cells(1, 5).Value = 12: PrevYear: Exit Sub
  GenCal
End Sub

Sub NextMon()
  Cells(1, 5).Value = Cells(1, 5).Value + 1
  If Cells(1, 5).Value > 12 Then Cells(1, 5).Value = 1: Call NextYear: Exit Sub
  GenCal
End Sub

Sub PrevYear()
  Dim bf As Double
  
  YR = YR - 1
  ReadSet
  bf = TimeZone
  If RefArea <> "" Then TimeZone = 9
  Cells(1, 4).Value = YR
  If RefArea <> "" And YR < 1911 Then TimeZone = 8

  If AutoConfig = True Then Call AutoChoose(YR)
  LSTBL2 JULIANDAY(CDbl(YR), Sheet1.Cells(1, 5).Value, 1, 12, 0), TimeZone, UseMeanSun, UseMeanMoon, UseJinsak '해당해의 음력표 만들기
  CalcSpecialDay YR  '해당해의 기념일 계산
  TimeZone = bf

  GenCal
End Sub

Sub NextYear()
  Dim bf As Double
  
  YR = YR + 1
  ReadSet
  bf = TimeZone
  If RefArea <> "" Then TimeZone = 9
  Cells(1, 4).Value = YR
  If RefArea <> "" And YR < 1911 Then TimeZone = 8

  If AutoConfig = True Then Call AutoChoose(YR)
  LSTBL2 JULIANDAY(CDbl(YR), Sheet1.Cells(1, 5).Value, 1, 12, 0), TimeZone, UseMeanSun, UseMeanMoon, UseJinsak
  CalcSpecialDay YR
  TimeZone = bf

  GenCal
End Sub

Private Sub Worksheet_Activate()
  ShowRST ShowRSTime
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
  Dim bf As Double

  If Target.Row = 1 And Target.Column = 4 Then
    If Int(Target.Cells(1, 1).Value) <> YR Then
      YR = Int(Target.Cells(1, 1).Value)
    End If
  End If
End Sub
