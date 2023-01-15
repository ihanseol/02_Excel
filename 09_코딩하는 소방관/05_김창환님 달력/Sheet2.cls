VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Deactivate()
  Dim bf As Double

  ReadSet
  bf = TimeZone
  If RefArea <> "" Then TimeZone = 9
  If RefArea <> "" And YR < 1911 Then TimeZone = 8

  If AutoConfig = True Then Call AutoChoose(YR)
  LSTBL2 JULIANDAY(CDbl(YR), Sheet1.Cells(1, 5).Value, 1, 12, 0), TimeZone, UseMeanSun, UseMeanMoon, UseJinsak
  CalcSpecialDay YR
  TimeZone = bf

  Sheet1.Cells(1, 4).Value = YR
  GenCal
End Sub
