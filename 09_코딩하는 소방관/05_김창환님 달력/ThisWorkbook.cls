VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
  ThisMon
End Sub

Sub ThisMon()
  Dim bf As Double
  
  YR = Year(Now())
  Sheet1.Cells(1, 4).Value = YR
  Sheet1.Cells(1, 5).Value = Month(Now())

  ReadSet
  bf = TimeZone
  If RefArea <> "" Then TimeZone = 9
  If RefArea <> "" And YR < 1911 Then TimeZone = 8

  LSTBL2 JULIANDAY(CDbl(YR), Sheet1.Cells(1, 5).Value, 1, 12, 0), TimeZone, UseMeanSun, UseMeanMoon, UseJinsak
  CalcSpecialDay YR
  TimeZone = bf

  GenCal
End Sub
