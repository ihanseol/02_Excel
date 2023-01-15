VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub DirectChange()
  Dim ly As Integer, lm As Byte, LD As Byte, ll As Boolean, JD As Double
  Dim X As Boolean, a As Double, B As Double, c As Double, bf As Double, bf2 As Integer
  
  ReadSet
  bf = TimeZone: bf2 = YR

  If RefArea <> "" Then TimeZone = 9
  If RefArea <> "" And Cells(3, 8).Value < 1911 Then TimeZone = 8
  JD = JULIANDAY(Cells(3, 8).Value, Cells(3, 10).Value, Cells(3, 12).Value, 12, 0)
  If AutoConfig = True Then Call AutoChoose(Cells(3, 8).Value)
  Call LuniSolarCal(JD, ly, lm, LD, ll, TimeZone, UseMeanSun, UseMeanMoon, UseJinsak)
  Cells(4, 8).Value = ly: Cells(4, 10).Value = lm: Cells(4, 12).Value = LD: Cells(4, 14).Value = IIf(ll, "(À±)", "")
  
  If RefArea <> "" Then TimeZone = 9
  If RefArea <> "" And Cells(7, 8).Value < 1911 Then TimeZone = 8
  If AutoConfig = True Then Call AutoChoose(Cells(7, 8).Value)
  X = InvLuniSolarCal(Cells(7, 8).Value, Cells(7, 10).Value, Cells(7, 12).Value, CheckBox1.Value, TimeZone, UseMeanSun, UseMeanMoon, UseJinsak, JD)
  If X Then
    InvJD JD, a, B, c, 1, 1
    Cells(8, 8).Value = a: Cells(8, 10).Value = B: Cells(8, 12).Value = c
  Else
    Cells(8, 8).Value = "-": Cells(8, 10).Value = "-": Cells(8, 12).Value = "-"
  End If
  TimeZone = bf: YR = bf2
End Sub

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
