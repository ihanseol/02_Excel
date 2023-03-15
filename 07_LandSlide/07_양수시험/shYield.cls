VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shYield"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Call hide_gachae
End Sub

Private Sub Worksheet_Activate()

    If (Range("B14").Value < Range("B15").Value) Then
        Call cellRED
    Else
        Call cellGREEN
    End If
    
    Range("D15").Select

End Sub


Sub cellRED()
    Range("A15:B15").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 192
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub

Sub cellGREEN()

    Range("A15:B15").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent5
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With

End Sub


