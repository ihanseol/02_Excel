Attribute VB_Name = "modEffectiveRadius"
Option Explicit

'0 : skin factor
'1 : Re1
'2 : Re2
'3 : Re3

Public Enum ER_VALUE
    erRE0 = 0
    erRE1 = 1
    erRE2 = 2
    erRE3 = 3
End Enum

Function GetEffectiveRadius() As Integer
    Dim er, r       As String
    
    er = Range("EffectiveRadius").Value
    'MsgBox er
    r = Mid(er, 5, 1)
    
    If r = "F" Then
        GetEffectiveRadius = 0
    Else
        GetEffectiveRadius = Val(r)
    End If
End Function
