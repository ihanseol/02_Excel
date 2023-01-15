Attribute VB_Name = "BaseData_modEffectiveRadius"
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

Function GetER_Mode(ByVal WB_NAME As String) As Integer
    Dim er, r       As String
    
    ' er = Range("h10").value
    er = Workbooks(WB_NAME).Worksheets("SkinFactor").Range("h10").value
    'MsgBox er
    r = Mid(er, 5, 1)
    
    If r = "F" Then
        GetER_Mode = 0
    Else
        GetER_Mode = Val(r)
    End If
End Function

Function GetEffectiveRadius(ByVal WB_NAME As String) As Double
    Dim i, er As Integer
    
    If Not IsWorkBookOpen(WB_NAME) Then
        MsgBox "Please open the yangsoo data ! " & WB_NAME
        Exit Function
    End If
    
    er = GetER_Mode(WB_NAME)
    
    Select Case er
        Case 1
            GetEffectiveRadius = Workbooks(WB_NAME).Worksheets("SkinFactor").Range("k8").value
        Case 2
            GetEffectiveRadius = Workbooks(WB_NAME).Worksheets("SkinFactor").Range("k9").value
        Case 3
            GetEffectiveRadius = Workbooks(WB_NAME).Worksheets("SkinFactor").Range("k10").value
        Case Else
            GetEffectiveRadius = Workbooks(WB_NAME).Worksheets("SkinFactor").Range("C8").value
    End Select

End Function



