Attribute VB_Name = "modInfluenceRadius"
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

Function findInfluenceRadius() As Double
    Dim erMODE, result As String
    Dim WBName, cell1 As String
    
    cell1 = Range("b2").value
    WBName = "A" & GetNumeric2(cell1) & "_ge_OriginalSaveFile.xlsm"
    
    If Not IsWorkBookOpen(WBName) Then
        MsgBox "Please open the yangsoo data ! " & WBName
        Exit Function
    End If
    
    erMODE = Workbooks(WBName).Worksheets("SkinFactor").Range("H10").value
    
    If Mid(erMODE, 5, 1) = "F" Then
        result = 0
    Else
        result = Val(Mid(erMODE, 5, 1))
    End If
    
    Select Case result
        Case erRE1
            findInfluenceRadius = Workbooks(WBName).Worksheets("SkinFactor").Range("k8").value
            
        Case erRE2
            findInfluenceRadius = Workbooks(WBName).Worksheets("SkinFactor").Range("k9").value
            
        Case erRE3
            findInfluenceRadius = Workbooks(WBName).Worksheets("SkinFactor").Range("k10").value
            
        Case Else
            findInfluenceRadius = Workbooks(WBName).Worksheets("SkinFactor").Range("c8").value
    End Select
End Function
