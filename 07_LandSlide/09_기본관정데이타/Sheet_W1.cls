VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_W1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Call find_average
End Sub

Private Sub CommandButton4_Click()
    Call delete_allWhpaData
End Sub



Private Sub CommandButton2_Click()
    Call main_drasticindex
    Call print_drastic_string
End Sub

Private Sub CommandButton3_Click()
    Call getWhpaData_AllWell
End Sub

Private Sub CommandButton7_Click()
   Call getWhpaData_EachWell
End Sub



Private Sub CommandButton5_Click()
    Call ToggleDirection
End Sub


Private Function get_rf_number() As String
    Dim rf_num As String

    '=(max*rf_1*E17/1000)
    get_rf_number = VBA.Mid(Range("F17").formula, 10, 1)

End Function


Private Sub Set_RechargeFactor_One()

    Range("F17").formula = "=(max*rf_1*E17/1000)"
    Range("F19").formula = "=(max*rf_1*E19/1000)/365"
    
    Range("G17").formula = "=F17*allow_ratio"
    Range("G19").formula = "=F19*allow_ratio"
    
    Range("E13").formula = "=Recharge!I24"
    Range("F13").formula = "=rf_1"
    Range("G13").formula = "=allow_ratio"
    
    Range("E26").formula = "=Recharge!C30"
    
End Sub

Private Sub Set_RechargeFactor_Two()

    Range("F17").formula = "=(max*rf_2*E17/1000)"
    Range("F19").formula = "=(max*rf_2*E19/1000)/365"
    
    Range("G17").formula = "=F17*allow_ratio2"
    Range("G19").formula = "=F19*allow_ratio2"
    
    
    Range("E13").formula = "=Recharge!I25"
    Range("F13").formula = "=rf_2"
    Range("G13").formula = "=allow_ratio2"
    
    
    Range("E26").formula = "=Recharge!D30"
End Sub


Private Sub Set_RechargeFactor_Three()

    Range("F17").formula = "=(max*rf_3*E17/1000)"
    Range("F19").formula = "=(max*rf_3*E19/1000)/365"
    
    Range("G17").formula = "=F17*allow_ratio3"
    Range("G19").formula = "=F19*allow_ratio3"
    
    Range("E13").formula = "=Recharge!I26"
    Range("F13").formula = "=rf_3"
    Range("G13").formula = "=allow_ratio3"
    
    Range("E26").formula = "=Recharge!E30"
    
End Sub



Private Sub CommandButton6_Click()
'Select Recharge Factor

    
   If Frame1.Controls("optionbutton1").value = True Then
        Call Set_RechargeFactor_One
   End If
    
   If Frame1.Controls("optionbutton2").value = True Then
        Call Set_RechargeFactor_Two
   End If
    
   If Frame1.Controls("optionbutton3").value = True Then
        Call Set_RechargeFactor_Three
   End If
    

End Sub



' 2022/6/9 Import YangSoo Data
' Radius of Influence - 양수영향반경
' Effective Radius - 유효우물반경

Private Sub CommandButton8_Click()
    Dim WkbkName As Object
    Dim WBName, cell1 As String
    Dim i As Integer
    Dim S1, S2, T1, T2, RI1, RI2, RI3, ir As Double
    
    ' nl : natural level, sl : stable level
    Dim nl, sl, deltas As Double
    Dim casing As Integer
    
    i = 2
    ' Range("i1") = Workbooks.count
    ' WBName = Range("i2").value
    
    cell1 = Range("b2").value
    WBName = "A" & GetNumeric2(cell1) & "_ge_OriginalSaveFile.xlsm"
    
    If Not IsWorkBookOpen(WBName) Then
        MsgBox "Please open the yangsoo data ! " & WBName
        Exit Sub
    End If

    ' delta s : 최초1분의 수위강하
    deltas = Workbooks(WBName).Worksheets("SkinFactor").Range("b4").value
    
    ' 자연수위, 안정수위, 케이싱 심도 결정
    nl = Workbooks(WBName).Worksheets("SkinFactor").Range("i4").value
    sl = Workbooks(WBName).Worksheets("SkinFactor").Range("i6").value
    casing = Workbooks(WBName).Worksheets("SkinFactor").Range("i10").value
    
    ' WkbkName.Close
    T1 = Workbooks(WBName).Worksheets("SkinFactor").Range("D5").value
    S1 = Workbooks(WBName).Worksheets("SkinFactor").Range("E10").value
    T2 = Workbooks(WBName).Worksheets("SkinFactor").Range("H13").value
    S2 = Workbooks(WBName).Worksheets("SkinFactor").Range("i16").value
    
    ' yangsoo radius of influence
    RI1 = Workbooks(WBName).Worksheets("SkinFactor").Range("C13").value
    RI2 = Workbooks(WBName).Worksheets("SkinFactor").Range("C18").value
    RI3 = Workbooks(WBName).Worksheets("SkinFactor").Range("C23").value
    
    ' 유효우물반경 , 설정값에 따른
    ir = GetEffectiveRadius(WBName)
    
    ' 자연수위, 안정수위, 케이싱 심도 결정
    Range("c20") = nl
    Range("c21") = sl
    Range("c10") = 5
    Range("c11") = casing - 5
        
    Range("E5") = T1
    Range("E6") = T2
    Range("g5") = S2
    Range("h7") = S1
    Range("h6") = ir 'find influence radius
    
    Range("e10") = RI1
    Range("f10") = RI2
    Range("g10") = RI3
    
    Range("c23") = Round(deltas, 2) 'deltas
        
End Sub

Private Sub Worksheet_Activate()

    Select Case get_rf_number
    
        Case "1"
             Frame1.Controls("optionbutton1").value = True
             
        Case "2"
             Frame1.Controls("optionbutton2").value = True
             
        Case "3"
             Frame1.Controls("optionbutton3").value = True
             
        Case Else
            Frame1.Controls("optionbutton1").value = True
           
    End Select

End Sub


