VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_AggSum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub

Private Sub CommandButton1_Click()
    Sheets("AggSum").Visible = False
    Sheets("Well").Select
End Sub


Private Sub CommandButton2_Click()
    Dim nofwell As Integer
    
    nofwell = GetNumberOfWell()
    If ActiveSheet.name <> "AggSum" Then Sheets("AggSum").Select


    ' Summary, Aquifer Characterization  Appropriated Water Analysis
    Call Write23_SummaryDevelopmentPotential
    Call Write26_AquiferCharacterization(nofwell)
    Call Write_RadiusOfInfluence(nofwell)
    Call Write_WaterIntake(nofwell)
    Call Write_DiggingDepth(nofwell)
    Call Write_MotorPower(nofwell)
    Call Write_DrasticIndex(nofwell)
    
    Call Write_NaturalLevel(nofwell)
    Call Write_StableLevel(nofwell)
    
End Sub

Sub Write23_SummaryDevelopmentPotential()
' Groundwater Development Potential, 지하수개발가능량
    
    Range("D4").value = Worksheets(CStr(1)).Range("e17").value
    Range("e4").value = Worksheets(CStr(1)).Range("g14").value
    Range("f4").value = Worksheets(CStr(1)).Range("f19").value
    Range("g4").value = Worksheets(CStr(1)).Range("g13").value
    
    Range("h4").value = Worksheets(CStr(1)).Range("g19").value
    Range("i4").value = Worksheets(CStr(1)).Range("f21").value
    Range("j4").value = Worksheets(CStr(1)).Range("e21").value
    Range("k4").value = Worksheets(CStr(1)).Range("g21").value
    
    ' --------------------------------------------------------------------
    
    Range("D8").value = Worksheets(CStr(1)).Range("e17").value
    Range("e8").value = Worksheets(CStr(1)).Range("g14").value
    Range("f8").value = Worksheets(CStr(1)).Range("f19").value
    Range("g8").value = Worksheets(CStr(1)).Range("g13").value
    
    Range("h8").value = Worksheets(CStr(1)).Range("g19").value
    Range("i8").value = Worksheets(CStr(1)).Range("f21").value
    Range("j8").value = Worksheets(CStr(1)).Range("h19").value
    Range("k8").value = Worksheets(CStr(1)).Range("e21").value

End Sub

Sub Write_NaturalLevel(nofwell As Integer)
    Dim i As Integer
    Dim unit As String
    
    Call EraseCellData("d59:p60")
    
    If Sheets("AggSum").CheckBox1.value = True Then
            unit = " m"
        Else
            unit = ""
    End If

    For i = 1 To nofwell
        Cells(59, Chr(Asc("D") + i - 1)).value = "W-" & CStr(i)
        Cells(60, Chr(Asc("D") + i - 1)).value = Worksheets(CStr(i)).Range("c20").value & unit
    Next i
End Sub

Sub Write_StableLevel(nofwell As Integer)
    Dim i As Integer
    Dim unit As String
    
    Call EraseCellData("d64:p65")

    If Sheets("AggSum").CheckBox1.value = True Then
            unit = " m"
        Else
            unit = ""
    End If


    For i = 1 To nofwell
        Cells(64, Chr(Asc("D") + i - 1)).value = "W-" & CStr(i)
        Cells(65, Chr(Asc("D") + i - 1)).value = Worksheets(CStr(i)).Range("c22").value & unit
    Next i
End Sub




Sub Write_DiggingDepth(nofwell As Integer)
    Dim i As Integer
    Dim unit As String
    
    Call EraseCellData("d49:p50")
    
    If Sheets("AggSum").CheckBox1.value = True Then
            unit = " m"
        Else
            unit = ""
    End If

    For i = 1 To nofwell
        Cells(49, Chr(Asc("D") + i - 1)).value = "W-" & CStr(i)
        Cells(50, Chr(Asc("D") + i - 1)).value = Worksheets(CStr(i)).Range("c7").value & unit
    Next i
End Sub


Sub Write_MotorPower(nofwell As Integer)
    Dim i As Integer
    Dim unit As String
    
    Call EraseCellData("d54:p55")
    
    If Sheets("AggSum").CheckBox1.value = True Then
            unit = " Hp"
        Else
            unit = ""
    End If

    For i = 1 To nofwell
        Cells(54, Chr(Asc("D") + i - 1)).value = "W-" & CStr(i)
        Cells(55, Chr(Asc("D") + i - 1)).value = Worksheets(CStr(i)).Range("c17").value & unit
    Next i
End Sub


Sub Write_WaterIntake(nofwell As Integer)
' 취수계획량
    Dim i As Integer
    Dim unit As String
    
    Call EraseCellData("d44:p45")
    
    If Sheets("AggSum").CheckBox1.value = True Then
            unit = Sheets("drastic").Range("a16").value
        Else
            unit = ""
    End If

    For i = 1 To nofwell
        ' WellNum
        Cells(44, Chr(Asc("D") + i - 1)).value = "W-" & CStr(i)
        Cells(45, Chr(Asc("D") + i - 1)).value = Worksheets(CStr(i)).Range("C15").value & unit
    Next i
End Sub


Sub Write_RadiusOfInfluence(nofwell As Integer)
    Dim i As Integer
    Dim unit As String
    
    Call EraseCellData("d28:f40")
    
    If Sheets("AggSum").CheckBox1.value = True Then
            unit = " m"
        Else
            unit = ""
    End If


    For i = 1 To nofwell
        ' WellNum
        Cells(27 + i, "D").value = "W-" & CStr(i)
        ' 양수영향반경, 이것은 보고서에 따라서 다른데,
        ' 일단은 최대값, shultz, webber, jcob 의 최대값을 선택하는것으로 한다.
        ' 그리고 필요한 부분은, 후에 추가시켜준다.
        Cells(27 + i, "E").value = Worksheets(CStr(i)).Range("G11").value & unit
    Next i
End Sub


Sub Write_DrasticIndex(nofwell As Integer)
    Dim i As Integer
    
    Call EraseCellData("h28:j40")
    For i = 1 To nofwell
        ' WellNum
        Cells(27 + i, "H").value = "W-" & CStr(i)
        Cells(27 + i, "I").value = Worksheets(CStr(i)).Range("k30").value
        Cells(27 + i, "J").value = Worksheets(CStr(i)).Range("k31").value
    Next i
End Sub


Sub Write26_AquiferCharacterization(nofwell As Integer)
' 대수층 분석및 적정채수량 분석
    Dim i As Integer
    
    ' Cells(13 + nowfell, "F").Value = Range("well!R5C10").Value
    
    Call EraseCellData("d12:j24")
    
    For i = 1 To nofwell
        ' WellNum --(J==10) / ='1'!$F$21
        Cells(11 + i, "D").value = "W-" & CStr(i)
        ' 심도
        Cells(11 + i, "E").value = Worksheets("Well").Cells(i + 3, 8).value
        ' 양수량
        Cells(11 + i, "F").value = Worksheets("Well").Cells(i + 3, 10).value
        ' 자연수위
        Cells(11 + i, "G").value = Worksheets(CStr(i)).Range("c20").value
        ' 안정수위
        Cells(11 + i, "H").value = Worksheets(CStr(i)).Range("c21").value
        ' 투수량계수
        Cells(11 + i, "I").value = Worksheets(CStr(i)).Range("E7").value
        ' 저류계수
        Cells(11 + i, "J").value = Worksheets(CStr(i)).Range("G7").value
    Next i
End Sub







