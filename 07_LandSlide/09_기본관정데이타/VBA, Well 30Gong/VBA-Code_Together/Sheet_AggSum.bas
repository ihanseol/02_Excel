Option Explicit

' ip_intake : 취수계획량
' ip_simdo : 굴착심도
' ip_pump : 펌프마력
' ip_natural_level  : 자연수위
' ip_stable_level  : 안정수위
' ip_tochul  : 토출구경
' ip_motor_simdo : 모터심도
'
' ip_roi : radius of influence
' ip_di : drastic index
'
' ip_ac : AquiferCharacterization
' ip_right_ac : Right AquiferCharacterization

Const G_WELL As Integer = 30

Private Sub EraseCellData(ByVal str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub

Private Sub CommandButton1_Click()
    Sheets("AggSum").Visible = False
    Sheets("Well").Select
End Sub


Private Sub Test_NameManager()
    Dim acColumn, acRow As Variant

  acColumn = Split(Range("ip_motor_simdo").Address, "$")(1)
  acRow = Split(Range("ip_motor_simdo").Address, "$")(2)

'  Row = ActiveCell.Row
'  col = ActiveCell.Column

    Debug.Print acColumn, acRow
End Sub


Function GetRC(name As String) As Variant
    Dim acColumn, acRow As Variant
    Dim result(1 To 2) As Variant

    acColumn = Split(Range(name).Address, "$")(1)
    acRow = Split(Range(name).Address, "$")(2)

    '  Row = ActiveCell.Row
    '  col = ActiveCell.Column
    
    
    result(1) = acColumn
    result(2) = acRow

    Debug.Print acColumn, acRow
    GetRC = result
End Function



' Summary Button
Private Sub CommandButton2_Click()
    Dim nofwell As Integer
    
    nofwell = GetNumberOfWell()
    If ActiveSheet.name <> "AggSum" Then Sheets("AggSum").Select


    ' Summary, Aquifer Characterization  Appropriated Water Analysis
    Call Write23_SummaryDevelopmentPotential
    Call Write26_AquiferCharacterization(nofwell)
    Call Write26_Right_AquiferCharacterization(nofwell)
    
    Call Write_RadiusOfInfluence(nofwell)
    Call Write_WaterIntake(nofwell)
    Call Check_DI
    
    Call Write_DiggingDepth(nofwell)
    Call Write_MotorPower(nofwell)
    Call Write_DrasticIndex(nofwell)
    
    Call Write_NaturalLevel(nofwell)
    Call Write_StableLevel(nofwell)
    
    
    Call Write_MotorTochool(nofwell)
    Call Write_MotorSimdo(nofwell)

    
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
' 자연수위
    Dim i As Integer
    Dim unit, rngString As String
    Dim Values As Variant
    
    Values = GetRC("ip_natural_level")
    rngString = Values(1) & Values(2) & ":AG" & (Values(2) + 1)
    
    
    Call EraseCellData(rngString)
    
    If Sheets("AggSum").CheckBox1.value = True Then
            unit = " m"
        Else
            unit = ""
    End If

    For i = 1 To nofwell
        Cells(Values(2), (i + 3)).value = "W-" & CStr(i)
        Cells(Values(2) + 1, (i + 3)).value = Worksheets(CStr(i)).Range("c20").value & unit
    Next i
End Sub

Sub Write_StableLevel(nofwell As Integer)
' 안정수위
    Dim i As Integer
    Dim unit, rngString As String
    Dim Values As Variant
    
    Values = GetRC("ip_stable_level")
    rngString = Values(1) & Values(2) & ":AG" & (Values(2) + 1)
    
    
    Call EraseCellData(rngString)
    If Sheets("AggSum").CheckBox1.value = True Then
            unit = " m"
        Else
            unit = ""
    End If


    For i = 1 To nofwell
        Cells(Values(2), (i + 3)).value = "W-" & CStr(i)
        Cells(Values(2) + 1, (i + 3)).value = Worksheets(CStr(i)).Range("c21").value & unit
    Next i
End Sub



' Write_MotorTochool
' Write_MotorSimdo

Sub Write_MotorPower(nofwell As Integer)
' 모터마력
    Dim i As Integer
    Dim unit, rngString As String
    Dim Values As Variant
    
    Values = GetRC("ip_pump")
    rngString = Values(1) & Values(2) & ":AG" & (Values(2) + 1)
    
    
    Call EraseCellData(rngString)
    If Sheets("AggSum").CheckBox1.value = True Then
            unit = " Hp"
        Else
            unit = ""
    End If

    For i = 1 To nofwell
        Cells(Values(2), (i + 3)).value = "W-" & CStr(i)
        Cells(Values(2) + 1, (i + 3)).value = Worksheets(CStr(i)).Range("c17").value & unit
    Next i
End Sub


Sub Write_MotorSimdo(nofwell As Integer)
' 모터심도
    Dim i As Integer
    Dim unit, rngString As String
    Dim Values As Variant
    
    Values = GetRC("ip_motor_simdo")
    rngString = Values(1) & Values(2) & ":AG" & (Values(2) + 1)
    
    
    Call EraseCellData(rngString)
    
    If Sheets("AggSum").CheckBox1.value = True Then
            unit = " m"
        Else
            unit = ""
    End If


    For i = 1 To nofwell
        Cells(Values(2), (i + 3)).value = "W-" & CStr(i)
        Cells(Values(2) + 1, (i + 3)).value = Worksheets(CStr(i)).Range("c18").value & unit
    Next i
End Sub


Sub Write_MotorTochool(nofwell As Integer)
' 토출구경
    Dim i As Integer
    Dim unit, rngString As String
    Dim Values As Variant
    
    Values = GetRC("ip_tochul")
    rngString = Values(1) & Values(2) & ":AG" & (Values(2) + 1)
    
    
    Call EraseCellData(rngString)
    

    If Sheets("AggSum").CheckBox1.value = True Then
            unit = " mm"
        Else
            unit = ""
    End If


    For i = 1 To nofwell
        Cells(Values(2), (i + 3)).value = "W-" & CStr(i)
        Cells(Values(2) + 1, (i + 3)).value = Worksheets(CStr(i)).Range("c19").value & unit
    Next i
End Sub



Sub Write_DiggingDepth(nofwell As Integer)
' 굴착심도
   Dim i As Integer
    Dim unit, rngString As String
    Dim Values As Variant
    
    Values = GetRC("ip_simdo")
    rngString = Values(1) & Values(2) & ":AG" & (Values(2) + 1)
    
    
    Call EraseCellData(rngString)
    
    If Sheets("AggSum").CheckBox1.value = True Then
            unit = " m"
        Else
            unit = ""
    End If

    For i = 1 To nofwell
        Cells(Values(2), (i + 3)).value = "W-" & CStr(i)
        Cells(Values(2) + 1, (i + 3)).value = Worksheets(CStr(i)).Range("c7").value & unit
    Next i
End Sub



Sub Write_WaterIntake(nofwell As Integer)
' 취수계획량
    Dim i As Integer
    Dim unit, rngString As String
    Dim Values As Variant
    
    Values = GetRC("ip_intake")
    rngString = Values(1) & Values(2) & ":AG" & (Values(2) + 1)
    
    
    Call EraseCellData(rngString)
    
    If Sheets("AggSum").CheckBox1.value = True Then
            unit = Sheets("drastic").Range("a16").value
        Else
            unit = ""
    End If

    For i = 1 To nofwell
        ' WellNum
        Cells(Values(2), (i + 3)).value = "W-" & CStr(i)
        Cells(Values(2) + 1, (i + 3)).value = Worksheets(CStr(i)).Range("C15").value & unit
    Next i
End Sub


Sub Write_RadiusOfInfluence(nofwell As Integer)
' 양수영향반경
    Dim i As Integer
    Dim unit, rngString01, rngString02 As String
    Dim Values As Variant
    
    Values = GetRC("ip_roi")
    rngString01 = "D" & Values(2) & ":G" & (Values(2) + G_WELL - 1)
    rngString02 = "M" & Values(2) & ":O" & (Values(2) + G_WELL - 1)
    
    
    ' Call EraseCellData(rngString)
    
    Call EraseCellData(rngString01)
    Call EraseCellData(rngString02)
        
    If Sheets("AggSum").CheckBox1.value = True Then
            unit = " m"
        Else
            unit = ""
    End If


    For i = 1 To nofwell
        ' WellNum
        Cells(Values(2) - 1 + i, "D").value = "W-" & CStr(i)
        ' 양수영향반경, 이것은 보고서에 따라서 다른데,
        ' 일단은 최대값, shultz, webber, jcob 의 최대값을 선택하는것으로 한다.
        ' 그리고 필요한 부분은, 후에 추가시켜준다.
        Cells(Values(2) - 1 + i, "E").value = Worksheets(CStr(i)).Range("H9").value & unit
        Cells(Values(2) - 1 + i, "F").value = Worksheets(CStr(i)).Range("K6").value & unit
        Cells(Values(2) - 1 + i, "G").value = Worksheets(CStr(i)).Range("K7").value & unit
        
        
        '영향반경의 최대, 최소, 평균값을 추가해준다.
        Cells(Values(2) - 1 + i, "M").value = Worksheets(CStr(i)).Range("H9").value & unit
        Cells(Values(2) - 1 + i, "N").value = Worksheets(CStr(i)).Range("H10").value & unit
        Cells(Values(2) - 1 + i, "O").value = Worksheets(CStr(i)).Range("H11").value & unit
        
        
    Next i
End Sub


Sub Write_DrasticIndex(nofwell As Integer)
' 드라스틱 인덱스
    Dim i As Integer
    Dim unit, rngString01, rngString02 As String
    Dim Values As Variant
    
    Values = GetRC("ip_di")
    rngString01 = "I" & Values(2) & ":K" & (Values(2) + G_WELL - 1)
    
    
    Call EraseCellData(rngString01)
    
    For i = 1 To nofwell
        ' WellNum
        Cells(Values(2) - 1 + i, "I").value = "W-" & CStr(i)
        Cells(Values(2) - 1 + i, "J").value = Worksheets(CStr(i)).Range("k30").value
        Cells(Values(2) - 1 + i, "K").value = Worksheets(CStr(i)).Range("k31").value
    Next i
End Sub

Function CheckDrasticIndex(val As Integer) As String
    
    Dim value As Integer
    Dim result As String
    
    Select Case val
        Case Is <= 100
            result = "매우낮음"
        Case Is <= 120
            result = "낮음"
        Case Is <= 140
            result = "비교적낮음"
        Case Is <= 160
            result = "중간정도"
        Case Is <= 180
            result = "높음"
        Case Else
            result = "매우높음"
    End Select
    
    CheckDrasticIndex = result
End Function


Sub Check_DI()
    
    Range("R59").value = CheckDrasticIndex(Range("Q59"))
    Range("R60").value = CheckDrasticIndex(Range("Q60"))

End Sub

Sub Write26_AquiferCharacterization(nofwell As Integer)
' 대수층 분석및 적정채수량 분석
    Dim i, remainder As Integer
    Dim unit, rngString01, rngString02 As String
    Dim Values As Variant
    
    Values = GetRC("ip_ac")
    rngString01 = "D" & Values(2) & ":J" & (Values(2) + G_WELL - 1)
    MsgBox rngString01
    
    Call EraseCellData(rngString01)
    ' Call EraseCellData("D12:J26")
        
    For i = 1 To nofwell
    
        remainder = i Mod 2
        If remainder = 0 Then
                With Range(Cells(11 + i, "d"), Cells(11 + i, "j"))
                    .Font.Bold = True
                End With
        Else
                With Range(Cells(11 + i, "d"), Cells(11 + i, "j"))
                    .Font.Bold = False
                End With
        End If
    
        ' WellNum --(J==10) / ='1'!$F$21
        Cells(11 + i, "D").value = "W-" & CStr(i)
        ' 심도
        Cells(11 + i, "E").value = Worksheets("Well").Cells(i + 3, 8).value
        ' 양수량
        Cells(11 + i, "F").value = Worksheets("Well").Cells(i + 3, 10).value
        
        ' 자연수위
        Cells(11 + i, "G").value = Worksheets(CStr(i)).Range("c20").value
        Cells(11 + i, "G").NumberFormat = "0.00"
        
        ' 안정수위
        Cells(11 + i, "H").value = Worksheets(CStr(i)).Range("c21").value
        Cells(11 + i, "H").NumberFormat = "0.00"
        
        ' 투수량계수
        Cells(11 + i, "I").value = Worksheets(CStr(i)).Range("E7").value
        Cells(11 + i, "I").NumberFormat = "0.0000"
        
        ' 저류계수
        Cells(11 + i, "J").value = Worksheets(CStr(i)).Range("G7").value
        Cells(11 + i, "J").NumberFormat = "0.0000000"
    Next i
End Sub


Sub Write26_Right_AquiferCharacterization(nofwell As Integer)
' 대수층 분석및 적정채수량 분석
    Dim i, remainder As Integer
    Dim unit, rngString01, rngString02 As String
    Dim Values As Variant
    
    Values = GetRC("ip_right_ac")
    rngString01 = "L" & Values(2) & ":S" & (Values(2) + G_WELL - 1)
    
    
    Call EraseCellData(rngString01)
    ' Call EraseCellData("L12:S26")
        
    For i = 1 To nofwell
    
        remainder = i Mod 2
        If remainder = 0 Then
                With Range(Cells(11 + i, "L"), Cells(11 + i, "S"))
                    .Font.Bold = True
                End With
        Else
                With Range(Cells(11 + i, "L"), Cells(11 + i, "S"))
                    .Font.Bold = False
                End With
        End If
    
        ' WellNum --(J==10) / ='1'!$F$21
        Cells(11 + i, "L").value = "W-" & CStr(i)
        ' 심도
        Cells(11 + i, "M").value = Worksheets("Well").Cells(i + 3, 8).value
        ' 양수량
        Cells(11 + i, "N").value = Worksheets("Well").Cells(i + 3, 10).value
        
        ' 자연수위
        Cells(11 + i, "O").value = Worksheets(CStr(i)).Range("c20").value
        Cells(11 + i, "O").NumberFormat = "0.00"
        
        ' 안정수위
        Cells(11 + i, "P").value = Worksheets(CStr(i)).Range("c21").value
        Cells(11 + i, "P").NumberFormat = "0.00"
        
        '수위강하량
        Cells(11 + i, "Q").value = Worksheets(CStr(i)).Range("c21").value - Worksheets(CStr(i)).Range("c20").value
        Cells(11 + i, "Q").NumberFormat = "0.00"
        
        ' 투수량계수
        Cells(11 + i, "R").value = Worksheets(CStr(i)).Range("E7").value
        Cells(11 + i, "R").NumberFormat = "0.0000"
         
        ' 저류계수
        Cells(11 + i, "S").value = Worksheets(CStr(i)).Range("G7").value
        Cells(11 + i, "S").NumberFormat = "0.0000000"
    Next i
End Sub





