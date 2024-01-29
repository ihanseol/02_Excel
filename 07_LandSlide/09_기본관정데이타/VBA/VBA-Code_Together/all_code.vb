
Private Sub Workbook_Open()
    Call InitialSetColorValue
    Sheets("Well").SingleColor.value = True
    Sheets("Recharge").cbCheSoo.value = True
End Sub

Private Sub Workbook_SheetActivate(ByVal sh As Object)
 ' Call InitialSetColorValue
End Sub


Private Sub CommandButton1_Click()
    Call importRainfall
End Sub

Private Sub CommandButton2_Click()
    Range("b5:n34").ClearContents
End Sub


Private Sub importRainfall()
    Dim myArray As Variant
    Dim rng As Range

    Select Case UCase(Range("T6").value)
        Case "SEJONG", "HONGSUNG"
            Exit Sub
    End Select

    Dim indexString As String
    indexString = "data_" & UCase(Range("T6").value)

    On Error Resume Next
    myArray = Application.Run(indexString)
    On Error GoTo 0

    If Not IsArray(myArray) Then
        MsgBox "An error occurred while fetching data.", vbExclamation
        Exit Sub
    End If

    Set rng = ThisWorkbook.ActiveSheet.Range("B5:N34")
    rng.value = myArray

    Range("B2").value = Range("T5").value & "기상청"
End Sub







' this is empty

Private Sub CommandButton1_Click()
    Call getMotorPower
    Call FindMaxMin
End Sub

Private Sub CommandButton2_Click()
    Rows("36:90").Select
    Selection.Delete Shift:=xlUp
    Range("g34").Select
End Sub

Private Sub FindMaxMin()

    Dim nWell As Integer
    Dim maxVal, minVal As Double
    Dim qMax, qMin As Double
    
    
    nWell = GetNumberOfWell()
    
    If nWell <= 1 Then
        Exit Sub
    End If
    
    maxVal = Application.WorksheetFunction.max(Range("B46:" & Chr(Asc("B") + nWell - 1) & "46"))
    minVal = Application.WorksheetFunction.min(Range("B46:" & Chr(Asc("B") + nWell - 1) & "46"))
    
    qMax = Application.WorksheetFunction.max(Range("B40:" & Chr(Asc("B") + nWell - 1) & "40"))
    qMin = Application.WorksheetFunction.min(Range("B40:" & Chr(Asc("B") + nWell - 1) & "40"))
    
    
    Range("k52") = minVal
    Range("k53") = maxVal
    
    Range("l52") = qMin
    Range("l53") = qMax

End Sub

Private Sub ShowLocation_Click()
      Sheets("location").Visible = True
      Sheets("location").Activate
End Sub



Private Sub CommandButton3_Click()
    Dim i As Integer
    Dim max, min As Single
    
    max = Range("o15").value
    min = Range("o16").value
    
    Range("B5:P14").Select
    Selection.Font.Bold = False
     
    Range("a1").Activate
    
    For i = 5 To 14
        If Cells(i, "O").value = max Or Cells(i, "O").value = min Then
            Union(Cells(i, "B"), Cells(i, "O")).Select
            Selection.Font.Bold = True
        End If
    Next i
End Sub


Option Explicit

Private Sub CommandButton1_Click()
' add well

    Call CopyOneSheet
End Sub

Private Sub CommandButton10_Click()
    Sheets("AggChart").Visible = True
    Sheets("AggChart").Select
End Sub

Private Sub CommandButton11_Click()
    Sheets("YangSoo").Visible = True
    Sheets("YangSoo").Select
End Sub

Private Sub CommandButton12_Click()
    Sheets("water").Visible = True
    Sheets("water").Select
End Sub

Private Sub CommandButton3_Click()
    Sheets("AggSum").Visible = True
    Sheets("AggSum").Select
End Sub

Private Sub CommandButton4_Click()
    Sheets("Aggregate1").Visible = True
    Sheets("Aggregate1").Select
End Sub

Private Sub CommandButton5_Click()
    Sheets("Aggregate2").Visible = True
    Sheets("Aggregate2").Select
End Sub


Private Sub CommandButton7_Click()
    Sheets("aggWhpa").Visible = True
    Sheets("aggWhpa").Select
End Sub


Private Sub CommandButton9_Click()
  Sheets("AggStep").Visible = True
  Sheets("AggStep").Select
End Sub


'Jojung Button
'add new feature - correct border frame ...
Private Sub CommandButton2_Click()
    Dim nofwell As Integer

    nofwell = sheets_count()
    Call JojungSheetData
    Call make_wellstyle
    Call DecorateWellBorder(nofwell)
    
    Worksheets("1").Range("E21") = "=Well!" & Cells(5 + GetNumberOfWell(), "I").Address
End Sub


' delete last
Private Sub CommandButton8_Click()
    Dim nofwell As Integer
    'nofwell = GetNumberOfWell()
    nofwell = sheets_count()
    
    If nofwell = 1 Then
        MsgBox "Last is not delete ... ", vbOK
        Exit Sub
    End If
    
    Rows(nofwell + 3).Delete
    Call DeleteWorksheet(nofwell)
    Call DecorateWellBorder(nofwell - 1)
End Sub

Sub DeleteWorksheet(well As Integer)
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(CStr(well)).Delete
    Application.DisplayAlerts = True
End Sub


Private Sub Worksheet_Activate()
    Call InitialSetColorValue
End Sub


Private Sub DecorateWellBorder(ByVal nofwell As Integer)
    Sheets("Well").Activate
    Range("A2:R" & CStr(nofwell + 3)).Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    
    Range("D15").Select
End Sub


Private Sub getDuoSolo(ByVal nofwell As Integer, ByRef nDuo As Integer, ByRef nSolo As Integer)
    Dim page, quotient, remainder As Integer
    
    quotient = WorksheetFunction.quotient(nofwell, 2)
    remainder = nofwell Mod 2
    
    If remainder = 0 Then
        nDuo = quotient
        nSolo = 0
    Else
        nDuo = quotient
        nSolo = 1
    End If

End Sub

'one button
'delete all well except for one ...

Private Sub CommandButton6_Click()
    Dim i, nofwell As Integer
    Dim response As VbMsgBoxResult
    
    nofwell = GetNumberOfWell()
    
    If nofwell = 1 Then Exit Sub

    response = MsgBox("Do you deletel all water well?", vbYesNo)
    If response = vbYes Then
         For i = 2 To nofwell
             RemoveSheetIfExists (CStr(i))
        Next i
        
        Sheets("Well").Activate
        Rows("5:" & CStr(nofwell + 3)).Select
        Selection.Delete Shift:=xlUp
        
        For i = 1 To 12
            If Not RemoveSheetIfExists("p" & CStr(i)) Then Exit For
        Next i
        
        Call DecorateWellBorder(1)
        Range("A1").Select
    End If
   
End Sub

Function RemoveSheetIfExists(shname As String) As Boolean
    Dim ws As Worksheet
    Dim sheetExists As Boolean
    
    sheetExists = False

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(shname)
    If Not ws Is Nothing Then sheetExists = True
    On Error GoTo 0

    If sheetExists Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
        RemoveSheetIfExists = True
        Exit Function
    Else
        RemoveSheetIfExists = False
        Exit Function
    End If
End Function



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
    Dim S1, S2, S3, T1, T2, RI1, RI2, RI3, ir As Double
    
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
    S3 = Workbooks(WBName).Worksheets("SkinFactor").Range("i13").value
    
    ' yangsoo radius of influence
    RI1 = Workbooks(WBName).Worksheets("SkinFactor").Range("C13").value
    RI2 = Workbooks(WBName).Worksheets("SkinFactor").Range("C18").value
    RI3 = Workbooks(WBName).Worksheets("SkinFactor").Range("C23").value
    
    ' 유효우물반경 , 설정값에 따른
    ir = GetEffectiveRadius(WBName)
    
    ' 자연수위, 안정수위, 케이싱 심도 결정
    Range("c20") = nl
    Range("c20").NumberFormat = "0.00"
    
    Range("c21") = sl
    Range("c21").NumberFormat = "0.00"
    
    Range("c10") = 5
    Range("c11") = casing - 5
    
    'in recover test, s' value
    Range("G6") = S3
        
    Range("E5") = T1
    Range("E5").NumberFormat = "0.0000"
     
    Range("E6") = T2
    Range("E6").NumberFormat = "0.0000"
    
    Range("g5") = S2
    Range("g5").NumberFormat = "0.0000000"
    
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


'기본관정데이타 - 드라스틱인덱스
Option Explicit

Dim Dr, Rr  As Single

Public Enum DRASTIC_MODE
    dmGENERAL = 0
    dmCHEMICAL = 1
End Enum

Sub ShiftNewYear()
    Range("B6:N34").Select
    ActiveWindow.SmallScroll Down:=-9
    Selection.Copy
    
    Range("B5").Select
    ActiveSheet.PasteSpecial Format:=3, Link:=1, DisplayAsIcon:=False, _
                             IconFileName:=False
    
    Range("B34:N34").Select
    Selection.ClearContents
    
    ActiveWindow.SmallScroll Down:=18
    Range("B42:N50").Select
    Selection.Copy
    Range("B41").Select
    ActiveSheet.PasteSpecial Format:=3, Link:=1, DisplayAsIcon:=False, _
                             IconFileName:=False
    Range("B50:N50").Select
    Selection.ClearContents
End Sub

Sub ToggleDirection()
    If Range("k12").Font.Bold Then
        Range("K12").Font.Bold = False
        Range("L12").Font.Bold = True
        
        CellBlack (ActiveSheet.Range("L12"))
        CellLight (ActiveSheet.Range("K12"))
    Else
        Range("K12").Font.Bold = True
        Range("L12").Font.Bold = False
        
        CellBlack (ActiveSheet.Range("K12"))
        CellLight (ActiveSheet.Range("L12"))
    End If
End Sub

Private Sub CellBlack(S As Range)
    S.Select
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent1
        .TintAndShade = -0.499984740745262
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
End Sub

Private Sub CellLight(S As Range)
    S.Select
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
End Sub

'Drastic Index 를 계산 해주기 위한 함수 ...
' 2017/11/21 화요일

' 1, 지하수위에 대한 등급의 계산
Private Function Rating_UnderGroundWater(ByVal water_level As Single) As Integer
    Select Case water_level
        Case Is < 1.52
            Rating_UnderGroundWater = 10
        Case Is < 4.57
            Rating_UnderGroundWater = 9
        Case Is < 9.14
            Rating_UnderGroundWater = 7
        Case Is < 15.24
            Rating_UnderGroundWater = 5
        Case Is < 22.86
            Rating_UnderGroundWater = 3
        Case Is < 30.48
            Rating_UnderGroundWater = 2
        Case Else
            Rating_UnderGroundWater = 1
    End Select
End Function


'2, 강수의 지하함양량
Private Function Rating_NetRecharge(ByVal value As Single) As Integer
    Select Case value
        Case Is < 5.08
            Rating_NetRecharge = 1
        Case Is < 10.16
            Rating_NetRecharge = 3
        Case Is < 17.78
            Rating_NetRecharge = 6
        Case Is < 25.4
            Rating_NetRecharge = 8
        Case Else
            Rating_NetRecharge = 9
    End Select
End Function



'3, 대수층
Private Function Rating_AqMedia(ByVal value As String) As Integer
    Dim ratings As New Dictionary
    
    ratings.Add "Massive Shale", 2
    ratings.Add "Metamorphic/Igneous", 3
    ratings.Add "Weathered Metamorphic / Igneous", 4
    ratings.Add "Glacial Till", 5
    ratings.Add "Bedded SandStone", 6
    ratings.Add "Massive Sandstone", 6
    ratings.Add "Massive Limestone", 6
    ratings.Add "Sand And Gravel", 8
    ratings.Add "Basalt", 9
    ratings.Add "Karst Limestone", 10

    If ratings.Exists(value) Then
        Rating_AqMedia = ratings(value)
    Else
        Rating_AqMedia = 0
    End If
End Function


'4 토양특성에 대한 등급

Private Function Rating_SoilMedia(ByVal value As String) As Integer
    Select Case value
        Case "Thin Or Absent", "Gravel"
            Rating_SoilMedia = 10
        Case "Sand"
            Rating_SoilMedia = 9
        Case "Peat"
            Rating_SoilMedia = 8
        Case "Shrinking Or Aggregated Clay"
            Rating_SoilMedia = 7
        Case "Sandy Loam"
            Rating_SoilMedia = 6
        Case "Loam"
            Rating_SoilMedia = 5
        Case "Silty Loam"
            Rating_SoilMedia = 4
        Case "Clay Loam"
            Rating_SoilMedia = 3
        Case "Mud"
            Rating_SoilMedia = 2
        Case "Nonshrinking And Nonaggregated Clay"
            Rating_SoilMedia = 1
    End Select
End Function


' 5, 지형구배
Private Function Rating_Topo(ByVal value As Single) As Integer
    Select Case value
        Case Is < 2
            Rating_Topo = 10
        Case Is < 6
            Rating_Topo = 9
        Case Is < 12
            Rating_Topo = 5
        Case Is < 18
            Rating_Topo = 3
        Case Else
            Rating_Topo = 1
    End Select
End Function



'6 비포화대의 영향에 대한 등급 Ir
'
Private Function Rating_Vadose(ByVal value As String) As Integer
    Select Case value
        Case "Confining Layer"
            Rating_Vadose = 1
        Case "Silt/Clay", "Shale"
            Rating_Vadose = 3
        Case "Limestone", "Sandstone", "Bedded Limestone, Sandstone, Shale", "Sand And Gravel With Significant Silt And Clay"
            Rating_Vadose = 6
        Case "Metamorphic/Igneous"
            Rating_Vadose = 4
        Case "Sand And Gravel"
            Rating_Vadose = 8
        Case "Basalt"
            Rating_Vadose = 9
        Case "Karst Limestone"
            Rating_Vadose = 10
    End Select
End Function


' 7, 대수층의 수리전도도에 대한 등급 : Cr

Private Function Rating_EC(ByVal value As Double) As Integer
    Select Case value
        Case Is < 0.0000472
            Rating_EC = 1
        Case Is < 0.000142
            Rating_EC = 2
        Case Is < 0.00033
            Rating_EC = 4
        Case Is < 0.000472
            Rating_EC = 6
        Case Is < 0.000944
            Rating_EC = 8
        Case Else
            Rating_EC = 10
    End Select
End Function



Public Sub find_average()
    ' 2019/10/18 일 작성함
    ' get 투수량계수, 대수층, 유향, 동수경사의 평균을 구해야한다.
    
    Dim n_sheets    As Integer
    Dim i           As Integer
    
    Dim nTooSoo     As Single: nTooSoo = 0
    Dim nDaeSoo     As Single: nDaeSoo = 0
    Dim nDirection  As Single: nDirection = 0
    Dim nGradient   As Single: nGradient = 0
    
    n_sheets = sheets_count()
    
    For i = 1 To n_sheets
        
        Worksheets(CStr(i)).Activate
        
        nTooSoo = nTooSoo + Range("E7").value
        nDaeSoo = nDaeSoo + Range("C14").value
        nDirection = nDirection + get_direction()
        nGradient = nGradient + Range("K18").value
        
    Next i
    
    Worksheets("1").Activate
    
    Range("J3").value = nTooSoo / n_sheets
    Range("J4").value = nDaeSoo / n_sheets
    Range("J5").value = nDirection / n_sheets
    Range("J6").value = nGradient / n_sheets
    
    Range("k3").formula = "=round(j3,4)"
    Range("k4").formula = "=round(j4,1)"
    Range("k5").formula = "=round(j5,1)"
    Range("k6").formula = "=round(j6,4)"
    
    Call make_frame
End Sub

Public Sub find_average2(ByVal sheet As Integer, ByVal nof_well As Integer)
    ' 2019/10/18 일 작성함
    ' get 투수량계수, 대수층, 유향, 동수경사의 평균을 구해야한다.
    
    Dim n_sheets    As Integer
    Dim i           As Integer
    
    Dim nTooSoo     As Single: nTooSoo = 0
    Dim nDaeSoo     As Single: nDaeSoo = 0
    Dim nDirection  As Single: nDirection = 0
    Dim nGradient   As Single: nGradient = 0
    
    Worksheets(CStr(sheet)).Activate
    
    For i = 1 To nof_well
        Worksheets(CStr(i + sheet - 1)).Activate
        
        nTooSoo = nTooSoo + Range("E7").value
        nDaeSoo = nDaeSoo + Range("C14").value
        nDirection = nDirection + get_direction()
        nGradient = nGradient + Range("K18").value
    Next i
    
    Worksheets(CStr(sheet)).Activate
    
    Range("J3").value = nTooSoo / nof_well
    Range("J4").value = nDaeSoo / nof_well
    Range("J5").value = nDirection / nof_well
    Range("J6").value = nGradient / nof_well
    
    Range("k3").formula = "=round(j3,4)"
    Range("k4").formula = "=round(j4,1)"
    Range("k5").formula = "=round(j5,1)"
    Range("k6").formula = "=round(j6,4)"
    
    Call make_frame2(sheet)
End Sub

Private Function get_direction() As Long
    ' get direction is cell is bold
    ' 셀이 볼드값이면 선택을 한다.  방향이 두개중에서 하나를 선택하게 된다.
    ' 2019/10/18일
    
    Range("k12").Select
    
    If Selection.Font.Bold Then
        get_direction = Range("k12").value
    Else
        get_direction = Range("L12").value
    End If
End Function

Sub main_drasticindex()
    Dim water_level, net_recharge, topo, EC As Single
    Dim AQ, Soil, Vadose As String
    Dim drastic_string As String
    
    Dim n_sheets    As Integer
    Dim i           As Integer
    
    ' 쉬트의 갯수 ..., 검사할 공의 갯수
    
    n_sheets = sheets_count()
    
    For i = 1 To n_sheets
        Worksheets(CStr(i)).Activate
        
        '1
        water_level = Range("D26").value
        Range("D27").value = Rating_UnderGroundWater(water_level)
        
        '2
        net_recharge = Range("E26").value
        Range("E27").value = Rating_NetRecharge(net_recharge)
        
        '3
        AQ = Range("F26").value
        Range("F27").value = Rating_AqMedia(AQ)
        
        '4
        Soil = Range("G26").value
        Range("G27").value = Rating_SoilMedia(Soil)
        
        '5
        topo = Range("H26").value
        Range("H27").value = Rating_Topo(topo)
        
        '6 Iv, Vadose
        Vadose = Range("I26").value
        Range("I27").value = Rating_Vadose(Vadose)
        
        '7
        EC = Range("J26").value
        Range("J27").value = Rating_EC(EC)
        
    Next i
End Sub


Function check_DrasticIndex(ByVal dmMode As Integer) As String
    ' dmGENERAL = 0
    ' dmCHEMICAL = 1
    
    Dim value As Integer
    Dim result As String
    
    If dmMode = dmGENERAL Then
        value = Range("K30").value
    Else
        value = Range("K31").value
    End If
    
    Select Case value
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
    
    check_DrasticIndex = result
End Function



Public Sub print_drastic_string()
    Dim n_sheets    As Integer
    Dim i           As Integer
    
    n_sheets = sheets_count()
    
    For i = 1 To n_sheets
        Worksheets(CStr(i)).Activate
        Range("k26").value = check_DrasticIndex(dmGENERAL)
        Range("k27").value = check_DrasticIndex(dmCHEMICAL)
    Next i
End Sub

' this is empty


Private Sub CommandButton1_Click()
UserFormTS.Show
End Sub


'Get Water Spec from YanSoo ilbo
Private Sub CommandButton2_Click()
  Dim thisname, fname As String
  Dim cell  As String
  Dim time As Date
  Dim bTemp, ec1, ph1 As Double
  
  
  cell = Range("d12").value
  
  thisname = ActiveWorkbook.Name
  fname = "A" & GetNumeric2(cell) & "_ge_OriginalSaveFile.xlsm"
 
  If Not IsWorkBookOpen(fname) Then
    MsgBox "Please open the yangsoo data ! " & fname
    Exit Sub
  End If
  
  ' Range("k2") = fname
   
  '------------------------------------------------------------------------
  time = Workbooks(fname).Worksheets("w1").Range("c6").value
  bTemp = Workbooks(fname).Worksheets("w1").Range("c7").value
  
  ec1 = Workbooks(fname).Worksheets("w1").Range("c8").value
  ph1 = Workbooks(fname).Worksheets("w1").Range("c9").value
  
  '------------------------------------------------------------------------
  
  Range("c6").value = time
  Range("c7").value = bTemp
  Range("c8").value = ec1
  Range("c9").value = ph1
    
  Call getModDataFromYangSooSingle(thisname, fname)
End Sub


Sub getModDataFromYangSooSingle(ByVal thisname As String, ByVal fname As String)
    Windows(fname).Activate
    Sheets("w1").Activate
    Sheets("w1").Range("H14:J23").Select
    Selection.Copy
    
    Windows(thisname).Activate
    Range("h14").Select
   
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
End Sub

Function IsSheet(shname As String) As Boolean
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(shname)
    
    Application.DisplayAlerts = False
    IsSheet = True
    Application.DisplayAlerts = True
    
    Exit Function
    
ErrorHandler:
    ' MsgBox "An error occurred while trying to delete the worksheet."
    
    IsSheet = False
    Application.DisplayAlerts = True
End Function


' Get(Ec, Ph, Temp) Range - 지열공에서 통계내는 함수 ....
' Ph, EC, Temp statistics, find range
' data gathering function in EarthThermal test ...
Private Sub CommandButton3_Click()
    Dim nofwell, i As Integer
    
    Dim lowEC() As Double
    Dim hiEC() As Double
    Dim lowPH() As Double
    Dim hiPH() As Double
    Dim lowTEMP() As Double
    Dim hiTEMP() As Double

    nofwell = sheets_count()
    
'    If nofwell < 2 Or Not Contains(Sheets, "a1") Then
'        MsgBox "first Generate Simple YangSoo"
'        Exit Sub
'    End If
    
    If Not IsSheet("p1") Then
        MsgBox "First Make Summary Page"
        Exit Sub
    End If
    
 
    ReDim lowPH(1 To nofwell)
    ReDim hiPH(1 To nofwell)
    
    ReDim lowEC(1 To nofwell)
    ReDim hiEC(1 To nofwell)
    
    ReDim lowTEMP(1 To nofwell)
    ReDim hiTEMP(1 To nofwell)
    
    For i = 1 To nofwell
        lowEC(i) = getEC(cellLOW, i)
        hiEC(i) = getEC(cellHI, i)
        
        lowPH(i) = getPH(cellLOW, i)
        hiPH(i) = getPH(cellHI, i)
        
        lowTEMP(i) = getTEMP(cellLOW, i)
        hiTEMP(i) = getTEMP(cellHI, i)
    Next i
    
    Debug.Print String(3, vbCrLf)
    
    Debug.Print "--Temp----------------------------------------"
    Debug.Print "low : " & Application.min(lowTEMP), Application.max(lowTEMP)
    Debug.Print "hi  : " & Application.min(hiTEMP), Application.max(hiTEMP)
    Debug.Print "----------------------------------------------"
    
    Debug.Print "--PH------------------------------------------"
    Debug.Print "low : " & Application.min(lowPH), Application.max(lowPH)
    Debug.Print "hi  : " & Application.min(hiPH), Application.max(hiPH)
    Debug.Print "----------------------------------------------"
       
    Debug.Print "--EC------------------------------------------"
    Debug.Print "low : " & Application.min(lowEC), Application.max(lowEC)
    Debug.Print "hi  : " & Application.min(hiEC), Application.max(hiEC)
    Debug.Print "----------------------------------------------"
End Sub

Function getEC(ByVal LOWHI As Integer, ByVal well As Integer)
    Sheets("p" & well).Activate
    
    If LOWHI = cellLOW Then
        getEC = Sheets("p" & CStr(well)).Range("e25").value
    Else
        getEC = Sheets("p" & CStr(well)).Range("e24").value
    End If
End Function

Function getPH(ByVal LOWHI As Integer, ByVal well As Integer)
    Sheets("p" & CStr(well)).Activate
    
    If LOWHI = cellLOW Then
        getPH = Sheets("p" & CStr(well)).Range("f25").value
    Else
        getPH = Sheets("p" & CStr(well)).Range("f24").value
    End If
    
End Function

Function getTEMP(ByVal LOWHI As Integer, ByVal well As Integer)
    Sheets("p" & CStr(well)).Activate

    If LOWHI = cellLOW Then
        getTEMP = Sheets("p" & CStr(well)).Range("d25").value
    Else
        getTEMP = Sheets("p" & CStr(well)).Range("d24").value
    End If
End Function


Sub DuplicateQ1Page(ByVal n As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q1")
    
    
    ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    ActiveSheet.Name = "p" & n
    
    With ActiveSheet.Tab
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0
    End With
    
    Call SetWellPropertyQ1(n)
    
End Sub

Sub SetWellPropertyQ1(ByVal i As Integer)
    ActiveSheet.Range("C4") = "W-" & CStr(i)
    ActiveSheet.Range("D12") = "W-" & CStr(i)
    ActiveSheet.Range("H12") = "W-" & CStr(i)
    
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton5")).Select
    Selection.Delete
End Sub


Private Sub DeleteWorksheet(shname As String)
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(shname)
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    
    Exit Sub
    
ErrorHandler:
    ' MsgBox "An error occurred while trying to delete the worksheet."
    Application.DisplayAlerts = True
End Sub


' make summary page
Private Sub CommandButton4_Click()
    Dim nofwell As Integer
    Dim i As Integer
    
    If IsSheet("p1") Then
        MsgBox "Sheet P1 Exist .... Delete First ... ", vbOKOnly
        Exit Sub
    End If
       
    
    nofwell = GetNumberOfWell()
    
    For i = 1 To nofwell
        DuplicateQ1Page (i)
    Next i
End Sub


' delete all summary page
Private Sub CommandButton5_Click()

    Dim nofwell As Integer
    Dim i As Integer
    
    nofwell = GetNumberOfWell()
    
    For i = 1 To nofwell
        DeleteWorksheet ("p" & i)
    Next i
    
    Sheets("Q1").Activate

End Sub























Public Sub MalgunGothic()
    ActiveWindow.SmallScroll Down:=78
    Cells.Select
    Range("A200").Activate
    With Selection.Font
        .Name = "맑은 고딕"
    End With
    
    Range("C186").Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
    Selection.End(xlUp).Select
End Sub

Public Sub make_frame2(ByVal sh As Integer)
    Worksheets(CStr(sh)).Activate
    
    Range("i3").value = "투수량계수"
    Range("i4").value = "대수층두께"
    Range("i5").value = "유향"
    Range("i6").value = "동수경사"
    
    Range("I3:K6").Select
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = False
    End With
    
    With Selection.Font
        .Name = "맑은 고딕"
        .Size = 11
    End With
    
    Range("I3:K6").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Range("k3").Select
End Sub

Public Sub make_frame()
    Range("i3").value = "투수량계수"
    Range("i4").value = "대수층두께"
    Range("i5").value = "유향"
    Range("i6").value = "동수경사"
    
    Range("I3:K6").Select
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = False
    End With
    
    With Selection.Font
        .Name = "맑은 고딕"
        .Size = 11
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
    End With
    
    Range("I3:K6").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub
Public Sub draw_motor_frame(ByVal nof_sheets As Integer, ByVal po As Integer)
    Dim mychar
    
    Debug.Print lastRow()
    mychar = Chr(65 + nof_sheets)
    
    Range("A" & (po) & ":" & mychar & (po + 12)).Select
    
    Call draw_border
    Range("A" & (po) & ":" & mychar & (po + 1)).Select
    Call draw_border
    Range("A" & (po + 11) & ":" & mychar & (po + 12)).Select
    Call draw_border
    Range("A" & (po) & ":" & "A" & (po + 12)).Select
    Call draw_border
    
    Range("A" & (po + 2) & ":" & "A" & (po + 10)).Select
    Call draw_border
    
    Range("A" & (po) & ":B" & (po)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    
    Range("a" & (po)).value = "펌프마력산정하는것"
    Range("a" & (po + 2)).value = "굴착심도"
    Range("a" & (po + 3)).value = "Q(물량)-양수량"
    Range("a" & (po + 4)).value = "Depth(모터설치심도)"
    Range("a" & (po + 5)).value = "Height(양정고)"
    Range("a" & (po + 6)).value = "Sum (합계)"
    Range("a" & (po + 7)).value = "E (효율)"
    Range("a" & (po + 9)).value = "계산식"
    Range("a" & (po + 11)).value = "허가필증의 마력"
    Range("a" & (po + 12)).value = "이론상 양수능력"
    
    Call decorationPumpHP(nof_sheets, po)
    Call decorationInerLine(nof_sheets, po)
    Call alignTitle(nof_sheets, po)
End Sub

Private Sub alignTitle(ByVal nof_sheets As Integer, ByVal po As Integer)
    Dim mychar
    mychar = Chr(65 + nof_sheets)
    
    'Range("A57:B57").Select
    Range("A" & (po) & ":" & "B" & (po)).Select
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    Selection.Font.Italic = False
    With Selection.Font
        .Name = "맑은 고딕"
        .Size = 12
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .ThemeFont = xlThemeFontMinor
    End With
    
    'Range("A59:A69").Select
    Range("A" & (po + 2) & ":" & "A" & (po + 12)).Select
    
    With Selection.Font
        .Name = "맑은 고딕"
        .Size = 11
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = False
    End With
    
    Range("O65").Select
End Sub

Private Sub decorationPumpHP(ByVal nof_sheets As Integer, ByVal po As Integer)
    Dim mychar
    
    mychar = Chr(65 + nof_sheets)
    'Range("B58:N69").Select
    Range("B" & (po + 1) & ":" & mychar & (po + 12)).Select
    
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    With Selection.Font
        .Name = "맑은 고딕"
        .Size = 12
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    'Range("B59:N59").Select
    Range("B" & (po + 2) & ":" & mychar & (po + 2)).Select
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    'Range("B60:N60").Select
    Range("B" & (po + 3) & ":" & mychar & (po + 3)).Select
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    
    'Range("B63:N63").Select
    Range("B" & (po + 6) & ":" & mychar & (po + 6)).Select
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    ActiveWindow.SmallScroll Down:=3
    
    'Range("B64:N64").Select
    Range("B" & (po + 7) & ":" & mychar & (po + 7)).Select
    Selection.NumberFormatLocal = "0.00"
    
    'Range("B67:N67").Select
    Range("B" & (po + 10) & ":" & mychar & (po + 10)).Select
    Selection.Font.Bold = True
    With Selection.Font
        .Name = "맑은 고딕"
        .Size = 14
        .Italic = True
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    'Range("B68:N69").Select
    Range("B" & (po + 11) & ":" & mychar & (po + 12)).Select
    
    Selection.Font.Bold = True
    Selection.Font.Italic = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
End Sub

Private Sub decorationInerLine(ByVal nof_sheets As Integer, ByVal po As Integer)
    Dim mychar
    
    mychar = Chr(65 + nof_sheets)
    'Range("A60:N61").Select
    Range("A" & (po + 3) & ":" & mychar & (po + 4)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    'Range("B67:N67").Select
    Range("B" & (po + 10) & ":" & mychar & (po + 10)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    'Range("B59:N67").Select
    Range("B" & (po + 2) & ":" & mychar & (po + 10)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    
    'Range("A59:A67").Select
    Range("A" & (po + 2) & ":" & "A" & (po + 10)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    
    'Range("B68:N69").Select
    Range("B" & (po + 11) & ":" & mychar & (po + 12)).Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    
    'Range("A68:A69").Select
    Range("A" & (po + 11) & ":" & "A" & (po + 12)).Select
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
End Sub

Private Sub draw_border()
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
End Sub
Option Explicit

Public ip, ip2  As Long

Private Function nColorsInArray(ByRef array_tabcolor() As Variant, ByVal CHECK As Variant) As Integer
    ' array_tabcolor :
    ' check : color값
    ' 관정에 지정하는 색갈은 모두 달라야 한다.
    ' 컬러값이 관정에 몇개가 있는지를 리턴
    
    Dim i, limit    As Integer
    Dim count       As Integer: count = 0
    
    limit = UBound(array_tabcolor, 1)
    
    For i = 1 To limit
        If array_tabcolor(i) = CHECK Then
            count = count + 1
        End If
    Next i
    
    nColorsInArray = count
End Function

Private Function getans_tabcolors() As Variant
    Dim n_sheets, i, j, limit As Integer
    Dim arr_tabcolors(), new_tabcolors(), ans_tabcolors() As Variant
    
    'uc : unique colors
    Dim uc          As Integer
    
    n_sheets = sheets_count()
    
    ReDim arr_tabcolors(1 To n_sheets)
    ReDim new_tabcolors(1 To n_sheets)
    ReDim ans_tabcolors(0 To n_sheets)
    
    For i = 1 To n_sheets
        arr_tabcolors(i) = Sheets(CStr(i)).Tab.Color
    Next i
    
    new_tabcolors = getUnique(arr_tabcolors)
    limit = GetLength(new_tabcolors)
    ans_tabcolors(0) = 1
    
    For i = 1 To limit
        uc = nColorsInArray(arr_tabcolors, new_tabcolors(i - 1))
        ans_tabcolors(i) = ans_tabcolors(i - 1) + uc
    Next i
    
    getans_tabcolors = ans_tabcolors
End Function

Private Function getkey_tabcolors() As Object
    ' return value using by dictionary
    ' 1:(3), 2:(2), 3:(2), 4:(1)
    ' c.Add Item:=CStr(uc), key:=CStr(i)
    
    Dim n_sheets, i, j, limit As Integer
    Dim arr_tabcolors(), new_tabcolors() As Variant
    
    'uc : unique colors
    Dim uc          As Integer
    
    'c colors code
    Dim C           As Collection
    Set C = New Collection
    
    n_sheets = sheets_count()
    
    ReDim arr_tabcolors(1 To n_sheets)
    ReDim new_tabcolors(1 To n_sheets)
    
    For i = 1 To n_sheets
        arr_tabcolors(i) = Sheets(CStr(i)).Tab.Color
    Next i
    
    new_tabcolors = getUnique(arr_tabcolors)
    limit = GetLength(new_tabcolors)
    
    For i = 1 To limit
        uc = nColorsInArray(arr_tabcolors, new_tabcolors(i - 1))
        C.Add item:=CStr(uc), key:=CStr(i)
    Next i
    
    Set getkey_tabcolors = C
End Function

Private Sub get_tabsize(ByRef nof_sheets As Integer, ByRef nof_unique_tab As Integer)
    Dim n_sheets, i, j, limit As Integer
    Dim arr_tabcolors(), new_tabcolors() As Variant
    
    n_sheets = sheets_count()
    
    ReDim arr_tabcolors(1 To n_sheets)
    ReDim new_tabcolors(1 To n_sheets)
    
    For i = 1 To n_sheets
        arr_tabcolors(i) = Sheets(CStr(i)).Tab.Color
    Next i
    
    new_tabcolors = getUnique(arr_tabcolors)
    limit = GetLength(new_tabcolors)
    
    nof_sheets = n_sheets
    nof_unique_tab = limit
End Sub

Sub getWhpaData_AllWell()
    Dim nof_sheets  As Integer
    Dim nof_unique_tab As Integer
    Dim i, sheet    As Integer
    
    Call get_tabsize(nof_sheets, nof_unique_tab)
    
    Application.ScreenUpdating = False
    
    For i = 1 To nof_sheets
        
        Call find_average2(i, 1)
        
    Next i
    
    Application.ScreenUpdating = True
End Sub

Sub getWhpaData_EachWell()
    ' 2019년 당진아파트 7지구에서 처럼, 2019년 6번 폴더
    ' rc : return collection
    
    Dim r_ans()     As Variant
    Dim rc          As Collection
    Dim nof_sheets  As Integer
    Dim nof_unique_tab As Integer
    Dim i, sheet    As Integer
    
    Set rc = getkey_tabcolors()
    r_ans = getans_tabcolors()
    Call get_tabsize(nof_sheets, nof_unique_tab)
    
    Debug.Print rc(1)
    Debug.Print r_ans(0)
    Debug.Print nof_sheets
    Debug.Print nof_unique_tab
    
    ' Call find_average2(1, rc(1))
    
    Application.ScreenUpdating = False
    
    For i = 1 To nof_unique_tab
        
        sheet = r_ans(i - 1)
        Call find_average2(sheet, rc(i))
        
    Next i
    
    Application.ScreenUpdating = True
End Sub

Sub delete_allWhpaData()
    Dim n_sheets    As Long
    Dim i           As Long
    
    n_sheets = sheets_count()
    
    For i = 1 To n_sheets
        Worksheets(CStr(i)).Activate
        
        Range("I3:K6").Select
        Selection.Clear
        Range("H9").Select
    Next i
End Sub


'@2023-7-15 , Refactor

Private Function get_efficiency_A(ByVal Q As Variant) As Variant
    Dim thresholds As Variant
    Dim results As Variant
    
    Dim i, result As Integer
    Dim CHECK As Boolean
    
    CHECK = True
    
    thresholds = Array(57.6, 72, 86.4, 115.2, 144, 216, 288, 432, 576, 720, 864, 1152, 1440)
    results = Array(40, 42, 45, 48, 50, 52, 54, 57, 59, 61, 62, 64, 65)
    
    result = results(0)

    For i = 1 To UBound(thresholds)
        If Q >= thresholds(i - 1) And Q < thresholds(i) Then
            result = results(i - 1)
            CHECK = False
            Exit For
        End If
    Next i
        
    get_efficiency_A = result
    
    If Q < 57.6 Then get_efficiency_A = 40
    If Q > 1440 Then get_efficiency_A = 65
    
End Function



'@2023-7-15 , Refactor

Private Function get_efficiency_B(ByVal Q As Variant) As Variant
    Dim thresholds As Variant
    Dim results As Variant
    
    Dim i, result As Integer
    Dim CHECK As Boolean
    
    CHECK = True
    
    thresholds = Array(57.6, 72, 86.4, 115.2, 144, 216, 288, 432, 576, 720, 864, 1152, 1440)
    results = Array(34, 36, 38, 41, 42, 44, 46, 48, 50, 52, 53, 54, 55)
    
    result = results(0)

    For i = 1 To UBound(thresholds)
        If Q >= thresholds(i - 1) And Q < thresholds(i) Then
            result = results(i - 1)
            CHECK = False
            Exit For
        End If
    Next i
        
    get_efficiency_B = result
    
    If Q < 57.6 Then get_efficiency_B = 34
    If Q > 1440 Then get_efficiency_B = 55
    
End Function



'@2023-7-15 , Refactor

Private Function get_efficiency_dongho(ByVal Q As Variant) As Variant
    Dim results As Variant
    Dim thresholds As Variant
    Dim i, result As Integer
    
    
    thresholds = Array(72, 86.4, 115.2, 144, 216, 288, 432, 576, 720, 864, 1152, 1440)
    results = Array(38, 40.25, 43, 45.25, 47, 49, 51.25, 53.5, 55.5, 57, 58.25, 59.5)
        
    result = results(0)

    For i = 1 To UBound(thresholds)
        If Q >= thresholds(i - 1) And Q < thresholds(i) Then
            result = results(i)
            Exit For
        End If
    Next i
    
    
    get_efficiency_dongho = result
    If Q < 57.6 Then get_efficiency_dongho = 37
    If Q > 1440 Then get_efficiency_dongho = 60
    
End Function



Public Sub getMotorPower()
    Dim r_ans()     As Variant
    Dim rc          As Collection        'return collection
    Dim nof_sheets  As Integer
    Dim nof_unique_tab As Integer
    Dim i, sheet    As Integer
    
    Dim title()     As Variant
    Dim simdo()     As Variant
    Dim pump_q()    As Variant
    Dim motor_depth() As Variant
    Dim efficiency() As Variant
    Dim hp()        As Variant
    
    Set rc = getkey_tabcolors()
    r_ans = getans_tabcolors()
    Call get_tabsize(nof_sheets, nof_unique_tab)
    
    ReDim title(1 To nof_sheets)
    ReDim simdo(1 To nof_sheets)
    ReDim pump_q(1 To nof_sheets)
    ReDim motor_depth(1 To nof_sheets)
    ReDim efficiency(1 To nof_sheets)
    ReDim hp(1 To nof_sheets)
    
    ip = lastRow() + 4
    ip2 = ip + 15
    
    Application.ScreenUpdating = False
    
    For i = 1 To nof_sheets
        Worksheets(CStr(i)).Activate
        
        title(i) = Range("b2").value
        simdo(i) = Range("c7").value
        
        ' 채수계획량을 선택할것인지, 양수량을 선택할것인지
        If Sheets("Recharge").cbCheSoo.value = True Then
            pump_q(i) = Range("c15").value
        Else
            pump_q(i) = Range("c16").value
        End If
        
        motor_depth(i) = Range("c18").value
        
        '2022/8/4 select efficiency
        If Sheets("Recharge").OptionButton1.value Then
          efficiency(i) = get_efficiency_A(pump_q(i))
        ElseIf Sheets("Recharge").OptionButton2.value Then
          efficiency(i) = get_efficiency_B(pump_q(i))
        Else
           efficiency(i) = get_efficiency_dongho(pump_q(i))
        End If
        
        hp(i) = Range("c17").value
    Next i
    
    Sheet_Recharge.Activate
    
    Call draw_motor_frame(nof_sheets, ip)
    
    For i = 1 To nof_sheets
        Call insert_basic_entry(title(i), simdo(i), pump_q(i), motor_depth(i), efficiency(i), hp(i), i, ip)
        Call insert_cell_function(i, ip)
    Next i
    
    
    ' -----------------------------------
    ' 2023-07-15
    ' -----------------------------------
    
    For i = 1 To nof_sheets
        Call insert_downform(pump_q(i), motor_depth(i), efficiency(i), title(i), ip2 + i - 1)
    Next i
    
    Call DecoLine(i, ip2)
    
    Application.ScreenUpdating = True
End Sub


Public Sub insert_downform(pump_q As Variant, motor_simdo As Variant, e As Variant, title As Variant, ByVal po As Integer)
    Dim tenper As Double
    Dim sum_simdo As Double
    
    
    tenper = Round(motor_simdo / 10, 1)
    sum_simdo = motor_simdo + tenper
    
    Cells(po, "A").value = title
    Cells(po, "B").value = pump_q
    Cells(po, "C").value = motor_simdo
    Cells(po, "D").value = tenper
    Cells(po, "E").value = sum_simdo
    Cells(po, "F").value = e
    Cells(po, "G").value = "-"
    Cells(po, "H").value = Round((pump_q * (motor_simdo + tenper)) / (6572.5 * (e / 100)), 4)
    Cells(po, "I").value = find_P2(Cells(po, "H").value)
    
    
    Debug.Print "{ " & pump_q & " TIMES " & sum_simdo & " } over { " & "6,572.5" & " TIMES " & (e / 100) & " }"

End Sub


Function find_P2(ByVal num As Double) As Double
    Dim thresholds As Variant
    Dim i As Integer
    thresholds = Array(1, 2, 3, 5, 7.5, 10, 15, 20, 25, 30)
    
    find_P2 = thresholds(0)
    
    For i = 1 To UBound(thresholds)
        If num >= thresholds(i - 1) And num < thresholds(i) Then
            find_P2 = thresholds(i)
            Exit For
        End If
    Next i
End Function



Sub DecoLine(ByVal i As Integer, ByVal po As Integer)
    Rows(po & ":" & (po + i - 2)).Select
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    With Selection.Font
        .Name = "Arial"
        .Size = 12
        .Italic = True
    End With
    
        
    Range("A" & po & ":I" & (po + i - 2)).Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
    End With
        
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlDot
        .Weight = xlThin
    End With
    
    Range("C62").Select
End Sub

    
Function RoundUpNumber(ByVal num As Double)
    Dim roundedNum As Double
    roundedNum = Application.WorksheetFunction.RoundUp(num, 0)
    RoundUpNumber = roundedNum
End Function


Private Sub insert_cell_function(ByVal n As Integer, ByVal position As Integer)
    'height1 : 양정고
    'height : 높이합계
    
    Dim mychar
    Dim height, height1, eq, round_hp, theory_hp As String
    Dim h1, h2      As Integer
    
    h1 = position + 4
    h2 = position
    
    mychar = Chr(65 + n)
    ' Debug.Print mychar
    
    height = "=" & mychar & CStr(h1) & "+" & mychar & CStr(h1 + 1)
    height1 = "=round(" & mychar & CStr(h2 + 4) & "/10,1)"
    
    eq = "=round((" & mychar & CStr(h2 + 3) & "*" & mychar & CStr(h2 + 6) & ")/(6572.5*" & mychar & CStr(h2 + 7) & "),4)"
    round_hp = "=roundup(" & mychar & CStr(h2 + 9) & ",0)"
    theory_hp = "=round((" & mychar & CStr(h2 + 11) & "*" & mychar & CStr(h2 + 7) & "*6572.5)" & "/" & mychar & CStr(h2 + 6) & ",1)"
    
    Range(mychar & CStr(h2 + 5)).formula = height1        '양정고
    Range(mychar & CStr(h2 + 6)).formula = height        '합계
    
    Range(mychar & CStr(h2 + 9)).formula = eq
    Range(mychar & CStr(h2 + 10)).formula = round_hp
    Range(mychar & CStr(h2 + 12)).formula = theory_hp

End Sub


'ip : insertion point

Private Sub insert_basic_entry(title As Variant, simdo As Variant, Q As Variant, motor_depth As Variant, _
                                e As Variant, hp As Variant, ByVal i As Integer, ByVal po As Variant)
    Dim mychar As String
    
    mychar = Chr(65 + i)
    Range(mychar & CStr(po + 1)).value = title
    Range(mychar & CStr(po + 2)).value = simdo
    Range(mychar & CStr(po + 3)).value = Q
    Range(mychar & CStr(po + 4)).value = motor_depth
    Range(mychar & CStr(po + 7)).value = e / 100
    Range(mychar & CStr(po + 11)).value = hp
    
    Call SetFontMalgun(mychar, ip)
End Sub

Sub SetFontMalgun(ByVal col As String, ByVal ip As Integer)
    With Range("col" & CStr(ip) & ":" & col & CStr((ip + 11))).Font
        .Name = "Arial"
        .Size = 12
        .Bold = True
        .Italic = True
    End With
End Sub



Option Explicit

Public Sub rows_and_column()
    Debug.Print Cells(20, 1).Address(RowAbsolute:=False, ColumnAbsolute:=False)
    Debug.Print Range("a20").Row & " , " & Range("a20").column
    
    Range("B2:Z44").Rows(3).Select
End Sub

Public Sub ShowNumberOfRowsInSheet1Selection()
    Dim area        As Range
    
    ' Worksheets("Sheet1").Activate
    Dim selectedRange As Excel.Range
    Set selectedRange = Selection
    
    Dim areaCount   As Long
    areaCount = Selection.Areas.count
    
    If areaCount <= 1 Then
        MsgBox "The selection contains " & _
               Selection.Rows.count & " rows."
    Else
        Dim areaIndex As Long
        areaIndex = 1
        For Each area In Selection.Areas
            MsgBox "Area " & areaIndex & " of the selection contains " & _
                   area.Rows.count & " rows." & " Selection 2 " & Selection.Areas(2).Rows.count & " rows."
            areaIndex = areaIndex + 1
        Next
    End If
End Sub


' Refactor 2023/10/20
Function myRandBetween(i As Integer, j As Integer, Optional div As Integer = 100) As Single
    Dim SIGN        As Integer
    
    SIGN = IIf(WorksheetFunction.RandBetween(0, 1), 1, -1)
    
    myRandBetween = (WorksheetFunction.RandBetween(i, j) / div) * SIGN
End Function

Function myRandBetween2(i As Integer, j As Integer, Optional div As Integer = 100) As Single
    Dim SIGN        As Integer
    
    myRandBetween = (WorksheetFunction.RandBetween(i, j) / div)
End Function


' Refactor 2023/10/20
Public Sub rnd_between()
    Dim i As Integer
    
    For i = 14 To 24
        Dim SIGN As Integer
        SIGN = IIf(WorksheetFunction.RandBetween(0, 1), 1, -1)
        
        Cells(i, 14).value = (WorksheetFunction.RandBetween(7, 12) / 100) * SIGN
        
        With Cells(i, 14)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .NumberFormatLocal = "0.00"
        End With
    Next i
End Sub


Private Sub CommandButton1_Click()
UserFormTS.Show
End Sub


Function DivideWellsBy3(ByVal numberOfWells As Integer) As Integer()

    Dim quotient As Integer
    Dim remainder As Integer
    Dim result(1) As Integer
    
    quotient = numberOfWells \ 3
    remainder = numberOfWells Mod 3
    
    result(0) = quotient
    result(1) = remainder
    
    DivideWellsBy3 = result
    
End Function

Sub SetWellPropertyQ3(ByVal i As Integer)
    
    ActiveSheet.Range("D12") = "W-" & CStr((i - 1) * 3 + 1)
    ActiveSheet.Range("D29") = "W-" & CStr((i - 1) * 3 + 1)
    
    ActiveSheet.Range("G12") = "W-" & CStr((i - 1) * 3 + 2)
    ActiveSheet.Range("H29") = "W-" & CStr((i - 1) * 3 + 2)
    
    ActiveSheet.Range("J12") = "W-" & CStr((i - 1) * 3 + 3)
    ActiveSheet.Range("L29") = "W-" & CStr((i - 1) * 3 + 3)
    
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
    
End Sub

Sub SetWellPropertyRest(ByVal wselect As Integer, ByVal w3page As Integer)
    Dim firstwell As Integer
      
    firstwell = 3 * w3page + 1
    
    If wselect = 2 Then
        ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
        Selection.Delete
        ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
        Selection.Delete
        
        
        ActiveSheet.Range("D12") = "W-" & CStr(firstwell)
        ActiveSheet.Range("D29") = "W-" & CStr(firstwell)
        
        ActiveSheet.Range("G12") = "W-" & CStr(firstwell + 1)
        ActiveSheet.Range("H29") = "W-" & CStr(firstwell + 1)
    Else
    
        ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
        Selection.Delete
        ActiveSheet.Shapes.Range(Array("CommandButton5")).Select
        Selection.Delete
    
        ActiveSheet.Range("D12") = "W-" & CStr(firstwell)
        ActiveSheet.Range("H12") = "W-" & CStr(firstwell)
    End If
End Sub

Sub DuplicateQ3Page(ByVal n As Integer)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q3")
    
    For i = 1 To n
        ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        ActiveSheet.Name = "p" & i
        
        With ActiveSheet.Tab
            .ThemeColor = xlThemeColorAccent3
            .TintAndShade = 0
        End With
        
        Call SetWellPropertyQ3(i)
    Next i
        
End Sub

Sub DuplicateRest(ByVal wselect As Integer, ByVal w3page As Integer)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q" & CStr(wselect))
    
    ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    ActiveSheet.Name = "p" & CStr(w3page + 1)
    
    With ActiveSheet.Tab
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0
    End With
        
    Call SetWellPropertyRest(wselect, w3page)
    
End Sub

'
' 2023/3/15, make summary page
' i think, this procedure made before .. but ican't the source excel file ..
' so i make again
' get page from pagenum
' quotient, remainder
' pagenum - 7 : 7/3 - 2, 1 (if remainder 1 = w1, remain = 2, w2)
'
Private Sub CommandButton3_Click()
' make summary page

    Dim n_sheets As Integer
    Dim result() As Integer
    Dim w3page, wselect, restpage As Integer
    'wselect = 1 --> only w1
    'wselect = 2 --> w1, w2
    
    
    If IsSheet("p1") Then
        MsgBox "Sheet P1 Exist .... Delete First ... ", vbOKOnly
        Exit Sub
    End If
       
    n_sheets = sheets_count()
    result = DivideWellsBy3(n_sheets)
    
    
    ' result(0) = quotient
    ' result(1) = remainder
    w3page = result(0)
    
    Select Case result(1)
        Case 0
            restpage = 0
            wselect = 0
            
        Case 1
            restpage = 1
            wselect = 1
            
        Case 2
            restpage = 1
            wselect = 2
    End Select
    
    
    Call DuplicateQ3Page(w3page)
    
    If restpage = 0 Then
        Exit Sub
    Else
        Call DuplicateRest(wselect, w3page)
    End If
    
End Sub


Private Sub CommandButton2_Click()
  Dim thisname, fname1, fname2, fname3 As String
  Dim cell1, cell2, cell3 As String
  Dim time1 As Date
  Dim bTemp, ec1, ec2, ec3, ph1, ph2, ph3 As Double
  
  cell1 = Range("d12").value
  cell2 = Range("g12").value
  cell3 = Range("j12").value
  
  
  thisname = ActiveWorkbook.Name
  fname1 = "A" & GetNumeric2(cell1) & "_ge_OriginalSaveFile.xlsm"
  fname2 = "A" & GetNumeric2(cell2) & "_ge_OriginalSaveFile.xlsm"
  fname3 = "A" & GetNumeric2(cell3) & "_ge_OriginalSaveFile.xlsm"
   
  If Not IsWorkBookOpen(fname1) Then
    MsgBox "Please open the yangsoo data ! " & fname1
    Exit Sub
  End If
  
  If Not IsWorkBookOpen(fname2) Then
    MsgBox "Please open the yangsoo data ! " & fname2
    Exit Sub
  End If
  
  If Not IsWorkBookOpen(fname3) Then
    MsgBox "Please open the yangsoo data ! " & fname3
    Exit Sub
  End If
  
  'Range("k2") = fname1
  'Range("k3") = fname2
  'Range("k4") = fname3
  
  '------------------------------------------------------------------------
  time1 = Workbooks(fname1).Worksheets("w1").Range("c6").value
  bTemp = Workbooks(fname1).Worksheets("w1").Range("c7").value
  
  ec1 = Workbooks(fname1).Worksheets("w1").Range("c8").value
  ph1 = Workbooks(fname1).Worksheets("w1").Range("c9").value
  
  ec2 = Workbooks(fname2).Worksheets("w1").Range("c8").value
  ph2 = Workbooks(fname2).Worksheets("w1").Range("c9").value
  
  ec3 = Workbooks(fname3).Worksheets("w1").Range("c8").value
  ph3 = Workbooks(fname3).Worksheets("w1").Range("c9").value
  '------------------------------------------------------------------------
  
  
  Range("c6").value = time1
  Range("c7").value = bTemp
  Range("c8").value = ec1
  Range("c9").value = ph1
  
  Range("d8").value = ec2
  Range("d9").value = ph2
  
  Range("e8").value = ec3
  Range("e9").value = ph3
  
  
  Call getModDataFromYangSooTripple(thisname, fname1)
  Call getModDataFromYangSooTripple(thisname, fname2)
  Call getModDataFromYangSooTripple(thisname, fname3)
End Sub


Sub getModDataFromYangSooTripple(ByVal thisname As String, ByVal fname As String)

    Dim f As Integer

    f = CInt(GetNumeric2(fname)) Mod 3

    Windows(fname).Activate
    Sheets("w1").Activate
    Sheets("w1").Range("H14:J23").Select
    Selection.Copy
    
    Windows(thisname).Activate
    
    If f = 0 Then
        Range("l31").Select
    ElseIf f = 1 Then
        Range("d31").Select
    Else
        Range("h31").Select
    End If
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
End Sub


Private Sub DeleteWorksheet(shname As String)
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(shname)
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    
    Exit Sub
    
ErrorHandler:
    ' MsgBox "An error occurred while trying to delete the worksheet."
    Application.DisplayAlerts = True
End Sub

Private Function IsSheet(shname As String) As Boolean
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(shname)
    
    Application.DisplayAlerts = False
    IsSheet = True
    Application.DisplayAlerts = True
    
    Exit Function
    
ErrorHandler:
    ' MsgBox "An error occurred while trying to delete the worksheet."
    
    IsSheet = False
    Application.DisplayAlerts = True
End Function


Private Sub CommandButton4_Click()
' delete all summary page

    Dim i, nofwell, pn As Integer
    Dim response As VbMsgBoxResult
        
    nofwell = GetNumberOfWell()
    pn = (nofwell \ 3) + (nofwell Mod 3)
    
    Sheets("Q3").Activate
    
    response = MsgBox("Do you deletel all water well?", vbYesNo)
    If response = vbYes Then
        For i = 1 To pn
            Call DeleteWorksheet("p" & i)
        Next i
    End If
    
    Sheets("Q3").Activate
    
End Sub

Private Sub CommandButton5_Click()
' get ec, ph, temp

    Call DataAnalysis
End Sub


' Get(Ec, Ph, Temp) Range - 지열공에서 통계내는 함수 ....
' Ph, EC, Temp statistics, find range
' data gathering function in EarthThermal test ...
Private Sub DataAnalysis()
    Dim nofwell, i As Integer
    
    Dim lowEC() As Double
    Dim hiEC() As Double
    Dim lowPH() As Double
    Dim hiPH() As Double
    Dim lowTEMP() As Double
    Dim hiTEMP() As Double

    nofwell = sheets_count()
    
'    If nofwell < 2 Or Not Contains(Sheets, "a1") Then
'        MsgBox "first Generate Simple YangSoo"
'        Exit Sub
'    End If
    
    If Not IsSheet("p1") Then
        MsgBox "First Make Summary Page"
        Exit Sub
    End If
    

    ReDim lowPH(1 To nofwell)
    ReDim hiPH(1 To nofwell)
    
    ReDim lowEC(1 To nofwell)
    ReDim hiEC(1 To nofwell)
    
    ReDim lowTEMP(1 To nofwell)
    ReDim hiTEMP(1 To nofwell)
    
    For i = 1 To nofwell
        lowEC(i) = getEC(cellLOW, i)
        hiEC(i) = getEC(cellHI, i)
        
        lowPH(i) = getPH(cellLOW, i)
        hiPH(i) = getPH(cellHI, i)
        
        lowTEMP(i) = getTEMP(cellLOW, i)
        hiTEMP(i) = getTEMP(cellHI, i)
    Next i
    
    Debug.Print String(3, vbCrLf)
    
    Debug.Print "--Temp----------------------------------------"
    Debug.Print "low : " & Application.min(lowTEMP), Application.max(lowTEMP)
    Debug.Print "hi  : " & Application.min(hiTEMP), Application.max(hiTEMP)
    Debug.Print "----------------------------------------------"
    
    Debug.Print "--PH------------------------------------------"
    Debug.Print "low : " & Application.min(lowPH), Application.max(lowPH)
    Debug.Print "hi  : " & Application.min(hiPH), Application.max(hiPH)
    Debug.Print "----------------------------------------------"
       
    Debug.Print "--EC------------------------------------------"
    Debug.Print "low : " & Application.min(lowEC), Application.max(lowEC)
    Debug.Print "hi  : " & Application.min(hiEC), Application.max(hiEC)
    Debug.Print "----------------------------------------------"
End Sub



' 1, 2, 3 --> p1
' 4, 5, 6 --> p2

Function getEC(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well - 1, 3)
    remainder = well Mod 3
    page = quo + 1
       
    Select Case remainder
        Case 0
            If LOWHI = cellLOW Then
                getEC = Sheets("p" & CStr(page)).Range("k25").value
            Else
                getEC = Sheets("p" & CStr(page)).Range("k24").value
            End If
            
        Case 1
            If LOWHI = cellLOW Then
                getEC = Sheets("p" & CStr(page)).Range("e25").value
            Else
                getEC = Sheets("p" & CStr(page)).Range("e24").value
            End If
        
        Case 2
            If LOWHI = cellLOW Then
                getEC = Sheets("p" & CStr(page)).Range("h25").value
            Else
                getEC = Sheets("p" & CStr(page)).Range("h24").value
            End If
    End Select
End Function

Function getPH(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well - 1, 3)
    remainder = well Mod 3
    page = quo + 1
       
    Select Case remainder
        Case 0
            If LOWHI = cellLOW Then
                getPH = Sheets("p" & CStr(page)).Range("l25").value
            Else
                getPH = Sheets("p" & CStr(page)).Range("l24").value
            End If
            
        Case 1
            If LOWHI = cellLOW Then
                getPH = Sheets("p" & CStr(page)).Range("f25").value
            Else
                getPH = Sheets("p" & CStr(page)).Range("f24").value
            End If
        
        Case 2
            If LOWHI = cellLOW Then
                getPH = Sheets("p" & CStr(page)).Range("i25").value
            Else
                getPH = Sheets("p" & CStr(page)).Range("i24").value
            End If
    End Select
End Function

Function getTEMP(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well - 1, 3)
    remainder = well Mod 3
    page = quo + 1
       
    Select Case remainder
        Case 0
            If LOWHI = cellLOW Then
                getTEMP = Sheets("p" & CStr(page)).Range("J25").value
            Else
                getTEMP = Sheets("p" & CStr(page)).Range("J24").value
            End If
            
        Case 1
            If LOWHI = cellLOW Then
                getTEMP = Sheets("p" & CStr(page)).Range("d25").value
            Else
                getTEMP = Sheets("p" & CStr(page)).Range("d24").value
            End If
        
        Case 2
            If LOWHI = cellLOW Then
                getTEMP = Sheets("p" & CStr(page)).Range("g25").value
            Else
                getTEMP = Sheets("p" & CStr(page)).Range("g24").value
            End If
    End Select
End Function



Option Explicit

'쉬트를 생성할때에는 전체 관정데이타를 건들지 않고, 우선먼저 쉬트복제를 누르는것이 기본으로 정해져 있다.
'Private Sub deleteCommandButton()
'
'     ActiveSheet.Shapes.Range(Array("CommandButton2")).Select
'     Selection.Delete
'
'End Sub

Private Sub DeleteCommandButton()
    ActiveSheet.Shapes.Range(Array("CommandButton2")).Delete
End Sub


Public Sub CopyOneSheet()
    Dim n_sheets    As Integer
    
    n_sheets = sheets_count()
    
    '2020/5/30 관정리스트의 목록삽입해주는 부분 추가
    InsertOneRow (n_sheets)
    
    If (n_sheets = 1) Then
        Sheets("1").Select
        Sheets("1").Copy Before:=Sheets("Q1")
        Call DeleteCommandButton
    Else
        Sheets("2").Select
        Sheets("2").Copy Before:=Sheets("Q1")
    End If
    
    ActiveSheet.Name = CStr(n_sheets + 1)
    Range("b2").value = "W-" & (n_sheets + 1)
    Range("e15").value = CStr(n_sheets + 1)
    
    '2022/6/9 일
    Range("i2") = "A" & CStr(n_sheets + 1) & "_ge_OriginalSaveFile.xlsm"
    
    If n_sheets = 1 Then
        Call ChangeCellData(n_sheets + 1, 1)
    Else
        Call ChangeCellData(n_sheets + 1, 2)
    End If
    
    Sheets("Well").Select
End Sub

Private Sub InsertOneRow(ByVal n_sheets As Integer)
    n_sheets = n_sheets + 4
    Rows(n_sheets & ":" & n_sheets).Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    
    Rows(CStr(n_sheets - 1) & ":" & CStr(n_sheets - 1)).Select
    Selection.Copy
    Rows(CStr(n_sheets) & ":" & CStr(n_sheets)).Select
    ActiveSheet.Paste
    
    Application.CutCopyMode = False
End Sub

Private Sub ChangeCellData(ByVal nsheet As Integer, ByVal nselect As Integer)
    ' change sheet data direct to well sheet data value
    ' https://stackoverflow.com/questions/18744537/vba-setting-the-formula-for-a-cell
    
    Range("C2, C3, C4, C5, C6, C7, C8, C15, C16, C17, C18, C19, E17, F21").Select
    
    nsheet = nsheet + 3
    Selection.Replace What:=CStr(nselect + 3), Replacement:=CStr(nsheet), LookAt:=xlPart, _
                      SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                      ReplaceFormat:=False
    
    
    
    ' minhwasoo, 2023-10-13
    ' block, this code ....
    ' Range("E21").Select
    ' Range("E21").formula = "=Well!" & Cells(nsheet, "I").Address
End Sub



' --------------------------------------------------------------------------------------------------------------


Private Sub CommandButton1_Click()
UserFormTS.Show
End Sub

Function IsSheet(shname As String) As Boolean
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(shname)
    
    Application.DisplayAlerts = False
    IsSheet = True
    Application.DisplayAlerts = True
    
    Exit Function
    
ErrorHandler:
    ' MsgBox "An error occurred while trying to delete the worksheet."
    
    IsSheet = False
    Application.DisplayAlerts = True
End Function


Private Sub CommandButton5_Click()
' Get(Ec, Ph, Temp) Range - 지열공에서 통계내는 함수 ....
' Ph, EC, Temp statistics, find range
' data gathering function in EarthThermal test ...

    Dim nofwell, i As Integer
    
    Dim lowEC() As Double
    Dim hiEC() As Double
    Dim lowPH() As Double
    Dim hiPH() As Double
    Dim lowTEMP() As Double
    Dim hiTEMP() As Double

    nofwell = sheets_count()
    
'    If nofwell < 2 Or Not Contains(Sheets, "a1") Then
'        MsgBox "first Generate Simple YangSoo"
'        Exit Sub
'    End If
    
    If Not IsSheet("p1") Then
        MsgBox "First Make Summary Page"
        Exit Sub
    End If
    
 
    ReDim lowPH(1 To nofwell)
    ReDim hiPH(1 To nofwell)
    
    ReDim lowEC(1 To nofwell)
    ReDim hiEC(1 To nofwell)
    
    ReDim lowTEMP(1 To nofwell)
    ReDim hiTEMP(1 To nofwell)
    
    For i = 1 To nofwell
        lowEC(i) = getEC(cellLOW, i)
        hiEC(i) = getEC(cellHI, i)
        
        lowPH(i) = getPH(cellLOW, i)
        hiPH(i) = getPH(cellHI, i)
        
        lowTEMP(i) = getTEMP(cellLOW, i)
        hiTEMP(i) = getTEMP(cellHI, i)
    Next i
    
    Debug.Print String(3, vbCrLf)
    
    Debug.Print "--Temp----------------------------------------"
    Debug.Print "low : " & Application.min(lowTEMP), Application.max(lowTEMP)
    Debug.Print "hi  : " & Application.min(hiTEMP), Application.max(hiTEMP)
    Debug.Print "----------------------------------------------"
    
    Debug.Print "--PH------------------------------------------"
    Debug.Print "low : " & Application.min(lowPH), Application.max(lowPH)
    Debug.Print "hi  : " & Application.min(hiPH), Application.max(hiPH)
    Debug.Print "----------------------------------------------"
       
    Debug.Print "--EC------------------------------------------"
    Debug.Print "low : " & Application.min(lowEC), Application.max(lowEC)
    Debug.Print "hi  : " & Application.min(hiEC), Application.max(hiEC)
    Debug.Print "----------------------------------------------"

End Sub


Function getEC(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well, 2)
    remainder = well Mod 2
    page = quo + remainder
    
    Sheets("p" & CStr(page)).Activate
    
    If remainder = 1 Then
        If LOWHI = cellLOW Then
            getEC = Sheets("p" & CStr(page)).Range("e25").value
        Else
            getEC = Sheets("p" & CStr(page)).Range("e24").value
        End If
    Else
        If LOWHI = cellLOW Then
            getEC = Sheets("p" & CStr(page)).Range("h25").value
        Else
            getEC = Sheets("p" & CStr(page)).Range("h24").value
        End If
    End If
End Function

Function getPH(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well, 2)
    remainder = well Mod 2
    page = quo + remainder
    
    Sheets("p" & CStr(page)).Activate
    
    If remainder = 1 Then
        If LOWHI = cellLOW Then
            getPH = Sheets("p" & CStr(page)).Range("f25").value
        Else
            getPH = Sheets("p" & CStr(page)).Range("f24").value
        End If
    Else
        If LOWHI = cellLOW Then
            getPH = Sheets("p" & CStr(page)).Range("i25").value
        Else
            getPH = Sheets("p" & CStr(page)).Range("i24").value
        End If
    End If
End Function

Function getTEMP(ByVal LOWHI As Integer, ByVal well As Integer)
    Dim page, quo, remainder As Integer
    
    quo = WorksheetFunction.quotient(well, 2)
    remainder = well Mod 2
    page = quo + remainder
    
    Sheets("p" & CStr(page)).Activate
    
    If remainder = 1 Then
        If LOWHI = cellLOW Then
            getTEMP = Sheets("p" & CStr(page)).Range("d25").value
        Else
            getTEMP = Sheets("p" & CStr(page)).Range("d24").value
        End If
    Else
        If LOWHI = cellLOW Then
            getTEMP = Sheets("p" & CStr(page)).Range("g25").value
        Else
            getTEMP = Sheets("p" & CStr(page)).Range("g24").value
        End If
    End If
End Function






Sub DuplicateQ2Page(ByVal n As Integer)
' n : Q2 page 복사할 회수
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q2")
    
    For i = 1 To n
        ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
        ActiveSheet.Name = "p" & i
        
        With ActiveSheet.Tab
            .ThemeColor = xlThemeColorAccent3
            .TintAndShade = 0
        End With
        
        Call SetWellPropertyQ2(i)
    Next i
End Sub


Sub SetWellPropertyQ2(ByVal i As Integer)
' i : index of well

    ActiveSheet.Range("D12") = "W-" & CStr((i - 1) * 2 + 1)
    ActiveSheet.Range("D29") = "W-" & CStr((i - 1) * 2 + 1)
    
    ActiveSheet.Range("G12") = "W-" & CStr((i - 1) * 2 + 2)
    ActiveSheet.Range("H29") = "W-" & CStr((i - 1) * 2 + 2)
    
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
End Sub


Sub SetWellPropertyRest(ByVal w2page As Integer)
    Dim firstwell As Integer
      
    firstwell = 2 * w2page + 1
    
    ActiveSheet.Range("D12") = "W-" & CStr(firstwell)
    ActiveSheet.Range("H12") = "W-" & CStr(firstwell)
    
    ActiveSheet.Shapes.Range(Array("CommandButton3")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton4")).Select
    Selection.Delete
    ActiveSheet.Shapes.Range(Array("CommandButton5")).Select
    Selection.Delete
End Sub


Sub DuplicateRest(ByVal w2page As Integer)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Q1")
    
    ws.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)
    ActiveSheet.Name = "p" & CStr(w2page + 1)
    
    With ActiveSheet.Tab
        .ThemeColor = xlThemeColorAccent3
        .TintAndShade = 0
    End With
        
    Call SetWellPropertyRest(w2page)
    
End Sub


Function DivideWellsBy2(ByVal numberOfWells As Integer) As Integer()
    Dim quotient As Integer
    Dim remainder As Integer
    Dim result(1) As Integer
    
    quotient = (numberOfWells - 1) \ 2
    remainder = numberOfWells Mod 2
    
    
    If remainder = 0 Then
        result(0) = quotient + 1
    Else
        result(0) = quotient
    End If
    
    result(1) = remainder
    
    DivideWellsBy2 = result
End Function


Private Sub CommandButton3_Click()
' make summary page

    Dim result() As Integer
    Dim w2page, wselect, restpage As Integer
    'wselect = 1 --> only w1
       
    If IsSheet("p1") Then
        MsgBox "Sheet P1 Exist .... Delete First ... ", vbOKOnly
        Exit Sub
    End If
       
       
       
    result = DivideWellsBy2(sheets_count())
    
    ' result(0) = quotient
    ' result(1) = remainder
    
    w2page = result(0)
    restpage = result(1)
    
    Call DuplicateQ2Page(w2page)
    
    If restpage = 0 Then
        Exit Sub
    Else
        Call DuplicateRest(w2page)
    End If

End Sub


Private Sub CommandButton2_Click()
' get waterspec from yangsoo
  
  Dim thisname, fname1, fname2 As String
  Dim cell1, cell2 As String
  Dim time1 As Date
  Dim bTemp, ec1, ec2, ph1, ph2 As Double
  
  
  
  cell1 = Range("d12").value
  cell2 = Range("g12").value
  
  thisname = ActiveWorkbook.Name
  fname1 = "A" & GetNumeric2(cell1) & "_ge_OriginalSaveFile.xlsm"
  fname2 = "A" & GetNumeric2(cell2) & "_ge_OriginalSaveFile.xlsm"
  
  If Not IsWorkBookOpen(fname1) Then
    MsgBox "Please open the yangsoo data ! " & fname1
    Exit Sub
  End If
  
  If Not IsWorkBookOpen(fname2) Then
    MsgBox "Please open the yangsoo data ! " & fname2
    Exit Sub
  End If
  
  ' Range("k2") = fname1
  ' Range("k3") = fname2
  
  '------------------------------------------------------------------------
  time1 = Workbooks(fname1).Worksheets("w1").Range("c6").value
  bTemp = Workbooks(fname1).Worksheets("w1").Range("c7").value
  
  ec1 = Workbooks(fname1).Worksheets("w1").Range("c8").value
  ph1 = Workbooks(fname1).Worksheets("w1").Range("c9").value
  
  ec2 = Workbooks(fname2).Worksheets("w1").Range("c8").value
  ph2 = Workbooks(fname2).Worksheets("w1").Range("c9").value
  
  '------------------------------------------------------------------------
  
  
  Range("c6").value = time1
  Range("c7").value = bTemp
  Range("c8").value = ec1
  Range("c9").value = ph1
  
  Range("d8").value = ec2
  Range("d9").value = ph2
  
  
  Call getModDataFromYangSooDual(thisname, fname1)
  Call getModDataFromYangSooDual(thisname, fname2)
  
  
End Sub


Sub getModDataFromYangSooDual(ByVal thisname As String, ByVal fname As String)

    Dim f As Integer

    f = CInt(GetNumeric2(fname)) Mod 2

    Windows(fname).Activate
    Sheets("w1").Activate
    Sheets("w1").Range("H14:J23").Select
    Selection.Copy
    
    Windows(thisname).Activate
    
    If f = 0 Then
        Range("h31").Select
    Else
        Range("d31").Select
    End If
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
End Sub


Private Sub DeleteWorksheet(shname As String)
    Dim ws As Worksheet
    
    On Error GoTo ErrorHandler
    
    Set ws = ThisWorkbook.Worksheets(shname)
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    
    Exit Sub
    
ErrorHandler:
    ' MsgBox "An error occurred while trying to delete the worksheet."
    Application.DisplayAlerts = True
End Sub



Private Sub CommandButton4_Click()
' delete all summary page

    Dim i, nofwell, pn As Integer
    Dim response As VbMsgBoxResult
        
    nofwell = GetNumberOfWell()
    pn = (nofwell \ 2) + (nofwell Mod 2)
    
    Sheets("Q2").Activate
    
    response = MsgBox("Do you deletel all water well?", vbYesNo)
    If response = vbYes Then
        For i = 1 To pn
            Call DeleteWorksheet("p" & i)
        Next i
    End If
    
    Sheets("Q2").Activate
    
End Sub




Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Function GetNowLast(inputDate As Date) As Date

    Dim dYear, dMonth, getDate As Date

    dYear = Year(inputDate)
    dMonth = Month(inputDate)

    getDate = DateSerial(dYear, dMonth + 1, 0)

    GetNowLast = getDate

End Function

Private Sub ComboBoxFix(ByVal SIGN As Boolean)

    Dim contr As Control
    
    If SIGN Then
        For Each contr In UserFormTS.Controls
            If TypeName(contr) = "ComboBox" Then
                contr.Style = fmStyleDropDownList
            End If
        Next
    Else
        For Each contr In UserFormTS.Controls
            If TypeName(contr) = "ComboBox" Then
                contr.Style = fmStyleDropDownCombo
            End If
        Next
    End If

End Sub

Private Function whichSection(n As Integer) As Integer

    whichSection = Round((n / 10), 0) * 10

End Function

Private Sub ComboBoxYear_Initialize()
    Dim nYear, nMonth, nDay As Integer
    Dim nHour, nMin As Integer
    
    Dim i, j As Integer
    Dim lastDay As Integer
    
    Dim sheetDate, currDate As Date
    Dim isThisYear As Boolean
    
    sheetDate = Range("c6").value
    'MsgBox (sheetDate)
    currDate = Now()
    
    If ((Year(currDate) - Year(sheetDate)) = 0) Then
    
        isThisYear = True
        
        nYear = Year(sheetDate)
        nMonth = Month(sheetDate)
        nDay = Day(sheetDate)
        
        nHour = Hour(sheetDate)
        nMin = Minute(sheetDate)
        
    Else
        
        isThisYear = False
        
        nYear = Year(currDate)
        nMonth = Month(currDate)
        nDay = Day(currDate)
        
        nHour = Hour(currDate)
        nMin = Minute(currDate)
            
    End If
    
    
    lastDay = Day(GetNowLast(IIf(isThisYear, sheetDate, currDate)))
    Debug.Print lastDay
    
    For i = nYear - 10 To nYear
        ComboBoxYear.AddItem (i)
    Next i
    
    For i = 1 To 12
        ComboBoxMonth.AddItem (i)
    Next i
    
    For i = 1 To lastDay
        ComboBoxDay.AddItem (i)
    Next i
    
            
    For i = 1 To 12
        ComboBoxHour.AddItem (i)
    Next i
    
    
    
    For i = 0 To 60 Step 10
        ComboBoxMinute.AddItem (i)
    Next i
    
    
    
    ComboBoxYear.value = nYear
    ComboBoxMonth.value = nMonth
    ComboBoxDay.value = nDay
    
    ComboBoxHour.value = IIf(nHour > 12, nHour - 12, nHour)
    ComboBoxMinute.value = whichSection(IIf(isThisYear, Minute(sheetDate), Minute(currDate)))
    
   
    If nHour > 12 Then
        OptionButtonPM.value = True
    Else
        OptionButtonAM.value = True
    End If
    
    Debug.Print nYear

End Sub

Sub ComboboxDay_ChangeItem(nYear As Integer, nMonth As Integer)
    Dim lasday, i As Integer
    
    lasday = Day(GetNowLast(DateSerial(nYear, nMonth, 1)))
    ComboBoxDay.Clear
    
    For i = 1 To lasday
        ComboBoxDay.AddItem (i)
    Next i
    
    ComboBoxDay.value = 1

End Sub

Private Sub ComboBoxHour_Change()
    ComboBoxMinute.value = 0
End Sub

Private Sub ComboBoxMonth_Change()
    '2019-11-26 change
    On Error GoTo Errcheck
    Call ComboboxDay_ChangeItem(ComboBoxYear.value, ComboBoxMonth.value)
Errcheck:
        
End Sub

Private Sub EnterButton_Click()
    Dim nYear, nMonth, nDay As Integer
    Dim nHour, nMinute As Integer
    
    Dim nDate, nTime As Date
    
    
    On Error GoTo Errcheck
    nYear = ComboBoxYear.value
    nMonth = ComboBoxMonth.value
    nDay = ComboBoxDay.value
        
    nHour = ComboBoxHour.value
    nMinute = ComboBoxMinute.value
            
            
    nHour = nHour + IIf(OptionButtonPM.value, 12, 0)
            
    nDate = DateSerial(nYear, nMonth, nDay)
    nTime = TimeSerial(nHour, nMinute, 0)
            
    nDate = nDate + nTime
         
    Range("c6").value = nDate
         
Errcheck:
     
    Unload Me
     
End Sub

Private Sub UserForm_Initialize()
    Call ComboBoxYear_Initialize
    
End Sub

Option Explicit

Public Sub TurnOffStuff()
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
End Sub

Public Sub TurnOnStuff()
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub


Function UppercaseString(inputString As String) As String
    UppercaseString = UCase(inputString)
End Function



Public Sub Range_End_Method()
    'Finds the last non-blank cell in a single row or column
    
    Dim lRow        As Long
    Dim lCol        As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.count, 1).End(xlUp).Row
    
    'Find the last non-blank cell in row 1
    lCol = Cells(1, Columns.count).End(xlToLeft).column
    
    MsgBox "Last Row: " & lRow & vbNewLine & _
           "Last Column: " & lCol
End Sub

Public Function lastRow() As Long
    Dim lRow        As Long
    lRow = Cells(Rows.count, 1).End(xlUp).Row
    
    lastRow = lRow
End Function

'Public Function Contains(Col As Collection, key As Variant) As Boolean
'    On Error Resume Next
'    Col (key)                                    ' Just try it. If it fails, Err.Number will be nonzero.
'    Contains = (Err.number = 0)
'    Err.Clear
'End Function

Function Contains(objCollection As Object, strName As String) As Boolean
    Dim o           As Object
    On Error Resume Next
    
    Set o = objCollection(strName)
    Contains = (Err.Number = 0)
    Err.Clear
End Function

Function RemoveDupesDict(myArray As Variant) As Variant
    'DESCRIPTION: Removes duplicates from your array using the dictionary method.
    'NOTES: (1.a) You must add a reference to the Microsoft Scripting Runtime library via
    ' the Tools > References menu.
    ' (1.b) This is necessary because I use Early Binding in this function.
    ' Early Binding greatly enhances the speed of the function.
    ' (2) The scripting dictionary will not work on the Mac OS.
    'SOURCE: https://wellsr.com
    '-----------------------------------------------------------------------
    Dim i           As Long
    Dim d           As Scripting.Dictionary
    Set d = New Scripting.Dictionary
    
    With d
        For i = LBound(myArray) To UBound(myArray)
            If IsMissing(myArray(i)) = False Then
                .item(myArray(i)) = 1
            End If
        Next
        RemoveDupesDict = .Keys
    End With
End Function

Public Function GetLength(A As Variant) As Integer
    ' if array is empty return 0
    ' else return number of array item
    
    If IsEmpty(A) Then
        GetLength = 0
    Else
        GetLength = UBound(A) - LBound(A) + 1
    End If
End Function

Public Function getUnique(ByRef array_tabcolor As Variant) As Variant
    ' remove duplicated item in array
    ' and return unique array value
    
    Dim array_size  As Variant
    Dim new_array   As Variant
    
    new_array = RemoveDupesDict(array_tabcolor)
    getUnique = new_array
End Function

Function ConvertToLongInteger(ByVal stValue As String) As Long
    On Error GoTo ConversionFailureHandler
    ConvertToLongInteger = CLng(stValue)        'TRY to convert to an Integer value
    Exit Function        'If we reach this point, then we succeeded so exit
    
ConversionFailureHandler:
    'IF we've reached this point, then we did not succeed in conversion
    'If the error is type-mismatch, clear the error and return numeric 0 from the function
    'Otherwise, disable the error handler, and re-run the code to allow the system to
    'display the error
    If Err.Number = 13 Then        'error # 13 is Type mismatch
        Err.Clear
        ConvertToLongInteger = 0
        Exit Function
    Else
        On Error GoTo 0
        Resume
    End If
End Function


Option Explicit

'------------------------------------------------------------------------------------------
' 2022/6/11

Public Enum cellLowHi
    cellLOW = 0
    cellHI = 1
End Enum

'Function GetNumberOfWell() As Integer
'    Dim save_name As String
'    Dim n As Integer
'
'    save_name = ActiveSheet.Name
'    Sheets("Well").Activate
'    Sheets("Well").Range("A30").Select
'    Selection.End(xlUp).Select
'    n = CInt(GetNumeric2(Selection.value))
'
'    GetNumberOfWell = n
'End Function


' 이것은, Well 탭의 값을 가지고 검사하하는것이라서, 차이가 생긴다.
Function GetNumberOfWell() As Integer
    Dim save_name As String
    Dim n As Integer
    
    save_name = ActiveSheet.Name
    With Sheets("Well")
        n = .Cells(.Rows.count, "A").End(xlUp).Row
        n = CInt(GetNumeric2(.Cells(n, "A").value))
    End With
    
    GetNumberOfWell = n
End Function


'Public Function sheets_count() As Long
'    Dim i, nSheetsCount, nWell  As Integer
'    Dim strSheetsName(50) As String
'
'    nSheetsCount = ThisWorkbook.Sheets.count
'    nWell = 0
'
'    For i = 1 To nSheetsCount
'        strSheetsName(i) = ThisWorkbook.Sheets(i).Name
'        'MsgBox (strSheetsName(i))
'        If (ConvertToLongInteger(strSheetsName(i)) <> 0) Then
'            nWell = nWell + 1
'        End If
'    Next
'
'    'MsgBox (CStr(nWell))
'    sheets_count = nWell
'End Function

Public Function sheets_count() As Long
    Dim i As Integer
    Dim nSheetsCount As Long
    Dim nWell As Long
    Dim strSheetsName() As String

    nSheetsCount = ThisWorkbook.Sheets.count
    nWell = 0

    ReDim strSheetsName(1 To nSheetsCount)

    For i = 1 To nSheetsCount
        strSheetsName(i) = ThisWorkbook.Sheets(i).Name
        If ConvertToLongInteger(strSheetsName(i)) <> 0 Then
            nWell = nWell + 1
        End If
    Next i

    sheets_count = nWell
End Function


Function GetNumeric2(ByVal CellRef As String)
    Dim StringLength, i  As Integer
    Dim result      As String
    
    StringLength = Len(CellRef)
    For i = 1 To StringLength
        If IsNumeric(Mid(CellRef, i, 1)) Then result = result & Mid(CellRef, i, 1)
    Next i
    GetNumeric2 = result
End Function

'********************************************************************************************************************************************************************************
'Function Name                    : IsWorkBookOpen(ByVal OWB As String)
'Function Description             : Function to check whether specified workbook is open
'Data Parameters                  : OWB:- Specify name or path to the workbook. eg: "Book1.xlsx" or "C:\Users\Kannan.S\Desktop\Book1.xlsm"

'********************************************************************************************************************************************************************************
Function IsWorkBookOpen(ByVal OWB As String) As Boolean
    IsWorkBookOpen = False
    Dim WB          As Excel.Workbook
    Dim WBName      As String
    Dim WBPath      As String
    Dim OWBArray    As Variant
    
    Err.Clear
    
    On Error Resume Next
    OWBArray = Split(OWB, Application.PathSeparator)
    Set WB = Application.Workbooks(OWBArray(UBound(OWBArray)))
    WBName = OWBArray(UBound(OWBArray))
    WBPath = WB.Path & Application.PathSeparator & WBName
    
    If Not WB Is Nothing Then
        If UBound(OWBArray) > 0 Then
            If LCase(WBPath) = LCase(OWB) Then IsWorkBookOpen = True
        Else
            IsWorkBookOpen = True
        End If
    End If
    Err.Clear
    
End Function

'------------------------------------------------------------------------------------------

Public Function GetLengthByColor(ByVal tabColor As Variant) As Integer
    Dim n_sheets, i, j, nTab As Integer
    n_sheets = sheets_count()
    
    nTab = 0
    
    For i = 1 To n_sheets
        If (Sheets(CStr(i)).Tab.Color = tabColor) Then
            nTab = nTab + 1
        End If
    Next i
    
    GetLengthByColor = nTab
End Function

Private Sub get_tabsize_by_well(ByRef nof_sheets As Integer, ByRef nof_unique_tab As Variant, ByRef n_tabcolors As Variant)
    ' n_tabcolors : return value
    ' nof_unique_tab : return value
    
    Dim n_sheets, i, j As Integer
    Dim limit()     As Integer
    Dim arr_tabcolors(), new_tabcolors() As Variant
    
    n_sheets = sheets_count()
    
    ReDim arr_tabcolors(1 To n_sheets)
    ReDim new_tabcolors(1 To n_sheets)
    ReDim limit(0 To n_sheets)
    
    For i = 1 To n_sheets
        arr_tabcolors(i) = Sheets(CStr(i)).Tab.Color
    Next i
    
    new_tabcolors = getUnique(arr_tabcolors)
    
    For i = 0 To UBound(new_tabcolors)
        limit(i) = GetLengthByColor(new_tabcolors(i))
    Next i
    
    nof_sheets = n_sheets
    nof_unique_tab = limit
    n_tabcolors = new_tabcolors
End Sub
Option Explicit

Dim ColorValue(1 To 20) As Long

Public Sub InitialSetColorValue()
    ColorValue(1) = RGB(192, 0, 0)
    ColorValue(2) = RGB(255, 0, 0)
    ColorValue(3) = RGB(255, 192, 0)
    ColorValue(4) = RGB(255, 255, 0)
    ColorValue(5) = RGB(146, 208, 80)
    ColorValue(6) = RGB(0, 176, 80)
    ColorValue(7) = RGB(0, 176, 240)
    ColorValue(8) = RGB(0, 112, 192)
    ColorValue(9) = RGB(0, 32, 96)
    ColorValue(10) = RGB(112, 48, 160)
    
    ColorValue(11) = RGB(192 + 10, 10, 0)
    ColorValue(12) = RGB(255, 0 + 10, 0)
    ColorValue(13) = RGB(255, 192 + 10, 0)
    ColorValue(14) = RGB(255, 255, 10)
    ColorValue(15) = RGB(146 + 10, 208 + 10, 80 + 10)
    ColorValue(16) = RGB(0 + 10, 176 + 10, 80)
    ColorValue(17) = RGB(0 + 10, 176 + 10, 240 + 10)
    ColorValue(18) = RGB(0 + 10, 112 + 10, 192)
    ColorValue(19) = RGB(0 + 10, 32 + 10, 96)
    ColorValue(20) = RGB(112, 48 + 10, 160 + 10)
End Sub

Private Sub initialize_wellstyle()
    Range("C3:C22").Select
    Selection.NumberFormat = "General"
        
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    With Selection.Font
        .Name = "맑은 고딕"
        .Size = 10
        .ThemeColor = xlThemeColorLight1
    End With
    
    Range("E19:G19").Select
    With Selection.Font
        .Name = "맑은 고딕"
        .Size = 12
        .ThemeColor = xlThemeColorLight1
        .ThemeFont = xlThemeFontNone
    End With
    
    Range("E21:G21").Select
    With Selection.Font
        .Name = "맑은 고딕"
        .Size = 12
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Range("B25:K29").Select
    With Selection.Font
        .Name = "맑은 고딕"
        .Size = 11
        .TintAndShade = 0
        .ThemeFont = xlThemeFontNone
    End With
    
    Range("d23").Select
End Sub

Private Sub change_font_size()
    Range("J25").Select
    Selection.Font.Size = 10
    Range("F26").Select
    Selection.Font.Size = 10
End Sub

Public Sub make_wellstyle()
    Dim n_sheets    As Integer
    Dim i           As Integer
    
    n_sheets = sheets_count()
    
    Call TurnOffStuff
    
    For i = 1 To n_sheets
        Worksheets(CStr(i)).Activate
        Call initialize_wellstyle
        Call change_font_size
    Next i
    
    Call TurnOnStuff
    
End Sub

Private Sub JojungData(ByVal nsheet As Integer)
    Dim nselect     As String
    
    Range("C2, C3, C4, C5, C6, C7, C8, C15, C16, C17, C18, C19, E17, F21").Select
    Range("F21").Activate
    
    nsheet = nsheet + 3
    '=Well!D7
    nselect = Mid(Range("c2").formula, 8)
    
    'Debug.Print Mid(Range("c2").Formula, 8) & ":" & nselect
    
    Selection.Replace What:=nselect, Replacement:=CStr(nsheet), LookAt:=xlPart, _
                      SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
                      ReplaceFormat:=False
    
    ' minhwasoo 2023/10/13
    ' Range("E21").Select
    ' Range("E21").formula = "=Well!" & Cells(nsheet, "I").Address
End Sub

Private Sub SetMyTabColor(ByVal index As Integer)
    If Sheets("Well").SingleColor.value Then
        With ActiveWorkbook.Sheets(CStr(index)).Tab
            .Color = 192
            .TintAndShade = 0
        End With
    Else
        With ActiveWorkbook.Sheets(CStr(index)).Tab
            .Color = ColorValue(index)
            .TintAndShade = 0
        End With
    End If
End Sub

'각각의 쉬트를 순회하면서, 셀의 참조값을 맟추어준다.
'
Public Sub JojungSheetData()
    Dim n_sheets    As Integer
    Dim i           As Integer
    
    n_sheets = sheets_count()
    
    For i = 1 To n_sheets
        Cells(i + 3, "A").value = "W" & i
    Next i
    
    For i = 1 To n_sheets
        Sheets(CStr(i)).Activate
        Call JojungData(i)
        Call SetMyTabColor(i)
    Next i
End Sub

Option Explicit

Private Sub HideLocation_Click()
  Sheets("location").Visible = False
  Sheets("Recharge").Activate
End Sub
Private Sub CommandButton1_Click()
    
    Sheets("Aggregate1").Visible = False
    Sheets("Well").Select
    
End Sub

'q1 - 한계양수량 - b13
'q2 - 가채수량 - b7
'q3 - 취수계획량 - b15
'ratio - b11
'qq1 - 1단계 양수량

Private Sub CommandButton2_Click()
' Collect Data

    Dim fname As String
    Dim nofwell, i As Integer
    Dim q1() As Double
    Dim qq1() As Double
    Dim q2() As Double
    Dim q3() As Double
    
    Dim ratio() As Double
    
    Dim C() As Double
    Dim B() As Double
    
    Dim S1() As Double
    Dim S2() As Double
    
    
    nofwell = GetNumberOfWell()
    Sheets("Aggregate1").Select
    
    ReDim q1(1 To nofwell) '한계양수량
    ReDim q2(1 To nofwell) '적정취수량
    ReDim q3(1 To nofwell) '취수계획량
    ReDim qq1(1 To nofwell) '1단계 양수량
    
    ReDim ratio(1 To nofwell)
    
    ReDim C(1 To nofwell)
    ReDim B(1 To nofwell)
    
    ReDim S1(1 To nofwell)
    ReDim S2(1 To nofwell)
    
    
     With Range("g3:k19")
        .value = ""
    End With
        
     With Range("q3:s19")
        .value = ""
    End With
    
    
    For i = 1 To nofwell
        q1(i) = Worksheets("YangSoo").Cells(4 + i, "aa").value
        qq1(i) = Worksheets("YangSoo").Cells(4 + i, "ac").value
        
        q2(i) = Worksheets("YangSoo").Cells(4 + i, "ab").value
        q3(i) = Worksheets("YangSoo").Cells(4 + i, "k").value
        
        ratio(i) = Worksheets("YangSoo").Cells(4 + i, "ah").value
        
        S1(i) = Worksheets("YangSoo").Cells(4 + i, "ad").value
        S2(i) = Worksheets("YangSoo").Cells(4 + i, "ae").value
        
        C(i) = Worksheets("YangSoo").Cells(4 + i, "af").value
        B(i) = Worksheets("YangSoo").Cells(4 + i, "ag").value
        
    Next i

    Call WriteWellData36(q1, q2, q3, ratio, C, B, nofwell)
    Call TransPoseWellData(nofwell)
    Call Write_Tentative_water_intake(qq1, S2, S1, q2, nofwell)
    
    
    Application.CutCopyMode = False
End Sub


Sub Write_Tentative_water_intake(q1 As Variant, S2 As Variant, S1 As Variant, q2 As Variant, nofwell As Variant)
    Dim i As Integer
    
    ' Range("E32:F67").Clear
    
    With Range("f27:i52")
        .value = ""
    End With
        
    For i = 1 To nofwell
        Cells(27 + (i - 1) * 2, "F").value = "W-" & CStr(i)
    
        Cells(27 + (i - 1) * 2, "G").value = q1(i)
        
        Cells(27 + (i - 1) * 2, "H").value = S2(i)
        Cells(28 + (i - 1) * 2, "H").value = S1(i)
    
        Cells(27 + (i - 1) * 2, "I").value = q2(i)
    Next i
End Sub

Sub TransPoseWellData(ByVal nofwell As Integer)
    
'    Range("i3:i" & (nofwell + 2)).Select
'    Selection.Copy
'    Range("M23").Select
'    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
'        False, Transpose:=True
'
'    Range("j3:j" & (nofwell + 2)).Select
'    Application.CutCopyMode = False
'    Selection.Copy
'    Range("M24").Select
'    Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
'        False, Transpose:=True

End Sub


Sub WriteWellData36(q1 As Variant, q2 As Variant, q3 As Variant, ratio As Variant, C As Variant, B As Variant, ByVal nofwell As Integer)
    
    Dim i As Integer
    
    For i = 1 To nofwell
        Range("G" & (i + 2)).value = "W-" & i
        Range("H" & (i + 2)).value = q1(i)
        Range("I" & (i + 2)).value = q2(i)
        Range("J" & (i + 2)).value = q3(i)
        Range("K" & (i + 2)).value = ratio(i)
        
        Range("Q" & (i + 2)).value = "W-" & i
        Range("R" & (i + 2)).value = C(i)
        Range("S" & (i + 2)).value = B(i)
    Next i
   
    Range("N3").value = Application.min(ratio)
    Range("O3").value = Application.max(ratio)
    
    Range("N4").value = Application.min(q2)
    Range("O4").value = Application.max(q2)
    
    Range("N5").value = Application.min(q3)
    Range("O5").value = Application.max(q3)

End Sub



Private Sub CommandButton3_Click()
    Range("g3:k19").Select
    Selection.ClearContents
    
    Range("n3:o5").Select
    Selection.ClearContents
     
    Range("q3:s19").Select
    Selection.ClearContents
    
    Range("B24").Select
    Application.CutCopyMode = False
End Sub
Private Sub CommandButton1_Click()
    Sheets("Aggregate2").Visible = False
    Sheets("Well").Select
End Sub

Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub



Private Sub CommandButton2_Click()
' Collect Data

    Dim fname As String
    Dim nofwell, i As Integer
    
    Dim Q() As Double          '양수량
    Dim natural() As Double    '자연수위
    Dim stable() As Double      '안정수위
    Dim recover() As Double     '회복수위
    
    Dim radius() As Double       ' 공반경
    Dim deltas() As Double       ' deltas
    Dim deltah() As Double       ' deltah : 수위강하량
    Dim daeSoo() As Double       ' 대수층 두께
    
    Dim T1() As Double            ' T1
    Dim T2() As Double            ' T2
    Dim TA() As Double            ' TA - (T1+T2)/2, TAverage
    
    Dim K() As Double
    Dim time_() As Double           ' 안정수위도달시간
    
    Dim S1() As Double            ' S1
    Dim S2() As Double            ' S2 - 스킨팩터 해석, s값
    
    Dim schultz() As Double
    Dim webber() As Double
    Dim jcob() As Double
    
    Dim skin() As Double ' skin factor
    Dim er() As Double   ' effective radius
    

    nofwell = GetNumberOfWell()
    Sheets("Aggregate2").Select
    
    ' --------------------------------------------------------------------------------------
    
    ReDim Q(1 To nofwell)
    ReDim natural(1 To nofwell)
    ReDim stable(1 To nofwell)
    ReDim recover(1 To nofwell)
    
    ReDim radius(1 To nofwell)
    ReDim deltas(1 To nofwell)
    ReDim deltah(1 To nofwell)
    ReDim daeSoo(1 To nofwell)
    
    ReDim T1(1 To nofwell)
    ReDim T2(1 To nofwell)
    ReDim TA(1 To nofwell)
    ReDim K(1 To nofwell)
    ReDim time_(1 To nofwell)
    
    ReDim S1(1 To nofwell)
    ReDim S2(1 To nofwell)
    
    ReDim shultz(1 To nofwell)
    ReDim webber(1 To nofwell)
    ReDim jcob(1 To nofwell)
    
    ReDim skin(1 To nofwell) ' skin factor
    ReDim er(1 To nofwell)   ' effective radius
    
    ' --------------------------------------------------------------------------------------
    
    Call EraseCellData("d3:j16")
    Call EraseCellData("l3:q16")
    Call EraseCellData("s3:u16")
            
            
    For i = 1 To nofwell
   
        Q(i) = Worksheets("YangSoo").Cells(4 + i, "k").value
        
        natural(i) = Worksheets("YangSoo").Cells(4 + i, "b").value
        stable(i) = Worksheets("YangSoo").Cells(4 + i, "c").value
        recover(i) = Worksheets("YangSoo").Cells(4 + i, "d").value
        
        radius(i) = Worksheets("YangSoo").Cells(4 + i, "h").value
        
        deltas(i) = Worksheets("YangSoo").Cells(4 + i, "l").value
        deltah(i) = Worksheets("YangSoo").Cells(4 + i, "f").value
        daeSoo(i) = Worksheets("YangSoo").Cells(4 + i, "n").value
        
        
        T1(i) = Worksheets("YangSoo").Cells(4 + i, "o").value
        T2(i) = Worksheets("YangSoo").Cells(4 + i, "p").value
        TA(i) = Worksheets("YangSoo").Cells(4 + i, "q").value
        
        time_(i) = Worksheets("YangSoo").Cells(4 + i, "u").value
                
        S1(i) = Worksheets("YangSoo").Cells(4 + i, "r").value
        S2(i) = Worksheets("YangSoo").Cells(4 + i, "s").value
        K(i) = Worksheets("YangSoo").Cells(4 + i, "t").value
        
        shultz(i) = Worksheets("YangSoo").Cells(4 + i, "v").value
        webber(i) = Worksheets("YangSoo").Cells(4 + i, "w").value
        jcob(i) = Worksheets("YangSoo").Cells(4 + i, "x").value
        
        
        skin(i) = Worksheets("YangSoo").Cells(4 + i, "y").value
        er(i) = Worksheets("YangSoo").Cells(4 + i, "z").value
    
    Next i

    Call WriteWellData(Q, natural, stable, recover, radius, deltas, daeSoo, T1, S1, nofwell)
    Call WriteData37_RadiusOfInfluence(TA, K, S2, time_, deltah, daeSoo, nofwell)
    Call WriteData36_TS_Analysis(T1, T2, TA, S2, nofwell)
    Call Write38_RadiusOfInfluence_Result(shultz, webber, jcob, nofwell)
    Call Wrote34_SkinFactor(skin, er, nofwell)
        
    Range("a1").Select
    Application.CutCopyMode = False
    
End Sub


' 3-3, 3-4, 3-5 결과출력
Sub WriteWellData(Q As Variant, natural As Variant, stable As Variant, recover As Variant, radius As Variant, deltas As Variant, daeSoo As Variant, T1 As Variant, S1 As Variant, ByVal nofwell As Integer)
    
    Dim i As Integer
    
    For i = 1 To nofwell
    
        ' 3-3, 장기양수시험결과 (Collect from yangsoo data)
        Range("D" & (i + 2)).value = "W-" & i
        
        Range("e" & (i + 2)).value = Q(i)
        Range("l" & (i + 2)).value = Q(i)
        
        Range("f" & (i + 2)).value = natural(i)
        Range("g" & (i + 2)).value = stable(i)
        Range("h" & (i + 2)).value = stable(i) - natural(i)
        
        Range("i" & (i + 2)).value = radius(i)
        Range("j" & (i + 2)).value = deltas(i)
        
        
        ' 3-4, aqtesolv 해석결과
        Range("m" & (i + 2)).value = radius(i)
        Range("n" & (i + 2)).value = radius(i)
        Range("o" & (i + 2)).value = daeSoo(i)
        Range("p" & (i + 2)).value = T1(i)
        Range("q" & (i + 2)).value = S1(i)
        
        
        '3-5, 수위회복시험 결과
        Range("s" & (i + 2)).value = stable(i)
        Range("t" & (i + 2)).value = recover(i)
        Range("u" & (i + 2)).value = stable(i) - recover(i)
    Next i
   
End Sub



' 3.4 스킨계수
Sub Wrote34_SkinFactor(skin As Variant, er As Variant, nofwell As Variant)
    Dim i As Integer
    
    Call EraseCellData("o32:r45")
    
    For i = 1 To nofwell
        Cells(32 + (i - 1), "p").value = "W-" & CStr(i)
           
        Cells(32 + (i - 1), "q").value = skin(i)
        Cells(32 + (i - 1), "q").NumberFormat = "0.0000"
        
        Cells(32 + (i - 1), "r").value = er(i)
        Cells(32 + (i - 1), "r").NumberFormat = "0.000"
    Next i
End Sub


' 3-6, 수리상수산정결과
Sub WriteData36_TS_Analysis(T1 As Variant, T2 As Variant, TA As Variant, S2 As Variant, nofwell As Variant)
    Dim i As Integer
    
    Call EraseCellData("c32:f70")
    
    For i = 1 To nofwell
        Cells(32 + (i - 1) * 3, "C").value = "W-" & CStr(i)
        
        
        Cells(32 + (i - 1) * 3, "D").value = "장기양수시험"
        Cells(33 + (i - 1) * 3, "D").value = "수위회복시험"
        Cells(34 + (i - 1) * 3, "D").value = "선택치"
    
        Cells(32 + (i - 1) * 3, "E").value = T1(i)
        Cells(32 + (i - 1) * 3, "E").NumberFormat = "0.0000"
        
        Cells(33 + (i - 1) * 3, "E").value = T2(i)
        Cells(33 + (i - 1) * 3, "E").NumberFormat = "0.0000"
        
        Cells(34 + (i - 1) * 3, "E").value = TA(i)
        Cells(34 + (i - 1) * 3, "E").NumberFormat = "0.0000"
        Cells(34 + (i - 1) * 3, "E").Font.Bold = True
        
        Cells(32 + (i - 1) * 3, "F").value = S2(i)
        Cells(32 + (i - 1) * 3, "F").NumberFormat = "0.0000000"
        
        Cells(34 + (i - 1) * 3, "F").value = S2(i)
        Cells(34 + (i - 1) * 3, "F").NumberFormat = "0.0000000"
        Cells(34 + (i - 1) * 3, "F").Font.Bold = True
    Next i
End Sub


' 3-7, 조사공별 수리상수
Sub WriteData37_RadiusOfInfluence(TA As Variant, K As Variant, S2 As Variant, time_ As Variant, deltah As Variant, daeSoo As Variant, nofwell As Variant)

    Dim i As Integer
       
    Call EraseCellData("e21:s27")
    
    For i = 1 To nofwell
        Cells(21, Chr(Asc("D") + i)).value = "W-" & CStr(i)
        Cells(22, Chr(Asc("D") + i)).value = TA(i)
        Cells(22, Chr(Asc("D") + i)).NumberFormat = "0.0000"
        
        Cells(23, Chr(Asc("D") + i)).value = K(i)
        Cells(23, Chr(Asc("D") + i)).NumberFormat = "0.0000"
        
        
        Cells(24, Chr(Asc("D") + i)).value = S2(i)
        Cells(24, Chr(Asc("D") + i)).NumberFormat = "0.0000000"
        
        Cells(25, Chr(Asc("D") + i)).value = time_(i)
        Cells(25, Chr(Asc("D") + i)).NumberFormat = "0.0000"
        
        Cells(26, Chr(Asc("D") + i)).value = deltah(i)
        Cells(26, Chr(Asc("D") + i)).NumberFormat = "0.00"
        
        Cells(27, Chr(Asc("D") + i)).value = daeSoo(i)
    Next i

End Sub

'3.8 영향반경
Sub Write38_RadiusOfInfluence_Result(shultz As Variant, webber As Variant, jcob As Variant, nofwell As Variant)

    Dim i As Integer
    Dim sum, average As Double
    
    Call EraseCellData("h32:n45")
    
    For i = 1 To nofwell
        Cells(32 + (i - 1), "h").value = "W-" & CStr(i)
        Cells(32 + (i - 1), "h").NumberFormat = "0.0"
        
        Cells(32 + (i - 1), "i").value = shultz(i)
        Cells(32 + (i - 1), "i").NumberFormat = "0.0"
        
        Cells(32 + (i - 1), "j").value = webber(i)
        Cells(32 + (i - 1), "j").NumberFormat = "0.0"
        
        Cells(32 + (i - 1), "k").value = jcob(i)
        Cells(32 + (i - 1), "k").NumberFormat = "0.0"
    
        Cells(32 + (i - 1), "l").value = Round((shultz(i) + webber(i) + jcob(i)) / 3, 1)
        Cells(32 + (i - 1), "l").NumberFormat = "0.0"
        
        Cells(32 + (i - 1), "m").value = Application.WorksheetFunction.max(shultz(i), webber(i), jcob(i))
        Cells(32 + (i - 1), "m").NumberFormat = "0.0"
        
        Cells(32 + (i - 1), "n").value = Application.WorksheetFunction.min(shultz(i), webber(i), jcob(i))
        Cells(32 + (i - 1), "n").NumberFormat = "0.0"
    Next i

End Sub



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
        GetER_Mode = val(r)
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
        Case erRE1
            GetEffectiveRadius = Workbooks(WB_NAME).Worksheets("SkinFactor").Range("k8").value
        Case erRE2
            GetEffectiveRadius = Workbooks(WB_NAME).Worksheets("SkinFactor").Range("k9").value
        Case erRE3
            GetEffectiveRadius = Workbooks(WB_NAME).Worksheets("SkinFactor").Range("k10").value
        Case Else
            GetEffectiveRadius = Workbooks(WB_NAME).Worksheets("SkinFactor").Range("C8").value
    End Select

End Function



Option Explicit

Private Sub CommandButton1_Click()
    Sheets("aggWhpa").Visible = False
    Sheets("Well").Select
End Sub


Function getDirectionFromWell(i) As Integer

    If Sheets(CStr(i)).Range("k12").Font.Bold Then
        getDirectionFromWell = Sheets(CStr(i)).Range("k12").value
    Else
        getDirectionFromWell = Sheets(CStr(i)).Range("l12").value
    End If

End Function



Private Sub CommandButton2_Click()
' Collect Data

    Dim fname As String
    Dim nofwell, i As Integer
    
    Dim Q() As Double
    Dim daeSoo() As Double
    
    Dim T1() As Double
    Dim S1() As Double
    
    Dim direction() As Integer
    Dim gradient() As Double
    
    
    nofwell = GetNumberOfWell()
    If ActiveSheet.Name <> "aggWhpa" Then Sheets("aggWhpa").Select
    
    ReDim Q(1 To nofwell) As Double
    ReDim daeSoo(1 To nofwell) As Double
    
    ReDim T1(1 To nofwell) As Double
    ReDim S1(1 To nofwell) As Double
    
    ReDim direction(1 To nofwell) As Integer
    ReDim gradient(1 To nofwell) As Double
      

    ' --------------------------------------------------------------------------------------
    
    For i = 1 To nofwell
        
        Sheets(CStr(i)).Select
        
        Q(i) = Sheets(CStr(i)).Range("c16").value
        daeSoo(i) = Sheets(CStr(i)).Range("c14").value
        
        T1(i) = Sheets(CStr(i)).Range("e7").value
        S1(i) = Sheets(CStr(i)).Range("g7").value
        
        direction(i) = getDirectionFromWell(i)
        gradient(i) = Sheets(CStr(i)).Range("k18").value
        
    Next i


    Sheets("aggWhpa").Select
    Call WriteWellData(Q, daeSoo, T1, S1, direction, gradient, nofwell)
    Call DrawOutline
    
    Range("a1").Select
    Application.CutCopyMode = False
    
End Sub

Sub DrawOutline()

    Application.ScreenUpdating = False
    
    Range("C3:O17").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("B31").Select
    
    Application.ScreenUpdating = True
End Sub


Private Sub WriteWellData(Q As Variant, daeSoo As Variant, T1 As Variant, S1 As Variant, direction As Variant, gradient As Variant, ByVal nofwell As Integer)
    Dim i As Integer
    Dim t_sum As Double
    Dim daesoo_sum As Double
    Dim gradient_sum As Double
    Dim direction_sum As Double
    
    t_sum = 0
    daesoo_sum = 0
    gradient_sum = 0
    direction_sum = 0
    
    
    Application.ScreenUpdating = False
    
    With Range("c4:o17")
        .value = ""
    End With
            
    Call UnmergeAllCells
    
    For i = 1 To nofwell
    
        Cells(3 + i, "c").value = "W-" & CStr(i)
        
        Cells(3 + i, "e").value = Q(i)
        Cells(3 + i, "f").value = T1(i)
        t_sum = t_sum + T1(i)
        
        Cells(3 + i, "i").value = daeSoo(i)
        daesoo_sum = daesoo_sum + daeSoo(i)
        
        Cells(3 + i, "k").value = direction(i)
        direction_sum = direction_sum + direction(i)
        
        Cells(3 + i, "m").value = Format(gradient(i), "###0.0000")
        gradient_sum = gradient_sum + gradient(i)
        
        Cells(4, "d").value = "5년"
    
    Next i
    
   
    Cells(4, "g").value = Round(t_sum / nofwell, 4)
    Cells(4, "g").NumberFormat = "0.0000"
    Call merge_cells("d", nofwell)
    Call merge_cells("g", nofwell)
    
    Cells(4, "j").value = Round(daesoo_sum / nofwell, 1)
    Cells(4, "j").NumberFormat = "0.0"
    Call merge_cells("j", nofwell)
    
    Cells(4, "l").value = Round(direction_sum / nofwell, 1)
    Cells(4, "l").NumberFormat = "0.0"
    Call merge_cells("l", nofwell)
    
    Cells(4, "n").value = Round(gradient_sum / nofwell, 4)
    Cells(4, "n").NumberFormat = "0.0000"
    Call merge_cells("n", nofwell)
    
    Cells(4, "o").value = "무경계조건"
    Call merge_cells("o", nofwell)
    
    Cells(4, "h").value = 0.03
    Call merge_cells("h", nofwell)
    
    Application.ScreenUpdating = True
    
End Sub



Sub merge_cells(cel As String, ByVal nofwell As Integer)

    Range(cel & CStr(4) & ":" & cel & CStr(nofwell + 3)).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    
End Sub

Sub FindMergedCellsRange()
    Dim mergedRange As Range
    Set mergedRange = ActiveSheet.Range("A1").MergeArea
    MsgBox "The range of merged cells is " & mergedRange.Address
End Sub



Sub UnmergeAllCells()
    Dim ws As Worksheet
    Dim cell As Range
    
    Set ws = ActiveSheet
    
    For Each cell In ws.UsedRange
        If cell.MergeCells Then
            cell.MergeCells = False
        End If
    Next cell
    
    Call SetBorderLine_Default
End Sub


Sub SetBorderLine_Default()

    Range("C4:O17").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Sub




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



' Summary Button
Private Sub CommandButton2_Click()
    Dim nofwell As Integer
    
    nofwell = GetNumberOfWell()
    If ActiveSheet.Name <> "AggSum" Then Sheets("AggSum").Select


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
' 안정수위
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
        Cells(65, Chr(Asc("D") + i - 1)).value = Worksheets(CStr(i)).Range("c21").value & unit
    Next i
End Sub



' Write_MotorTochool
' Write_MotorSimdo



Sub Write_MotorPower(nofwell As Integer)
' 모터마력
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


Sub Write_MotorSimdo(nofwell As Integer)
' 모터심도
    Dim i As Integer
    Dim unit As String
    
    Call EraseCellData("d74:p75")

    If Sheets("AggSum").CheckBox1.value = True Then
            unit = " m"
        Else
            unit = ""
    End If


    For i = 1 To nofwell
        Cells(74, Chr(Asc("D") + i - 1)).value = "W-" & CStr(i)
        Cells(75, Chr(Asc("D") + i - 1)).value = Worksheets(CStr(i)).Range("c18").value & unit
    Next i
End Sub


Sub Write_MotorTochool(nofwell As Integer)
' 토출구경
    Dim i As Integer
    Dim unit As String
    
    Call EraseCellData("d69:p70")

    If Sheets("AggSum").CheckBox1.value = True Then
            unit = " mm"
        Else
            unit = ""
    End If


    For i = 1 To nofwell
        Cells(69, Chr(Asc("D") + i - 1)).value = "W-" & CStr(i)
        Cells(70, Chr(Asc("D") + i - 1)).value = Worksheets(CStr(i)).Range("c19").value & unit
    Next i
End Sub



Sub Write_DiggingDepth(nofwell As Integer)
' 굴착심도
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
' 양수영향반경
    Dim i As Integer
    Dim unit As String
    
    Call EraseCellData("d28:g40")
        
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
        Cells(27 + i, "E").value = Worksheets(CStr(i)).Range("H9").value & unit
        Cells(27 + i, "F").value = Worksheets(CStr(i)).Range("K6").value & unit
        Cells(27 + i, "G").value = Worksheets(CStr(i)).Range("K7").value & unit
    Next i
End Sub


Sub Write_DrasticIndex(nofwell As Integer)
' 드라스틱 인덱스
    Dim i As Integer
    
    Call EraseCellData("i28:k40")
    
    For i = 1 To nofwell
        ' WellNum
        Cells(27 + i, "I").value = "W-" & CStr(i)
        Cells(27 + i, "J").value = Worksheets(CStr(i)).Range("k30").value
        Cells(27 + i, "K").value = Worksheets(CStr(i)).Range("k31").value
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
    
    Range("n28").value = CheckDrasticIndex(Range("M28"))
    Range("n29").value = CheckDrasticIndex(Range("M29"))

End Sub

Sub Write26_AquiferCharacterization(nofwell As Integer)
' 대수층 분석및 적정채수량 분석
    Dim i, remainder As Integer
    
    ' Cells(13 + nowfell, "F").Value = Range("well!R5C10").Value
    
    Call EraseCellData("d12:j24")
        
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
    
    ' Cells(13 + nowfell, "F").Value = Range("well!R5C10").Value
    
    Call EraseCellData("L12:S24")
        
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





Option Explicit



Private Sub CommandButton1_Click()
    Sheets("AggStep").Visible = False
    Sheets("Well").Select
End Sub


Private Sub CommandButton2_Click()
    If ActiveSheet.Name <> "AggStep" Then Sheets("AggStep").Select
    Call WriteStepTestData
End Sub


Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub

Private Sub WriteStepTestData()
    Dim fname As String
    Dim nofwell, i As Integer
    
    
    Dim a1() As String
    Dim a2() As String
    Dim a3() As String
    
    Dim Q() As String
    Dim h() As String
    Dim delta_h() As String
    Dim qsw() As String
    Dim swq() As String
    
    nofwell = GetNumberOfWell()
    ' --------------------------------------------------------------------------------------
    ReDim a1(1 To nofwell)
    ReDim a2(1 To nofwell)
    ReDim a3(1 To nofwell)
    
    ReDim Q(1 To nofwell)
    ReDim h(1 To nofwell)
    ReDim delta_h(1 To nofwell)
    ReDim qsw(1 To nofwell)
    ReDim swq(1 To nofwell)
    
    ' --------------------------------------------------------------------------------------
    
    If ActiveSheet.Name <> "AggStep" Then Sheets("AggStep").Select
    
    For i = 1 To nofwell
    
        fname = "A" & CStr(i) & "_ge_OriginalSaveFile.xlsm"
        If Not IsWorkBookOpen(fname) Then
            MsgBox "Please open the yangsoo data ! " & fname
            Exit Sub
        End If
        
        Q(i) = Workbooks(fname).Worksheets("Input").Range("q64").value
        h(i) = Workbooks(fname).Worksheets("Input").Range("r64").value
        delta_h(i) = Workbooks(fname).Worksheets("Input").Range("s64").value
        qsw(i) = Workbooks(fname).Worksheets("Input").Range("t64").value
        swq(i) = Workbooks(fname).Worksheets("Input").Range("u64").value

        a1(i) = Workbooks(fname).Worksheets("Input").Range("v64").value
        a2(i) = Workbooks(fname).Worksheets("Input").Range("w64").value
        a3(i) = Workbooks(fname).Worksheets("Input").Range("x64").value
        
    Next i
    
    Call Write31_StepTestData(a1, a2, a3, Q, h, delta_h, qsw, swq, nofwell)
End Sub


Sub Write31_StepTestData(a1 As Variant, a2 As Variant, a3 As Variant, Q As Variant, h As Variant, delta_h As Variant, qsw As Variant, swq As Variant, nofwell As Variant)

    Dim i As Integer
        
    Call EraseCellData("c5:k19")
    
    For i = 1 To nofwell
        Cells(4 + i, "c").value = "W-" & CStr(i)
        
        Cells(4 + i, "d").value = a1(i)
        Cells(4 + i, "e").value = a2(i)
        Cells(4 + i, "f").value = a3(i)
    
        Cells(4 + i, "g").value = Q(i)
        Cells(4 + i, "h").value = h(i)
        Cells(4 + i, "i").value = delta_h(i)
        Cells(4 + i, "j").value = qsw(i)
        Cells(4 + i, "k").value = swq(i)
    Next i
End Sub


Option Explicit


Private Sub CommandButton1_Click()
    Sheets("AggChart").Visible = False
    Sheets("Well").Select
End Sub


Private Sub CommandButton2_Click()
    If ActiveSheet.Name <> "AggChart" Then Sheets("AggChart").Select
    Call WriteAllCharts
End Sub


Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub





Option Explicit

Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)


Sub DeleteAllCharts()
    Dim ws As Worksheet
    Dim chartObj As ChartObject
    
    Set ws = ThisWorkbook.Worksheets("AggChart")
    
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj
End Sub

Sub DeleteAllImages()
    Dim ws As Worksheet
    Dim sh As Shape
    
    Set ws = ThisWorkbook.Worksheets("AggChart")
    For Each sh In ws.Shapes
        If sh.Type = msoPicture Then
            sh.Delete
        End If
    Next sh
End Sub

Sub WriteAllCharts()
    Dim fname, source_name As String
    Dim nofwell, i As Integer
    
    nofwell = GetNumberOfWell()
    If ActiveSheet.Name <> "AggChart" Then Sheets("AggChart").Select
    
    ' Call DeleteAllCharts
    Call DeleteAllImages
    
    source_name = ActiveWorkbook.Name
    
    For i = 1 To nofwell
        Call Write_InsertChart(i, source_name)
    Next i
    
End Sub

Sub Write_InsertChart(well As Integer, source_name As String)
    Dim fname As String
    Dim imagePath As String

    imagePath = Environ("TEMP") & "\tempChartImage.png"

    fname = "A" & CStr(well) & "_ge_OriginalSaveFile.xlsm"
    If Not IsWorkBookOpen(fname) Then
        MsgBox "Please open the yangsoo data ! " & fname
        Exit Sub
    End If

    Call SaveAndInsertChart(well, source_name, "Chart 5", "d" & CStr(3 + 16 * (well - 1)))
    Call SaveAndInsertChart(well, source_name, "Chart 7", "j" & CStr(3 + 16 * (well - 1)))
    Call SaveAndInsertChart(well, source_name, "Chart 9", "p" & CStr(3 + 16 * (well - 1)))
End Sub


Sub SaveAndInsertChart(well As Integer, source_name As String, chartName As String, targetRange As String)
    Dim imagePath As String
    Dim fname As String
    Dim targetCell As Range
    Dim picWidth As Double, picHeight As Double
    
    imagePath = Environ("TEMP") & "\tempChartImage.png"
    fname = "A" & CStr(well) & "_ge_OriginalSaveFile.xlsm"

    Windows(fname).Activate
    Worksheets("Input").ChartObjects(chartName).Activate
    ActiveChart.Export fileName:=imagePath, FilterName:="PNG"
    
    With ActiveChart.Parent
        picWidth = .Width
        picHeight = .height
    End With
    

    Windows(source_name).Activate
    Set targetCell = Sheets("AggChart").Range(targetRange)
        
    
    Sleep (1000)
    Sheets("AggChart").Shapes.AddPicture _
        fileName:=imagePath, _
        LinkToFile:=False, _
        SaveWithDocument:=True, _
        Left:=targetCell.Left, _
        Top:=targetCell.Top, _
        Width:=picWidth, _
        height:=picHeight
        
End Sub


Sub ActivateChart()
    Dim ws As Worksheet
    Dim chartObj As ChartObject

    Set ws = ThisWorkbook.Worksheets("Input")
    Set chartObj = ws.ChartObjects("Chart 5")
    chartObj.Activate
End Sub




Private Sub CommandButton1_Click()
    Sheets("YangSoo").Visible = False
    Sheets("Well").Select
End Sub

Private Sub CommandButton2_Click()
    Call GetBaseDataFromYangSoo
End Sub

Private Sub EraseCellData(str_range As String)
    With Range(str_range)
        .value = ""
    End With
End Sub


Sub GetBaseDataFromYangSoo()
    Dim fname As String
    Dim nofwell, i As Integer
        
    Dim natural() As Double ' 자연수위, natural depth
    Dim stable() As Double  ' 안정수위, stable depth
    Dim recover() As Double ' 회복수위, recover depth
    Dim Sw() As Double ' 수위회복량 - 안정수위 - 회복수위
    
    Dim delta_h() As Double ' deltah : 수위강하량
    
    Dim radius() As Double ' 공반경
    Dim Rw() As Double      ' 공반경 / 2000
    
    Dim well_depth() As Double     ' 관정심도, well depth
    Dim casing() As Double  ' 케이싱심도
    
    Dim Q() As Double       '취수계획량
    Dim delta_s() As Double
    Dim hp() As Double
    
    Dim daeSoo() As Double  ' 대수층 두께
    
    Dim T1() As Double      ' T1
    Dim T2() As Double      ' T2
    Dim TA() As Double      ' TA - (T1+T2)/2, TAverage
    
    Dim S1() As Double      ' S1
    Dim S2() As Double      ' S2 - 스킨팩터 해석, s값
    
    Dim T0() As Double         ' 단계양수시험의 T값
    Dim S0() As Double         ' S값, 0.005: 피압대수층, 0.001: 누수대수층, 0.1: 자유면대수층
    Dim ER_MODE() As String    ' 영향반경산정공식 선정
    Dim ER1() As Double
    Dim ER2() As Double
    Dim ER3() As Double
    
    Dim K() As Double
    Dim time_() As Double   ' 안정수위도달시간
    
    Dim shultze() As Double
    Dim webber() As Double
    Dim jacob() As Double
    
    Dim skin() As Double ' skin factor
    Dim er() As Double   ' effective radius, 유효우물반경
        
    
    Dim qh() As Double ' 한계양수량
    Dim qg() As Double ' 가채수량
    Dim q1() As Double ' Q1
    
    
    Dim sd1() As Double ' 1단계 수위강하량
    Dim sd2() As Double ' 4단계 수위강하량
    
    
    Dim C() As Double
    Dim B() As Double
    
    Dim ratio() As Double
    
    
 ' --------------------------------------------------------------------------------------

    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
    
 ' --------------------------------------------------------------------------------------
    
    ReDim natural(1 To nofwell)
    ReDim stable(1 To nofwell)
    ReDim recover(1 To nofwell)
    ReDim delta_h(1 To nofwell)
    ReDim Sw(1 To nofwell)
    
    
    ReDim radius(1 To nofwell)
    ReDim Rw(1 To nofwell)
    
    ReDim well_depth(1 To nofwell)
    ReDim casing(1 To nofwell)
    
    ReDim Q(1 To nofwell)
    ReDim delta_s(1 To nofwell)
    ReDim hp(1 To nofwell)
    
    ReDim daeSoo(1 To nofwell)
    
    ReDim T1(1 To nofwell)
    ReDim T2(1 To nofwell)
    ReDim TA(1 To nofwell)
    
    ReDim S1(1 To nofwell)
    ReDim S2(1 To nofwell)
    
    ReDim K(1 To nofwell)
    ReDim time_(1 To nofwell)
    
    ReDim shultze(1 To nofwell)
    ReDim webber(1 To nofwell)
    ReDim jacob(1 To nofwell)
    
    ReDim skin(1 To nofwell)
    ReDim er(1 To nofwell)
        
    ReDim ER1(1 To nofwell)
    ReDim ER2(1 To nofwell)
    ReDim ER3(1 To nofwell)
    
    ReDim qh(1 To nofwell)
    ReDim qg(1 To nofwell)
    
    
    ReDim sd1(1 To nofwell)
    ReDim sd2(1 To nofwell)
    ReDim q1(1 To nofwell)
    
    ReDim C(1 To nofwell)
    ReDim B(1 To nofwell)
    
    ReDim ratio(1 To nofwell)
    
    
    ReDim T0(1 To nofwell)         ' 단계양수시험의 T값
    ReDim S0(1 To nofwell)         ' S값, 0.005: 피압대수층, 0.001: 누수대수층, 0.1: 자유면대수층
    ReDim ER_MODE(1 To nofwell)
    
    Call EraseCellData("a5:an19")
    
    For i = 1 To nofwell
    
        fname = "A" & CStr(i) & "_ge_OriginalSaveFile.xlsm"
        If Not IsWorkBookOpen(fname) Then
            MsgBox "Please open the yangsoo data ! " & fname
            Exit Sub
        End If
        
        
        Q(i) = Workbooks(fname).Worksheets("Input").Range("m51").value
        hp(i) = Workbooks(fname).Worksheets("Input").Range("i48").value
        
        natural(i) = Workbooks(fname).Worksheets("Input").Range("m48").value
        stable(i) = Workbooks(fname).Worksheets("Input").Range("m49").value
        radius(i) = Workbooks(fname).Worksheets("Input").Range("m44").value
        Rw(i) = radius(i) / 2000
        
        well_depth(i) = Workbooks(fname).Worksheets("Input").Range("m45").value
        casing(i) = Workbooks(fname).Worksheets("Input").Range("i52").value
        
        
        C(i) = Workbooks(fname).Worksheets("Input").Range("A31").value
        B(i) = Workbooks(fname).Worksheets("Input").Range("B31").value
        
        
        
        recover(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c10").value
        Sw(i) = stable(i) - recover(i)
        
        delta_h(i) = Workbooks(fname).Worksheets("SkinFactor").Range("b16").value
        delta_s(i) = Workbooks(fname).Worksheets("SkinFactor").Range("b4").value
        
        daeSoo(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c16").value
        
        '----------------------------------------------------------------------------------
        
        T0(i) = Workbooks(fname).Worksheets("SkinFactor").Range("d4").value
        S0(i) = Workbooks(fname).Worksheets("SkinFactor").Range("f4").value
        ER_MODE(i) = Workbooks(fname).Worksheets("SkinFactor").Range("h10").value
        
        T1(i) = Workbooks(fname).Worksheets("SkinFactor").Range("d5").value
        T2(i) = Workbooks(fname).Worksheets("SkinFactor").Range("h13").value
        TA(i) = (T1(i) + T2(i)) / 2
        
        S1(i) = Workbooks(fname).Worksheets("SkinFactor").Range("e10").value
        S2(i) = Workbooks(fname).Worksheets("SkinFactor").Range("i16").value
        
        K(i) = Workbooks(fname).Worksheets("SkinFactor").Range("e16").value
        time_(i) = Workbooks(fname).Worksheets("SkinFactor").Range("h16").value
        
        shultze(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c13").value
        webber(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c18").value
        jacob(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c23").value
        
        skin(i) = Workbooks(fname).Worksheets("SkinFactor").Range("g6").value
        er(i) = Workbooks(fname).Worksheets("SkinFactor").Range("c8").value
        
        
        ' 경험식, 1번, 2번, 3번의 유효우물반경
        ER1(i) = Workbooks(fname).Worksheets("SkinFactor").Range("K8").value
        ER2(i) = Workbooks(fname).Worksheets("SkinFactor").Range("K9").value
        ER3(i) = Workbooks(fname).Worksheets("SkinFactor").Range("K10").value
        
        '----------------------------------------------------------------------------------
        
        qh(i) = Workbooks(fname).Worksheets("SafeYield").Range("b13").value
        qg(i) = Workbooks(fname).Worksheets("SafeYield").Range("b7").value
        
        sd1(i) = Workbooks(fname).Worksheets("SafeYield").Range("b3").value
        sd2(i) = Workbooks(fname).Worksheets("SafeYield").Range("b4").value
        q1(i) = Workbooks(fname).Worksheets("SafeYield").Range("b2").value
        
        ratio(i) = Workbooks(fname).Worksheets("SafeYield").Range("b11").value
        
        '*****************************************************************************************
        
        Cells(4 + i, "a").value = "W-" & CStr(i)
        Cells(4 + i, "b").value = natural(i)
        Cells(4 + i, "c").value = stable(i)
        
        Cells(4 + i, "d").value = recover(i)
        Cells(4 + i, "d").NumberFormat = "0.00"
        
        Cells(4 + i, "e").value = Sw(i)
        Cells(4 + i, "e").NumberFormat = "0.00"
        
        Cells(4 + i, "f").value = delta_h(i)
        Cells(4 + i, "f").NumberFormat = "0.00"
        
        Cells(4 + i, "g").value = radius(i)
        Cells(4 + i, "h").value = Rw(i)
        Cells(4 + i, "i").value = well_depth(i)
        Cells(4 + i, "j").value = casing(i)
        Cells(4 + i, "k").value = Q(i)
        
        Cells(4 + i, "l").value = delta_s(i)
        Cells(4 + i, "l").NumberFormat = "0.00"
        
        Cells(4 + i, "m").value = hp(i)
        Cells(4 + i, "n").value = daeSoo(i)
        
        Cells(4 + i, "o").value = T1(i)
        Cells(4 + i, "o").NumberFormat = "0.0000"
         
        Cells(4 + i, "p").value = T2(i)
        Cells(4 + i, "p").NumberFormat = "0.0000"
         
        Cells(4 + i, "q").value = TA(i)
        Cells(4 + i, "q").NumberFormat = "0.0000"
        
        Cells(4 + i, "r").value = S1(i)
        
        Cells(4 + i, "s").value = S2(i)
        Cells(4 + i, "s").NumberFormat = "0.0000000"
        
        Cells(4 + i, "t").value = K(i)
        Cells(4 + i, "t").NumberFormat = "0.0000"
        
        Cells(4 + i, "u").value = time_(i)
        
        Cells(4 + i, "v").value = shultze(i)
        Cells(4 + i, "v").NumberFormat = "0.0"
        
        Cells(4 + i, "w").value = webber(i)
        Cells(4 + i, "w").NumberFormat = "0.0"
        
        Cells(4 + i, "x").value = jacob(i)
        Cells(4 + i, "x").NumberFormat = "0.0"
        
        
        
        Cells(4 + i, "y").value = Format(skin(i), "0.0000")
        
        Cells(4 + i, "z").value = er(i)
        Cells(4 + i, "z").NumberFormat = "0.0000"
        
        Cells(4 + i, "aa").value = Format(qh(i), "0.")
        Cells(4 + i, "ab").value = Format(qg(i), "0.00")
        Cells(4 + i, "ac").value = Format(q1(i), "0.")
        
        Cells(4 + i, "ad").value = Format(sd1(i), "0.00")
        Cells(4 + i, "ae").value = Format(sd2(i), "0.00")
        
        Cells(4 + i, "af").value = C(i)
        Cells(4 + i, "ag").value = B(i)
        
        Cells(4 + i, "ah").value = ratio(i)
        Cells(4 + i, "ah").NumberFormat = "0.0%"
        
        
        ' 2023/09/22 새로 추가한 값들 ...
        Cells(4 + i, "AI").value = Format(T0(i), "0.0000")
        Cells(4 + i, "AJ").value = Format(S0(i), "0.0000")
        Cells(4 + i, "AK").value = ER_MODE(i)
        
        Cells(4 + i, "AL").value = Format(ER1(i), "0.0000")
        Cells(4 + i, "AM").value = Format(ER2(i), "0.0000")
        Cells(4 + i, "AN").value = Format(ER3(i), "0.0000")
        
    Next i
End Sub


Public Sub MyDebug(sPrintStr As String, Optional bClear As Boolean = False)
   If bClear = True Then
      Application.SendKeys "^g^{END}", True

      DoEvents '  !!! DoEvents is VERY IMPORTANT here !!!

      Debug.Print String(30, vbCrLf)
   End If

   Debug.Print sPrintStr
End Sub


'0 : skin factor, cell, C8
'1 : Re1,         cell, E8
'2 : Re2,         cell, H8
'3 : Re3,         cell, G10

Function DetermineEffectiveRadius(ERMode As String) As Integer
    Dim er, r As String
    
    er = ERMode
    'MsgBox er
    r = Mid(er, 5, 1)
    
    If r = "F" Then
        DetermineEffectiveRadius = erRE0
    Else
        DetermineEffectiveRadius = val(r)
    End If
End Function



Private Sub CommandButton3_Click()
' Formula Button

Dim formula1, formula2 As String
Dim nofwell As String
Dim i As Integer
Dim T, Q, radius, skin, er As Double
Dim T0, S0 As Double
Dim ERMode As String
Dim ER1, ER2, ER3, B, S1 As Double



nofwell = GetNumberOfWell()
Sheets("YangSoo").Select
    
    
Call MyDebug("Formula SkinFactor ... ", True)

Debug.Print "************************************************************************************************************************************************************************************************"

For i = 1 To nofwell

    T = Format(Cells(4 + i, "o").value, "0.0000")
    Q = Cells(4 + i, "k").value
    
    T0 = Format(Cells(4 + i, "AI").value, "0.0000")
    S0 = Format(Cells(4 + i, "AJ").value, "0.0000")
    S1 = Cells(4 + i, "R").value
            
    delta_s = Format(Cells(4 + i, "l").value, "0.00")
    radius = Format(Cells(4 + i, "h").value, "0.000")
    skin = Cells(4 + i, "y").value
    er = Cells(4 + i, "z").value
    
    
    B = Format(Cells(4 + i, "AG").value, "0.0000")
    ER1 = Cells(4 + i, "AL").value
    ER2 = Cells(4 + i, "AM").value
    ER3 = Cells(4 + i, "AN").value
    
    
    Select Case DetermineEffectiveRadius(Cells(4 + i, "AK").value)
    ' 경험식 1번
    Case erRE1
        formula1 = "W-" & i & "호공~~r _{e-" & i & "} `=~ sqrt {{2.25 TIMES  " & T0 & " TIMES  0.0833333} over {" & S0 & " TIMES  10 ^{5.46 TIMES  " & T & " TIMES  " & B & "}}} `=~" & ER1 & "m"
        formula2 = "erRE1, 경험식 1번"
        
    ' 경험식 2번
    Case erRE2
        formula1 = "W-" & i & "호공~~r _{e-" & i & "} `=~ sqrt {{2.25 TIMES  " & T0 & " TIMES  0.0833333} over {" & S0 & " TIMES  10 ^{4 pi TIMES " & T & " TIMES  " & B & "}}} `=~" & ER2 & "m"
        formula2 = "erRE2, 경험식 2번"
    ' 경험식 3번
    Case erRE3
        formula1 = "W-" & i & "호공~~r _{e-" & i & "} `=~" & radius & " TIMES  sqrt {{" & S1 & "} over {" & S0 & "}} `=~" & ER3 & "m"
        formula2 = "erRE3, 경험식 3번"
        
    Case Else
        ' 스킨계수
        formula1 = "W-" & i & "호공~~ sigma  _{w-" & i & "} = {2 pi  TIMES  " & T & " TIMES  " & delta_s & " } over {" & Q & "} -1.15 TIMES  log {2.25 TIMES  " & T & " TIMES  (1/1440)} over {0.0005 TIMES  (" & radius & " TIMES  " & radius & ")} =`" & skin
        ' 유효우물반경
        formula2 = "W-" & i & "호공~~r _{e-" & i & "} `=~r _{w} e ^{- sigma  _{w-" & i & "}} =" & radius & " TIMES e ^{-(" & skin & ")} =" & er & "m"
    End Select
    
    Debug.Print formula1
    Debug.Print formula2
    Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"

Next i

Debug.Print "************************************************************************************************************************************************************************************************"

Call FormulaChwiSoo
' 3-7, 적정취수량

Call FormulaRadiusOfInfluence
' 양수영향반경
End Sub


Sub FormulaChwiSoo()
' 3-7, 적정취수량

    Dim formula As String
    Dim nofwell As String
    Dim i As Integer
    Dim q1, S1, S2, res As Double
    
    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
        
        
    For i = 1 To 3
        Debug.Print " "
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    
    For i = 1 To nofwell
        q1 = Cells(4 + i, "ac").value
 
        S1 = Format(Cells(4 + i, "ad").value, "0.00")
        S2 = Format(Cells(4 + i, "ae").value, "0.00")
        
        res = Cells(4 + i, "ab").value
        formula = "W-" & i & "호공~~Q _{ & 2} `＝" & q1 & "` TIMES  `(` {" & S2 & "} over {" & S1 & "} `) ^{2/3} `＝" & res & "㎥/일"
        
        Debug.Print formula
        Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
End Sub


Sub FormulaRadiusOfInfluence()
' 양수영향반경

    Dim formula1, formula2, formula3 As String
    ' 슐츠, 웨버, 제이콥
    
    Dim nofwell As String
    Dim i As Integer
    Dim shultze, webber, jacob, T, K, S, time_, delta_h As String
    
    nofwell = GetNumberOfWell()
    Sheets("YangSoo").Select
        
        
    For i = 1 To 3
        Debug.Print " "
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    
    For i = 1 To nofwell
        shultze = CStr(Cells(4 + i, "v").value)
        webber = CStr(Cells(4 + i, "w").value)
        jacob = CStr(Cells(4 + i, "x").value)
        
        T = CStr(Format(Cells(4 + i, "q").value, "0.0000"))
        S = CStr(Format(Cells(4 + i, "s").value, "0.0000000"))
        K = CStr(Format(Cells(4 + i, "t").value, "0.0000"))
    
        delta_h = CStr(Cells(4 + i, "f").value)
        time_ = CStr(Format(Cells(4 + i, "u").value, "0.0000"))
        
        
        ' Cells(4 + i, "y").value = Format(skin(i), "0.0000")
        
        formula1 = "W-" & i & "호공~~R _{W-" & i & "} ``=`` sqrt {6 TIMES  " & delta_h & " TIMES  " & K & " TIMES  " & time_ & "/" & S & "} ``=~" & shultze & "m"
        formula2 = "W-" & i & "호공~~R _{W-" & i & "} ``=3`` sqrt {" & delta_h & " TIMES " & K & " TIMES " & time_ & "/" & S & "} `=`" & webber & "`m"
        formula3 = "W-" & i & "호공~~r _{0(W-" & i & ")} `=~ sqrt {{2.25 TIMES  " & T & " TIMES  " & time_ & "} over {" & S & "}} `=~" & jacob & "m"
        
        Debug.Print ""
        Debug.Print formula1
        Debug.Print formula2
        Debug.Print formula3
        Debug.Print ""
        
        Debug.Print "-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------"
    Next i
    
    Debug.Print "************************************************************************************************************************************************************************************************"
    
End Sub







Private Sub CommandButton1_Click()
   Sheets("water").Visible = False
    Sheets("Well").Select
End Sub

Private Sub CommandButton2_Click()
    Dim WB_NAME, cpRange  As String

    If Workbooks.count = 1 Then
        MsgBox "Please Open 기사용관정, 파일 ", vbOKOnly
        Exit Sub
    End If
   
    
    WB_NAME = GetOtherFileName
    
    cpRange = GetCopyPoint(WB_NAME)
    Call CopyFromGWAN_JUNG(WB_NAME, cpRange)
    Call FormulaInjection
    
End Sub

' 2024-01-14
' inject formula ...

Private Sub FormulaInjection()
    Dim nofwell, i As Integer
    
    nofwell = GetNumberOfWell()
    For i = 4 To nofwell + 3
        Sheets("Well").Cells(i, "O").formula = "=ROUND(water!$F$7, 1)"
    Next i

End Sub

'
'Function lastRowByKey(cell As String) As Long
'    lastRowByKey = Range(cell).End(xlDown).Row
'End Function


Function GetCopyPoint(ByVal fname As String) As String

  Dim ip1, ip2 As Integer

  ip1 = Workbooks(fname).Worksheets("ss").Range("b1").End(xlDown).Row + 4
  ip2 = ip1 + 2
  
  GetCopyPoint = "B" & ip1 & ":J" & ip2
  ThisWorkbook.Activate

End Function


Sub CopyFromGWAN_JUNG(ByVal fname As String, ByVal cpRange As String)

    Workbooks(fname).Worksheets("ss").Activate
    Workbooks(fname).Worksheets("ss").Range(cpRange).Select
    Selection.Copy
    
    ThisWorkbook.Sheets("water").Activate
    
    Range("d6").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

End Sub

Sub ListOpenWorkbookNames()
    Dim Workbook As Workbook
    Dim workbookNames As String
    Dim i As Long
        
    workbookNames = ""
    
    For Each Workbook In Application.Workbooks
        workbookNames = workbookNames & Workbook.Name & vbCrLf
    Next
    
    Cells(1, 1).value = workbookNames
End Sub

Function GetOtherFileName() As String
    Dim Workbook As Workbook
    Dim workbookNames As String
    Dim i As Long

    workbookNames = ""
    
    For Each Workbook In Application.Workbooks
        If StrComp(ThisWorkbook.Name, Workbook.Name, vbTextCompare) = 0 Then
            GoTo NEXT_ITERATION
        End If
        
        If CheckSubstring(Workbook.Name, "관정") Then
            Exit For
        End If
        
NEXT_ITERATION:
    Next
    
    GetOtherFileName = Workbook.Name
End Function


Function CheckSubstring(str As String, chk As String) As Boolean
    
    If InStr(str, chk) > 0 Then
        ' The string contains "chk"
        CheckSubstring = True
    Else
        ' The string does not contain "chk"
        CheckSubstring = False
    End If
End Function

Sub DumpRangeToArrayAndSaveTest()
    Dim myArray() As Variant
    Dim rng As Range
    Dim i As Integer, j As Integer
    
    ' Set the range you want to dump to an array
    Set rng = ThisWorkbook.ActiveSheet.Range("B6:N35")
    
    ' Read the range into an array
    myArray = rng.value
    
    ' Save array to a file
    Dim filePath As String
    filePath = ThisWorkbook.Path & "\" & ActiveSheet.Name & ".csv"
    SaveArrayToFileByExcelForm myArray, filePath
    
End Sub



Sub SaveArrayToFileByExcelForm(myArray As Variant, filePath As String)
    Dim i As Integer, j As Integer
    Dim FileNum As Integer
    
    FileNum = FreeFile
    
    Open filePath For Output As FileNum
    
    For i = LBound(myArray, 1) To UBound(myArray, 1)
        For j = LBound(myArray, 2) To UBound(myArray, 2)
            Print #FileNum, "myArray(" & i & ", " & j & ") = ";
            
            ' Separate values with a comma (CSV format)
            If j <= UBound(myArray, 2) Then
                Print #FileNum, myArray(i, j);
            End If
            
            Print #FileNum, ""
        Next j
        ' Start a new line for each row
        Print #FileNum, ""
    Next i
    
    Close FileNum
End Sub


Sub importFromArray()
    Dim myArray As Variant
    Dim rng As Range
    
    indexString = "data_" & UCase(Range("s11").value)
    
    myArray = Application.Run(indexString)
    
    
    Set rng = ThisWorkbook.ActiveSheet.Range("B6:N35")
    rng.value = myArray
       
End Sub





'***********************
'Year 2024
'***********************
'data_GEUMSAN
'data_BORYUNG
'data_DAEJEON
'data_BUYEO
'data_SEOSAN
'data_CHEONAN
'data_CHEUNGJU
'***********************
'data_HONGSUNG
'data_SEJONG
'***********************


'
'Function data_CHEUNGJU() As Variant
'    Dim myArray() As Variant
'    ReDim myArray(1 To 30, 1 To 13)
'
'
'    data_CHEUNGJU = myArray
'End Function
'
'Function data_GEUMSAN() As Variant
'    Dim myArray() As Variant
'    ReDim myArray(1 To 30, 1 To 13)
'
'    data_GEUMSAN = myArray
'End Function
'
'
'
'Function data_DAEJEON() As Variant
'    Dim myArray() As Variant
'    ReDim myArray(1 To 30, 1 To 13)
'
'
'    data_DAEJEON = myArray
'End Function
'
'
'Function data_BUYEO() As Variant
'    Dim myArray() As Variant
'    ReDim myArray(1 To 30, 1 To 13)
'
'
'    data_BUYEO = myArray
'End Function
'
'
'Function data_SEOSAN() As Variant
'    Dim myArray() As Variant
'    ReDim myArray(1 To 30, 1 To 13)
'
'
'    data_SEOSAN = myArray
'End Function
'
'
'Function data_CHEONAN() As Variant
'    Dim myArray() As Variant
'    ReDim myArray(1 To 30, 1 To 13)
'
'
'
'    data_CHEONAN = myArray
'End Function
'
'
'Function data_BORYUNG() As Variant
'    Dim myArray() As Variant
'    ReDim myArray(1 To 30, 1 To 13)
'
'
'    data_BORYUNG = myArray
'End Function
'



Function data_TEMP() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    data_TEMP = myArray

End Function



Function data_HONGSUNG() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)
    
    data_HONGSUNG = myArray

End Function

Function data_BORYUNG() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

    myArray(1, 1) = 1994
    myArray(1, 2) = 17.9
    myArray(1, 3) = 6
    myArray(1, 4) = 58.2
    myArray(1, 5) = 51.5
    myArray(1, 6) = 135.5
    myArray(1, 7) = 207
    myArray(1, 8) = 137
    myArray(1, 9) = 443.5
    myArray(1, 10) = 21
    myArray(1, 11) = 155
    myArray(1, 12) = 17.5
    myArray(1, 13) = 18.9

    myArray(2, 1) = 1995
    myArray(2, 2) = 15.7
    myArray(2, 3) = 11
    myArray(2, 4) = 19.6
    myArray(2, 5) = 65.5
    myArray(2, 6) = 49.5
    myArray(2, 7) = 26.5
    myArray(2, 8) = 144.5
    myArray(2, 9) = 996.5
    myArray(2, 10) = 70.5
    myArray(2, 11) = 24.5
    myArray(2, 12) = 23
    myArray(2, 13) = 12.7

    myArray(3, 1) = 1996
    myArray(3, 2) = 33.4
    myArray(3, 3) = 6.8
    myArray(3, 4) = 104.5
    myArray(3, 5) = 34
    myArray(3, 6) = 22.5
    myArray(3, 7) = 235
    myArray(3, 8) = 192.5
    myArray(3, 9) = 44.5
    myArray(3, 10) = 14
    myArray(3, 11) = 106.5
    myArray(3, 12) = 74.2
    myArray(3, 13) = 31.7

    myArray(4, 1) = 1997
    myArray(4, 2) = 15.1
    myArray(4, 3) = 38.4
    myArray(4, 4) = 30.5
    myArray(4, 5) = 57.5
    myArray(4, 6) = 203
    myArray(4, 7) = 272
    myArray(4, 8) = 353.5
    myArray(4, 9) = 211.5
    myArray(4, 10) = 23
    myArray(4, 11) = 10
    myArray(4, 12) = 193.5
    myArray(4, 13) = 34.3

    myArray(5, 1) = 1998
    myArray(5, 2) = 29.9
    myArray(5, 3) = 40.2
    myArray(5, 4) = 30.5
    myArray(5, 5) = 138
    myArray(5, 6) = 100
    myArray(5, 7) = 209.5
    myArray(5, 8) = 263
    myArray(5, 9) = 341.7
    myArray(5, 10) = 150.3
    myArray(5, 11) = 61
    myArray(5, 12) = 29.3
    myArray(5, 13) = 3.8

    myArray(6, 1) = 1999
    myArray(6, 2) = 7.9
    myArray(6, 3) = 9.5
    myArray(6, 4) = 71
    myArray(6, 5) = 88.5
    myArray(6, 6) = 124.5
    myArray(6, 7) = 192.5
    myArray(6, 8) = 98
    myArray(6, 9) = 180
    myArray(6, 10) = 292.5
    myArray(6, 11) = 169
    myArray(6, 12) = 24.9
    myArray(6, 13) = 25.8

    myArray(7, 1) = 2000
    myArray(7, 2) = 42.1
    myArray(7, 3) = 3.2
    myArray(7, 4) = 7
    myArray(7, 5) = 35
    myArray(7, 6) = 53.5
    myArray(7, 7) = 159.5
    myArray(7, 8) = 155
    myArray(7, 9) = 701.5
    myArray(7, 10) = 241
    myArray(7, 11) = 46
    myArray(7, 12) = 39.5
    myArray(7, 13) = 32.1

    myArray(8, 1) = 2001
    myArray(8, 2) = 73.3
    myArray(8, 3) = 46
    myArray(8, 4) = 15.9
    myArray(8, 5) = 26
    myArray(8, 6) = 17
    myArray(8, 7) = 129
    myArray(8, 8) = 286.5
    myArray(8, 9) = 170
    myArray(8, 10) = 10
    myArray(8, 11) = 85
    myArray(8, 12) = 13
    myArray(8, 13) = 32

    myArray(9, 1) = 2002
    myArray(9, 2) = 50.8
    myArray(9, 3) = 5.5
    myArray(9, 4) = 32
    myArray(9, 5) = 169
    myArray(9, 6) = 155.5
    myArray(9, 7) = 72
    myArray(9, 8) = 217.5
    myArray(9, 9) = 477
    myArray(9, 10) = 27
    myArray(9, 11) = 134
    myArray(9, 12) = 61.1
    myArray(9, 13) = 51.8

    myArray(10, 1) = 2003
    myArray(10, 2) = 30.7
    myArray(10, 3) = 44.5
    myArray(10, 4) = 39.5
    myArray(10, 5) = 168.5
    myArray(10, 6) = 78.5
    myArray(10, 7) = 153
    myArray(10, 8) = 309.5
    myArray(10, 9) = 310
    myArray(10, 10) = 128
    myArray(10, 11) = 23
    myArray(10, 12) = 45.5
    myArray(10, 13) = 13

    myArray(11, 1) = 2004
    myArray(11, 2) = 22.1
    myArray(11, 3) = 28.5
    myArray(11, 4) = 45.7
    myArray(11, 5) = 58
    myArray(11, 6) = 105.5
    myArray(11, 7) = 234.5
    myArray(11, 8) = 263.5
    myArray(11, 9) = 164
    myArray(11, 10) = 195
    myArray(11, 11) = 4
    myArray(11, 12) = 56.5
    myArray(11, 13) = 38.9

    myArray(12, 1) = 2005
    myArray(12, 2) = 5.8
    myArray(12, 3) = 35.8
    myArray(12, 4) = 30
    myArray(12, 5) = 73.5
    myArray(12, 6) = 48.5
    myArray(12, 7) = 156
    myArray(12, 8) = 260.5
    myArray(12, 9) = 291.5
    myArray(12, 10) = 282.5
    myArray(12, 11) = 21
    myArray(12, 12) = 18
    myArray(12, 13) = 43.4

    myArray(13, 1) = 2006
    myArray(13, 2) = 27
    myArray(13, 3) = 25.9
    myArray(13, 4) = 10.6
    myArray(13, 5) = 81.5
    myArray(13, 6) = 94.5
    myArray(13, 7) = 114.5
    myArray(13, 8) = 321
    myArray(13, 9) = 21.5
    myArray(13, 10) = 23.5
    myArray(13, 11) = 24.5
    myArray(13, 12) = 61.5
    myArray(13, 13) = 25.4

    myArray(14, 1) = 2007
    myArray(14, 2) = 23.4
    myArray(14, 3) = 29.8
    myArray(14, 4) = 102
    myArray(14, 5) = 29.5
    myArray(14, 6) = 79
    myArray(14, 7) = 85
    myArray(14, 8) = 214
    myArray(14, 9) = 239.5
    myArray(14, 10) = 384
    myArray(14, 11) = 59
    myArray(14, 12) = 17.5
    myArray(14, 13) = 33.1

    myArray(15, 1) = 2008
    myArray(15, 2) = 20.9
    myArray(15, 3) = 10.8
    myArray(15, 4) = 48.2
    myArray(15, 5) = 40.5
    myArray(15, 6) = 78.9
    myArray(15, 7) = 101.3
    myArray(15, 8) = 257.2
    myArray(15, 9) = 119.5
    myArray(15, 10) = 46.9
    myArray(15, 11) = 26.7
    myArray(15, 12) = 37.6
    myArray(15, 13) = 25

    myArray(16, 1) = 2009
    myArray(16, 2) = 18.5
    myArray(16, 3) = 23.3
    myArray(16, 4) = 55.1
    myArray(16, 5) = 41.5
    myArray(16, 6) = 154.5
    myArray(16, 7) = 115.1
    myArray(16, 8) = 320.9
    myArray(16, 9) = 176.6
    myArray(16, 10) = 25.5
    myArray(16, 11) = 39.5
    myArray(16, 12) = 52.9
    myArray(16, 13) = 58

    myArray(17, 1) = 2010
    myArray(17, 2) = 30.1
    myArray(17, 3) = 73.5
    myArray(17, 4) = 75.9
    myArray(17, 5) = 58.5
    myArray(17, 6) = 122.8
    myArray(17, 7) = 60.8
    myArray(17, 8) = 396.5
    myArray(17, 9) = 402.7
    myArray(17, 10) = 213.1
    myArray(17, 11) = 19.2
    myArray(17, 12) = 16.3
    myArray(17, 13) = 32.9

    myArray(18, 1) = 2011
    myArray(18, 2) = 11.1
    myArray(18, 3) = 37.5
    myArray(18, 4) = 18
    myArray(18, 5) = 72.1
    myArray(18, 6) = 115.3
    myArray(18, 7) = 318
    myArray(18, 8) = 723.1
    myArray(18, 9) = 289.4
    myArray(18, 10) = 70.8
    myArray(18, 11) = 13.9
    myArray(18, 12) = 61.3
    myArray(18, 13) = 12.5

    myArray(19, 1) = 2012
    myArray(19, 2) = 24.2
    myArray(19, 3) = 9.2
    myArray(19, 4) = 45
    myArray(19, 5) = 68.9
    myArray(19, 6) = 14.6
    myArray(19, 7) = 76.8
    myArray(19, 8) = 231.1
    myArray(19, 9) = 450.2
    myArray(19, 10) = 207.7
    myArray(19, 11) = 65
    myArray(19, 12) = 61.1
    myArray(19, 13) = 65.2

    myArray(20, 1) = 2013
    myArray(20, 2) = 28.4
    myArray(20, 3) = 40.7
    myArray(20, 4) = 53.4
    myArray(20, 5) = 68.2
    myArray(20, 6) = 116.6
    myArray(20, 7) = 159.9
    myArray(20, 8) = 267.5
    myArray(20, 9) = 214.6
    myArray(20, 10) = 320
    myArray(20, 11) = 10.9
    myArray(20, 12) = 81.1
    myArray(20, 13) = 26.4

    myArray(21, 1) = 2014
    myArray(21, 2) = 3.4
    myArray(21, 3) = 20.5
    myArray(21, 4) = 56.3
    myArray(21, 5) = 70
    myArray(21, 6) = 47.1
    myArray(21, 7) = 125.8
    myArray(21, 8) = 104
    myArray(21, 9) = 168.5
    myArray(21, 10) = 152
    myArray(21, 11) = 156
    myArray(21, 12) = 39.9
    myArray(21, 13) = 66.6

    myArray(22, 1) = 2015
    myArray(22, 2) = 29.9
    myArray(22, 3) = 23.4
    myArray(22, 4) = 30.9
    myArray(22, 5) = 129.7
    myArray(22, 6) = 38.8
    myArray(22, 7) = 83.9
    myArray(22, 8) = 94.7
    myArray(22, 9) = 30.2
    myArray(22, 10) = 13.3
    myArray(22, 11) = 90
    myArray(22, 12) = 155.6
    myArray(22, 13) = 65

    myArray(23, 1) = 2016
    myArray(23, 2) = 7.8
    myArray(23, 3) = 54.2
    myArray(23, 4) = 18.7
    myArray(23, 5) = 105.1
    myArray(23, 6) = 146.5
    myArray(23, 7) = 23.7
    myArray(23, 8) = 200.2
    myArray(23, 9) = 5.1
    myArray(23, 10) = 73.4
    myArray(23, 11) = 108
    myArray(23, 12) = 5.6
    myArray(23, 13) = 44.5

    myArray(24, 1) = 2017
    myArray(24, 2) = 14.8
    myArray(24, 3) = 30.2
    myArray(24, 4) = 14.4
    myArray(24, 5) = 57.6
    myArray(24, 6) = 58.9
    myArray(24, 7) = 21.1
    myArray(24, 8) = 278.1
    myArray(24, 9) = 210
    myArray(24, 10) = 90
    myArray(24, 11) = 26.6
    myArray(24, 12) = 15.9
    myArray(24, 13) = 38.6

    myArray(25, 1) = 2018
    myArray(25, 2) = 15
    myArray(25, 3) = 33.6
    myArray(25, 4) = 92
    myArray(25, 5) = 128.1
    myArray(25, 6) = 104.5
    myArray(25, 7) = 71
    myArray(25, 8) = 262.7
    myArray(25, 9) = 239.6
    myArray(25, 10) = 158.2
    myArray(25, 11) = 154.7
    myArray(25, 12) = 46.7
    myArray(25, 13) = 31.1

    myArray(26, 1) = 2019
    myArray(26, 2) = 1.9
    myArray(26, 3) = 17.8
    myArray(26, 4) = 18.2
    myArray(26, 5) = 71.9
    myArray(26, 6) = 31.3
    myArray(26, 7) = 56
    myArray(26, 8) = 149
    myArray(26, 9) = 131.3
    myArray(26, 10) = 118.7
    myArray(26, 11) = 63.9
    myArray(26, 12) = 130.6
    myArray(26, 13) = 31.3

    myArray(27, 1) = 2020
    myArray(27, 2) = 49.4
    myArray(27, 3) = 75.3
    myArray(27, 4) = 22.8
    myArray(27, 5) = 16.5
    myArray(27, 6) = 92.4
    myArray(27, 7) = 139.7
    myArray(27, 8) = 345.9
    myArray(27, 9) = 321.5
    myArray(27, 10) = 177.1
    myArray(27, 11) = 16.2
    myArray(27, 12) = 35.4
    myArray(27, 13) = 9.7

    myArray(28, 1) = 2021
    myArray(28, 2) = 32
    myArray(28, 3) = 18.7
    myArray(28, 4) = 76.1
    myArray(28, 5) = 43.4
    myArray(28, 6) = 110
    myArray(28, 7) = 55
    myArray(28, 8) = 131.3
    myArray(28, 9) = 253.7
    myArray(28, 10) = 215.9
    myArray(28, 11) = 39.6
    myArray(28, 12) = 117.8
    myArray(28, 13) = 14.4

    myArray(29, 1) = 2022
    myArray(29, 2) = 8.4
    myArray(29, 3) = 5.3
    myArray(29, 4) = 60.9
    myArray(29, 5) = 34.8
    myArray(29, 6) = 5.7
    myArray(29, 7) = 225
    myArray(29, 8) = 119.9
    myArray(29, 9) = 637.1
    myArray(29, 10) = 102
    myArray(29, 11) = 112
    myArray(29, 12) = 23.3
    myArray(29, 13) = 14

    myArray(30, 1) = 2023
    myArray(30, 2) = 19.2
    myArray(30, 3) = 0.4
    myArray(30, 4) = 7.2
    myArray(30, 5) = 42.4
    myArray(30, 6) = 190.8
    myArray(30, 7) = 95.1
    myArray(30, 8) = 772.2
    myArray(30, 9) = 107.8
    myArray(30, 10) = 288.4
    myArray(30, 11) = 25.5
    myArray(30, 12) = 65.2
    myArray(30, 13) = 110.6

    data_BORYUNG = myArray

End Function


Function data_BUYEO() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

    myArray(1, 1) = 1994
    myArray(1, 2) = 18.5
    myArray(1, 3) = 14.4
    myArray(1, 4) = 54.4
    myArray(1, 5) = 41
    myArray(1, 6) = 137
    myArray(1, 7) = 137.5
    myArray(1, 8) = 96.5
    myArray(1, 9) = 286.5
    myArray(1, 10) = 34
    myArray(1, 11) = 162
    myArray(1, 12) = 23
    myArray(1, 13) = 25

    myArray(2, 1) = 1995
    myArray(2, 2) = 22.6
    myArray(2, 3) = 23.5
    myArray(2, 4) = 24.4
    myArray(2, 5) = 62
    myArray(2, 6) = 59.5
    myArray(2, 7) = 34.5
    myArray(2, 8) = 171.5
    myArray(2, 9) = 839
    myArray(2, 10) = 46.5
    myArray(2, 11) = 22
    myArray(2, 12) = 15
    myArray(2, 13) = 5.7

    myArray(3, 1) = 1996
    myArray(3, 2) = 26.4
    myArray(3, 3) = 2.8
    myArray(3, 4) = 131
    myArray(3, 5) = 45
    myArray(3, 6) = 33
    myArray(3, 7) = 289
    myArray(3, 8) = 235
    myArray(3, 9) = 67
    myArray(3, 10) = 16
    myArray(3, 11) = 90.5
    myArray(3, 12) = 76
    myArray(3, 13) = 35

    myArray(4, 1) = 1997
    myArray(4, 2) = 9
    myArray(4, 3) = 54.9
    myArray(4, 4) = 44
    myArray(4, 5) = 70
    myArray(4, 6) = 229.5
    myArray(4, 7) = 236.5
    myArray(4, 8) = 404.5
    myArray(4, 9) = 263
    myArray(4, 10) = 24.5
    myArray(4, 11) = 8
    myArray(4, 12) = 219.5
    myArray(4, 13) = 39.5

    myArray(5, 1) = 1998
    myArray(5, 2) = 40.6
    myArray(5, 3) = 47
    myArray(5, 4) = 45
    myArray(5, 5) = 200.5
    myArray(5, 6) = 130.5
    myArray(5, 7) = 324
    myArray(5, 8) = 323
    myArray(5, 9) = 451.3
    myArray(5, 10) = 313.1
    myArray(5, 11) = 75.5
    myArray(5, 12) = 46.3
    myArray(5, 13) = 3.5

    myArray(6, 1) = 1999
    myArray(6, 2) = 3.5
    myArray(6, 3) = 10
    myArray(6, 4) = 75.7
    myArray(6, 5) = 92.5
    myArray(6, 6) = 127.5
    myArray(6, 7) = 203
    myArray(6, 8) = 149
    myArray(6, 9) = 119.5
    myArray(6, 10) = 426
    myArray(6, 11) = 290
    myArray(6, 12) = 15.5
    myArray(6, 13) = 17.4

    myArray(7, 1) = 2000
    myArray(7, 2) = 41.4
    myArray(7, 3) = 2.3
    myArray(7, 4) = 14.1
    myArray(7, 5) = 62
    myArray(7, 6) = 40
    myArray(7, 7) = 248.5
    myArray(7, 8) = 248.5
    myArray(7, 9) = 543
    myArray(7, 10) = 238.5
    myArray(7, 11) = 39
    myArray(7, 12) = 29.5
    myArray(7, 13) = 13.8

    myArray(8, 1) = 2001
    myArray(8, 2) = 65
    myArray(8, 3) = 69.5
    myArray(8, 4) = 9.8
    myArray(8, 5) = 25
    myArray(8, 6) = 23.5
    myArray(8, 7) = 132
    myArray(8, 8) = 216
    myArray(8, 9) = 98
    myArray(8, 10) = 10.5
    myArray(8, 11) = 76.5
    myArray(8, 12) = 10.5
    myArray(8, 13) = 16.3

    myArray(9, 1) = 2002
    myArray(9, 2) = 72.3
    myArray(9, 3) = 6
    myArray(9, 4) = 32.5
    myArray(9, 5) = 142.5
    myArray(9, 6) = 159
    myArray(9, 7) = 70.5
    myArray(9, 8) = 208
    myArray(9, 9) = 358.5
    myArray(9, 10) = 57.5
    myArray(9, 11) = 78.5
    myArray(9, 12) = 31.5
    myArray(9, 13) = 57.2

    myArray(10, 1) = 2003
    myArray(10, 2) = 24.2
    myArray(10, 3) = 59
    myArray(10, 4) = 52
    myArray(10, 5) = 208.5
    myArray(10, 6) = 144.5
    myArray(10, 7) = 228
    myArray(10, 8) = 626.5
    myArray(10, 9) = 202
    myArray(10, 10) = 167.5
    myArray(10, 11) = 24.5
    myArray(10, 12) = 29.5
    myArray(10, 13) = 13.8

    myArray(11, 1) = 2004
    myArray(11, 2) = 18.1
    myArray(11, 3) = 26.2
    myArray(11, 4) = 63.1
    myArray(11, 5) = 73.5
    myArray(11, 6) = 109
    myArray(11, 7) = 388
    myArray(11, 8) = 296
    myArray(11, 9) = 249
    myArray(11, 10) = 176.5
    myArray(11, 11) = 1
    myArray(11, 12) = 50.5
    myArray(11, 13) = 43

    myArray(12, 1) = 2005
    myArray(12, 2) = 6
    myArray(12, 3) = 39
    myArray(12, 4) = 26.5
    myArray(12, 5) = 75
    myArray(12, 6) = 65.5
    myArray(12, 7) = 186
    myArray(12, 8) = 448.5
    myArray(12, 9) = 381.5
    myArray(12, 10) = 225.5
    myArray(12, 11) = 30.5
    myArray(12, 12) = 21
    myArray(12, 13) = 22

    myArray(13, 1) = 2006
    myArray(13, 2) = 30.2
    myArray(13, 3) = 29.5
    myArray(13, 4) = 7.8
    myArray(13, 5) = 99
    myArray(13, 6) = 81.5
    myArray(13, 7) = 111
    myArray(13, 8) = 503
    myArray(13, 9) = 83.5
    myArray(13, 10) = 37.5
    myArray(13, 11) = 15
    myArray(13, 12) = 51
    myArray(13, 13) = 27.5

    myArray(14, 1) = 2007
    myArray(14, 2) = 21.8
    myArray(14, 3) = 47.8
    myArray(14, 4) = 159
    myArray(14, 5) = 28
    myArray(14, 6) = 104
    myArray(14, 7) = 101
    myArray(14, 8) = 286
    myArray(14, 9) = 319.5
    myArray(14, 10) = 502.5
    myArray(14, 11) = 37
    myArray(14, 12) = 13
    myArray(14, 13) = 31.7

    myArray(15, 1) = 2008
    myArray(15, 2) = 39.6
    myArray(15, 3) = 11.2
    myArray(15, 4) = 42.2
    myArray(15, 5) = 38.8
    myArray(15, 6) = 51.6
    myArray(15, 7) = 260
    myArray(15, 8) = 194.3
    myArray(15, 9) = 154
    myArray(15, 10) = 48.8
    myArray(15, 11) = 24.1
    myArray(15, 12) = 14.1
    myArray(15, 13) = 23.4

    myArray(16, 1) = 2009
    myArray(16, 2) = 10.6
    myArray(16, 3) = 23.6
    myArray(16, 4) = 63.9
    myArray(16, 5) = 51
    myArray(16, 6) = 135.5
    myArray(16, 7) = 113.2
    myArray(16, 8) = 408
    myArray(16, 9) = 140.2
    myArray(16, 10) = 30.5
    myArray(16, 11) = 23.7
    myArray(16, 12) = 54.5
    myArray(16, 13) = 34.9

    myArray(17, 1) = 2010
    myArray(17, 2) = 37.1
    myArray(17, 3) = 89.5
    myArray(17, 4) = 94.9
    myArray(17, 5) = 69.6
    myArray(17, 6) = 140.7
    myArray(17, 7) = 36.1
    myArray(17, 8) = 262.1
    myArray(17, 9) = 431.1
    myArray(17, 10) = 149.8
    myArray(17, 11) = 17.8
    myArray(17, 12) = 18.6
    myArray(17, 13) = 31

    myArray(18, 1) = 2011
    myArray(18, 2) = 3.7
    myArray(18, 3) = 60.7
    myArray(18, 4) = 16
    myArray(18, 5) = 70
    myArray(18, 6) = 111.2
    myArray(18, 7) = 316
    myArray(18, 8) = 599.6
    myArray(18, 9) = 618.1
    myArray(18, 10) = 104.2
    myArray(18, 11) = 26.6
    myArray(18, 12) = 81.6
    myArray(18, 13) = 7

    myArray(19, 1) = 2012
    myArray(19, 2) = 16
    myArray(19, 3) = 3.2
    myArray(19, 4) = 60.2
    myArray(19, 5) = 109.3
    myArray(19, 6) = 19.5
    myArray(19, 7) = 71.3
    myArray(19, 8) = 302.9
    myArray(19, 9) = 573.3
    myArray(19, 10) = 186.2
    myArray(19, 11) = 83
    myArray(19, 12) = 60.7
    myArray(19, 13) = 60.2

    myArray(20, 1) = 2013
    myArray(20, 2) = 45.4
    myArray(20, 3) = 58.7
    myArray(20, 4) = 50.3
    myArray(20, 5) = 93.7
    myArray(20, 6) = 159
    myArray(20, 7) = 151.7
    myArray(20, 8) = 240.4
    myArray(20, 9) = 119.5
    myArray(20, 10) = 184.8
    myArray(20, 11) = 17.5
    myArray(20, 12) = 79.4
    myArray(20, 13) = 35.9

    myArray(21, 1) = 2014
    myArray(21, 2) = 2.2
    myArray(21, 3) = 15.3
    myArray(21, 4) = 69.3
    myArray(21, 5) = 94.1
    myArray(21, 6) = 61.5
    myArray(21, 7) = 77.8
    myArray(21, 8) = 174.7
    myArray(21, 9) = 225.1
    myArray(21, 10) = 157.5
    myArray(21, 11) = 170.5
    myArray(21, 12) = 42.4
    myArray(21, 13) = 51.7

    myArray(22, 1) = 2015
    myArray(22, 2) = 35.4
    myArray(22, 3) = 35.6
    myArray(22, 4) = 42.4
    myArray(22, 5) = 99.5
    myArray(22, 6) = 53.5
    myArray(22, 7) = 92.7
    myArray(22, 8) = 119.9
    myArray(22, 9) = 56.9
    myArray(22, 10) = 22
    myArray(22, 11) = 104
    myArray(22, 12) = 130
    myArray(22, 13) = 56.9

    myArray(23, 1) = 2016
    myArray(23, 2) = 6.6
    myArray(23, 3) = 59.6
    myArray(23, 4) = 19
    myArray(23, 5) = 164.6
    myArray(23, 6) = 121.6
    myArray(23, 7) = 49.4
    myArray(23, 8) = 341.1
    myArray(23, 9) = 33.4
    myArray(23, 10) = 133.7
    myArray(23, 11) = 120.1
    myArray(23, 12) = 17.1
    myArray(23, 13) = 63.1

    myArray(24, 1) = 2017
    myArray(24, 2) = 16
    myArray(24, 3) = 28.5
    myArray(24, 4) = 8.8
    myArray(24, 5) = 78.4
    myArray(24, 6) = 35.8
    myArray(24, 7) = 51.4
    myArray(24, 8) = 326.7
    myArray(24, 9) = 358.5
    myArray(24, 10) = 97.1
    myArray(24, 11) = 51.9
    myArray(24, 12) = 22.8
    myArray(24, 13) = 36.1

    myArray(25, 1) = 2018
    myArray(25, 2) = 25
    myArray(25, 3) = 43.1
    myArray(25, 4) = 99.3
    myArray(25, 5) = 156.5
    myArray(25, 6) = 116.1
    myArray(25, 7) = 107.1
    myArray(25, 8) = 278.8
    myArray(25, 9) = 277
    myArray(25, 10) = 98.3
    myArray(25, 11) = 159.2
    myArray(25, 12) = 66
    myArray(25, 13) = 31.5

    myArray(26, 1) = 2019
    myArray(26, 2) = 0.5
    myArray(26, 3) = 37.6
    myArray(26, 4) = 35
    myArray(26, 5) = 73.7
    myArray(26, 6) = 44.3
    myArray(26, 7) = 59.9
    myArray(26, 8) = 216.7
    myArray(26, 9) = 102.1
    myArray(26, 10) = 191.9
    myArray(26, 11) = 85.6
    myArray(26, 12) = 113.5
    myArray(26, 13) = 31.2

    myArray(27, 1) = 2020
    myArray(27, 2) = 79.6
    myArray(27, 3) = 92.4
    myArray(27, 4) = 19.3
    myArray(27, 5) = 17.7
    myArray(27, 6) = 108.5
    myArray(27, 7) = 188.4
    myArray(27, 8) = 492.6
    myArray(27, 9) = 367.8
    myArray(27, 10) = 208.9
    myArray(27, 11) = 4.4
    myArray(27, 12) = 41.8
    myArray(27, 13) = 3.4

    myArray(28, 1) = 2021
    myArray(28, 2) = 32.1
    myArray(28, 3) = 18.1
    myArray(28, 4) = 95.7
    myArray(28, 5) = 42.3
    myArray(28, 6) = 136.9
    myArray(28, 7) = 76.9
    myArray(28, 8) = 187.7
    myArray(28, 9) = 227.6
    myArray(28, 10) = 187.1
    myArray(28, 11) = 36.9
    myArray(28, 12) = 73.4
    myArray(28, 13) = 8.7

    myArray(29, 1) = 2022
    myArray(29, 2) = 3.5
    myArray(29, 3) = 2.5
    myArray(29, 4) = 76.1
    myArray(29, 5) = 62.6
    myArray(29, 6) = 4
    myArray(29, 7) = 123.4
    myArray(29, 8) = 168.5
    myArray(29, 9) = 615.6
    myArray(29, 10) = 87
    myArray(29, 11) = 103.7
    myArray(29, 12) = 36.4
    myArray(29, 13) = 17.8

    myArray(30, 1) = 2023
    myArray(30, 2) = 35.7
    myArray(30, 3) = 4.3
    myArray(30, 4) = 13.2
    myArray(30, 5) = 60.6
    myArray(30, 6) = 248.1
    myArray(30, 7) = 122.2
    myArray(30, 8) = 880.3
    myArray(30, 9) = 300.6
    myArray(30, 10) = 303
    myArray(30, 11) = 16.7
    myArray(30, 12) = 58.1
    myArray(30, 13) = 120.3

    data_BUYEO = myArray

End Function


Function data_CHEONAN() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

    myArray(1, 1) = 1994
    myArray(1, 2) = 9.1
    myArray(1, 3) = 10.1
    myArray(1, 4) = 39.5
    myArray(1, 5) = 13.5
    myArray(1, 6) = 106.5
    myArray(1, 7) = 160.5
    myArray(1, 8) = 98
    myArray(1, 9) = 418
    myArray(1, 10) = 52
    myArray(1, 11) = 220
    myArray(1, 12) = 22
    myArray(1, 13) = 21

    myArray(2, 1) = 1995
    myArray(2, 2) = 19
    myArray(2, 3) = 8.2
    myArray(2, 4) = 25.3
    myArray(2, 5) = 47
    myArray(2, 6) = 48
    myArray(2, 7) = 14.5
    myArray(2, 8) = 239.9
    myArray(2, 9) = 1082.5
    myArray(2, 10) = 29
    myArray(2, 11) = 23.5
    myArray(2, 12) = 40.2
    myArray(2, 13) = 8.9

    myArray(3, 1) = 1996
    myArray(3, 2) = 29.5
    myArray(3, 3) = 10.2
    myArray(3, 4) = 115
    myArray(3, 5) = 54.5
    myArray(3, 6) = 19
    myArray(3, 7) = 237
    myArray(3, 8) = 177.5
    myArray(3, 9) = 116.5
    myArray(3, 10) = 8
    myArray(3, 11) = 102.5
    myArray(3, 12) = 71.6
    myArray(3, 13) = 26.2

    myArray(4, 1) = 1997
    myArray(4, 2) = 10.7
    myArray(4, 3) = 44.1
    myArray(4, 4) = 30
    myArray(4, 5) = 66.5
    myArray(4, 6) = 211
    myArray(4, 7) = 191.5
    myArray(4, 8) = 305
    myArray(4, 9) = 175.5
    myArray(4, 10) = 14.5
    myArray(4, 11) = 23
    myArray(4, 12) = 153.5
    myArray(4, 13) = 43.5

    myArray(5, 1) = 1998
    myArray(5, 2) = 20.4
    myArray(5, 3) = 27.9
    myArray(5, 4) = 29.5
    myArray(5, 5) = 120.5
    myArray(5, 6) = 85
    myArray(5, 7) = 219.5
    myArray(5, 8) = 277
    myArray(5, 9) = 408.1
    myArray(5, 10) = 283
    myArray(5, 11) = 51.5
    myArray(5, 12) = 52.8
    myArray(5, 13) = 8.5

    myArray(6, 1) = 1999
    myArray(6, 2) = 2.7
    myArray(6, 3) = 2.8
    myArray(6, 4) = 46.5
    myArray(6, 5) = 88.5
    myArray(6, 6) = 121.5
    myArray(6, 7) = 163.7
    myArray(6, 8) = 138.5
    myArray(6, 9) = 313.5
    myArray(6, 10) = 324.5
    myArray(6, 11) = 134.5
    myArray(6, 12) = 16.5
    myArray(6, 13) = 11.9

    myArray(7, 1) = 2000
    myArray(7, 2) = 52.3
    myArray(7, 3) = 2.7
    myArray(7, 4) = 7.1
    myArray(7, 5) = 36
    myArray(7, 6) = 36
    myArray(7, 7) = 181
    myArray(7, 8) = 83
    myArray(7, 9) = 636
    myArray(7, 10) = 287.5
    myArray(7, 11) = 32
    myArray(7, 12) = 32
    myArray(7, 13) = 22.5

    myArray(8, 1) = 2001
    myArray(8, 2) = 43.5
    myArray(8, 3) = 44
    myArray(8, 4) = 16.5
    myArray(8, 5) = 19
    myArray(8, 6) = 15
    myArray(8, 7) = 227.5
    myArray(8, 8) = 178
    myArray(8, 9) = 194.5
    myArray(8, 10) = 12
    myArray(8, 11) = 63.5
    myArray(8, 12) = 6.3
    myArray(8, 13) = 18.4

    myArray(9, 1) = 2002
    myArray(9, 2) = 45.3
    myArray(9, 3) = 6
    myArray(9, 4) = 25.5
    myArray(9, 5) = 128
    myArray(9, 6) = 104
    myArray(9, 7) = 54
    myArray(9, 8) = 229.5
    myArray(9, 9) = 481.5
    myArray(9, 10) = 57
    myArray(9, 11) = 91.5
    myArray(9, 12) = 42.1
    myArray(9, 13) = 48.1

    myArray(10, 1) = 2003
    myArray(10, 2) = 18.6
    myArray(10, 3) = 44
    myArray(10, 4) = 38.1
    myArray(10, 5) = 172.3
    myArray(10, 6) = 106
    myArray(10, 7) = 178.6
    myArray(10, 8) = 381.2
    myArray(10, 9) = 334.6
    myArray(10, 10) = 264.2
    myArray(10, 11) = 27
    myArray(10, 12) = 46.7
    myArray(10, 13) = 17

    myArray(11, 1) = 2004
    myArray(11, 2) = 16.4
    myArray(11, 3) = 21.3
    myArray(11, 4) = 21.5
    myArray(11, 5) = 67.5
    myArray(11, 6) = 127.6
    myArray(11, 7) = 235
    myArray(11, 8) = 365.2
    myArray(11, 9) = 229.3
    myArray(11, 10) = 189
    myArray(11, 11) = 4.5
    myArray(11, 12) = 53
    myArray(11, 13) = 33

    myArray(12, 1) = 2005
    myArray(12, 2) = 3
    myArray(12, 3) = 29.8
    myArray(12, 4) = 37
    myArray(12, 5) = 53.7
    myArray(12, 6) = 48
    myArray(12, 7) = 183
    myArray(12, 8) = 313.8
    myArray(12, 9) = 202
    myArray(12, 10) = 377.5
    myArray(12, 11) = 26.7
    myArray(12, 12) = 23.5
    myArray(12, 13) = 11.3

    myArray(13, 1) = 2006
    myArray(13, 2) = 25.2
    myArray(13, 3) = 18.5
    myArray(13, 4) = 6.1
    myArray(13, 5) = 78.6
    myArray(13, 6) = 79
    myArray(13, 7) = 120
    myArray(13, 8) = 535.1
    myArray(13, 9) = 63.5
    myArray(13, 10) = 22.2
    myArray(13, 11) = 21.6
    myArray(13, 12) = 56.3
    myArray(13, 13) = 17.2

    myArray(14, 1) = 2007
    myArray(14, 2) = 9.4
    myArray(14, 3) = 34.1
    myArray(14, 4) = 108.3
    myArray(14, 5) = 35.3
    myArray(14, 6) = 126.2
    myArray(14, 7) = 106.7
    myArray(14, 8) = 215.6
    myArray(14, 9) = 470.6
    myArray(14, 10) = 353.3
    myArray(14, 11) = 43.4
    myArray(14, 12) = 15.6
    myArray(14, 13) = 43.9

    myArray(15, 1) = 2008
    myArray(15, 2) = 17.5
    myArray(15, 3) = 11.1
    myArray(15, 4) = 40.1
    myArray(15, 5) = 34
    myArray(15, 6) = 62.6
    myArray(15, 7) = 126.7
    myArray(15, 8) = 287.2
    myArray(15, 9) = 138.8
    myArray(15, 10) = 89.3
    myArray(15, 11) = 30.4
    myArray(15, 12) = 16.6
    myArray(15, 13) = 15.8

    myArray(16, 1) = 2009
    myArray(16, 2) = 13.3
    myArray(16, 3) = 16
    myArray(16, 4) = 51.6
    myArray(16, 5) = 30.6
    myArray(16, 6) = 112.6
    myArray(16, 7) = 55.6
    myArray(16, 8) = 335.8
    myArray(16, 9) = 212.3
    myArray(16, 10) = 30.8
    myArray(16, 11) = 61.1
    myArray(16, 12) = 39.7
    myArray(16, 13) = 40.5

    myArray(17, 1) = 2010
    myArray(17, 2) = 40.7
    myArray(17, 3) = 50.4
    myArray(17, 4) = 73.8
    myArray(17, 5) = 61
    myArray(17, 6) = 84
    myArray(17, 7) = 37
    myArray(17, 8) = 171
    myArray(17, 9) = 486.1
    myArray(17, 10) = 316.9
    myArray(17, 11) = 19.4
    myArray(17, 12) = 13.5
    myArray(17, 13) = 24.5

    myArray(18, 1) = 2011
    myArray(18, 2) = 7.9
    myArray(18, 3) = 31
    myArray(18, 4) = 26.5
    myArray(18, 5) = 133.2
    myArray(18, 6) = 103.3
    myArray(18, 7) = 374.6
    myArray(18, 8) = 645.1
    myArray(18, 9) = 268.2
    myArray(18, 10) = 153.2
    myArray(18, 11) = 26.5
    myArray(18, 12) = 65.8
    myArray(18, 13) = 10.5

    myArray(19, 1) = 2012
    myArray(19, 2) = 14.5
    myArray(19, 3) = 2.3
    myArray(19, 4) = 44.9
    myArray(19, 5) = 81.6
    myArray(19, 6) = 16.8
    myArray(19, 7) = 75.1
    myArray(19, 8) = 252.5
    myArray(19, 9) = 483.7
    myArray(19, 10) = 190.1
    myArray(19, 11) = 66.6
    myArray(19, 12) = 52.6
    myArray(19, 13) = 56

    myArray(20, 1) = 2013
    myArray(20, 2) = 28.5
    myArray(20, 3) = 35.2
    myArray(20, 4) = 40
    myArray(20, 5) = 56.3
    myArray(20, 6) = 123.5
    myArray(20, 7) = 102.1
    myArray(20, 8) = 308.2
    myArray(20, 9) = 173.6
    myArray(20, 10) = 117.5
    myArray(20, 11) = 12.2
    myArray(20, 12) = 58.2
    myArray(20, 13) = 40.3

    myArray(21, 1) = 2014
    myArray(21, 2) = 4.9
    myArray(21, 3) = 14.7
    myArray(21, 4) = 40.9
    myArray(21, 5) = 62.1
    myArray(21, 6) = 34.6
    myArray(21, 7) = 73.9
    myArray(21, 8) = 239
    myArray(21, 9) = 218.7
    myArray(21, 10) = 144
    myArray(21, 11) = 119.5
    myArray(21, 12) = 28.9
    myArray(21, 13) = 38.9

    myArray(22, 1) = 2015
    myArray(22, 2) = 12.7
    myArray(22, 3) = 21.5
    myArray(22, 4) = 23.3
    myArray(22, 5) = 87.6
    myArray(22, 6) = 27.5
    myArray(22, 7) = 86
    myArray(22, 8) = 136.8
    myArray(22, 9) = 64.2
    myArray(22, 10) = 29
    myArray(22, 11) = 69
    myArray(22, 12) = 128.6
    myArray(22, 13) = 41.8

    myArray(23, 1) = 2016
    myArray(23, 2) = 8
    myArray(23, 3) = 43.6
    myArray(23, 4) = 16.5
    myArray(23, 5) = 118.3
    myArray(23, 6) = 107.2
    myArray(23, 7) = 36.2
    myArray(23, 8) = 364.3
    myArray(23, 9) = 82
    myArray(23, 10) = 55
    myArray(23, 11) = 95.9
    myArray(23, 12) = 33.5
    myArray(23, 13) = 44.3

    myArray(24, 1) = 2017
    myArray(24, 2) = 13.9
    myArray(24, 3) = 32.2
    myArray(24, 4) = 6.5
    myArray(24, 5) = 42.9
    myArray(24, 6) = 14.3
    myArray(24, 7) = 15.6
    myArray(24, 8) = 788.1
    myArray(24, 9) = 291.5
    myArray(24, 10) = 43.3
    myArray(24, 11) = 14.1
    myArray(24, 12) = 23.8
    myArray(24, 13) = 18.8

    myArray(25, 1) = 2018
    myArray(25, 2) = 14
    myArray(25, 3) = 31.6
    myArray(25, 4) = 62.2
    myArray(25, 5) = 117
    myArray(25, 6) = 82.7
    myArray(25, 7) = 88.9
    myArray(25, 8) = 185.8
    myArray(25, 9) = 282.7
    myArray(25, 10) = 124.6
    myArray(25, 11) = 99.8
    myArray(25, 12) = 48.3
    myArray(25, 13) = 25.8

    myArray(26, 1) = 2019
    myArray(26, 2) = 0.6
    myArray(26, 3) = 25.5
    myArray(26, 4) = 26.9
    myArray(26, 5) = 43.9
    myArray(26, 6) = 15.1
    myArray(26, 7) = 84.9
    myArray(26, 8) = 234.7
    myArray(26, 9) = 90.7
    myArray(26, 10) = 102.8
    myArray(26, 11) = 81.9
    myArray(26, 12) = 120.6
    myArray(26, 13) = 18

    myArray(27, 1) = 2020
    myArray(27, 2) = 59.7
    myArray(27, 3) = 63.1
    myArray(27, 4) = 21.7
    myArray(27, 5) = 15.1
    myArray(27, 6) = 86.4
    myArray(27, 7) = 121.9
    myArray(27, 8) = 356.4
    myArray(27, 9) = 481.7
    myArray(27, 10) = 167.2
    myArray(27, 11) = 18.9
    myArray(27, 12) = 45.9
    myArray(27, 13) = 5.5

    myArray(28, 1) = 2021
    myArray(28, 2) = 17.8
    myArray(28, 3) = 9.2
    myArray(28, 4) = 75.3
    myArray(28, 5) = 54.7
    myArray(28, 6) = 135.9
    myArray(28, 7) = 44.8
    myArray(28, 8) = 117.6
    myArray(28, 9) = 230
    myArray(28, 10) = 250.8
    myArray(28, 11) = 49.5
    myArray(28, 12) = 67.9
    myArray(28, 13) = 5.4

    myArray(29, 1) = 2022
    myArray(29, 2) = 3.3
    myArray(29, 3) = 3.3
    myArray(29, 4) = 57.6
    myArray(29, 5) = 51.6
    myArray(29, 6) = 9.8
    myArray(29, 7) = 168
    myArray(29, 8) = 117
    myArray(29, 9) = 366.6
    myArray(29, 10) = 133.3
    myArray(29, 11) = 98.2
    myArray(29, 12) = 43.2
    myArray(29, 13) = 28.8

    myArray(30, 1) = 2023
    myArray(30, 2) = 31
    myArray(30, 3) = 3.1
    myArray(30, 4) = 16.4
    myArray(30, 5) = 29.6
    myArray(30, 6) = 116.9
    myArray(30, 7) = 178.9
    myArray(30, 8) = 574.9
    myArray(30, 9) = 196.5
    myArray(30, 10) = 180.1
    myArray(30, 11) = 28.7
    myArray(30, 12) = 56.9
    myArray(30, 13) = 86.2

    data_CHEONAN = myArray

End Function



Function data_CHEONGJU() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

    myArray(1, 1) = 1994
    myArray(1, 2) = 13.3
    myArray(1, 3) = 12.8
    myArray(1, 4) = 54.2
    myArray(1, 5) = 21.3
    myArray(1, 6) = 108.8
    myArray(1, 7) = 140.5
    myArray(1, 8) = 85.5
    myArray(1, 9) = 318.5
    myArray(1, 10) = 48.1
    myArray(1, 11) = 160.1
    myArray(1, 12) = 29.6
    myArray(1, 13) = 19.3

    myArray(2, 1) = 1995
    myArray(2, 2) = 21.5
    myArray(2, 3) = 14
    myArray(2, 4) = 34.4
    myArray(2, 5) = 64
    myArray(2, 6) = 70.7
    myArray(2, 7) = 30.9
    myArray(2, 8) = 204.9
    myArray(2, 9) = 835.4
    myArray(2, 10) = 17.5
    myArray(2, 11) = 22.6
    myArray(2, 12) = 20.3
    myArray(2, 13) = 3.6

    myArray(3, 1) = 1996
    myArray(3, 2) = 27.9
    myArray(3, 3) = 4.2
    myArray(3, 4) = 98.4
    myArray(3, 5) = 28.6
    myArray(3, 6) = 36.8
    myArray(3, 7) = 255.8
    myArray(3, 8) = 170.5
    myArray(3, 9) = 128.6
    myArray(3, 10) = 11.2
    myArray(3, 11) = 67.1
    myArray(3, 12) = 77.2
    myArray(3, 13) = 22.5

    myArray(4, 1) = 1997
    myArray(4, 2) = 12.9
    myArray(4, 3) = 39.1
    myArray(4, 4) = 31.6
    myArray(4, 5) = 58.5
    myArray(4, 6) = 179.1
    myArray(4, 7) = 210.3
    myArray(4, 8) = 425.5
    myArray(4, 9) = 211.1
    myArray(4, 10) = 55.5
    myArray(4, 11) = 8.4
    myArray(4, 12) = 180.3
    myArray(4, 13) = 44.3

    myArray(5, 1) = 1998
    myArray(5, 2) = 22
    myArray(5, 3) = 28.9
    myArray(5, 4) = 30.9
    myArray(5, 5) = 153.1
    myArray(5, 6) = 92.8
    myArray(5, 7) = 247
    myArray(5, 8) = 253
    myArray(5, 9) = 460.6
    myArray(5, 10) = 225.9
    myArray(5, 11) = 74.2
    myArray(5, 12) = 44.7
    myArray(5, 13) = 7.1

    myArray(6, 1) = 1999
    myArray(6, 2) = 1.6
    myArray(6, 3) = 3.6
    myArray(6, 4) = 54.1
    myArray(6, 5) = 91.4
    myArray(6, 6) = 102.4
    myArray(6, 7) = 191.1
    myArray(6, 8) = 122.4
    myArray(6, 9) = 197.4
    myArray(6, 10) = 281.3
    myArray(6, 11) = 252.4
    myArray(6, 12) = 15.4
    myArray(6, 13) = 13.4

    myArray(7, 1) = 2000
    myArray(7, 2) = 38.7
    myArray(7, 3) = 1.3
    myArray(7, 4) = 10.4
    myArray(7, 5) = 56.1
    myArray(7, 6) = 42.1
    myArray(7, 7) = 185.7
    myArray(7, 8) = 300
    myArray(7, 9) = 390.4
    myArray(7, 10) = 244.6
    myArray(7, 11) = 32.1
    myArray(7, 12) = 37.3
    myArray(7, 13) = 18.9

    myArray(8, 1) = 2001
    myArray(8, 2) = 56.9
    myArray(8, 3) = 50.3
    myArray(8, 4) = 11.3
    myArray(8, 5) = 12.7
    myArray(8, 6) = 14.3
    myArray(8, 7) = 217.5
    myArray(8, 8) = 171.5
    myArray(8, 9) = 135.5
    myArray(8, 10) = 11.8
    myArray(8, 11) = 75.9
    myArray(8, 12) = 6.9
    myArray(8, 13) = 19.5

    myArray(9, 1) = 2002
    myArray(9, 2) = 58.7
    myArray(9, 3) = 9
    myArray(9, 4) = 25.9
    myArray(9, 5) = 132
    myArray(9, 6) = 106.9
    myArray(9, 7) = 57.9
    myArray(9, 8) = 186.2
    myArray(9, 9) = 482.4
    myArray(9, 10) = 90.5
    myArray(9, 11) = 58
    myArray(9, 12) = 26.3
    myArray(9, 13) = 48

    myArray(10, 1) = 2003
    myArray(10, 2) = 16.2
    myArray(10, 3) = 45
    myArray(10, 4) = 38.9
    myArray(10, 5) = 192.7
    myArray(10, 6) = 113.5
    myArray(10, 7) = 186
    myArray(10, 8) = 467.2
    myArray(10, 9) = 293.9
    myArray(10, 10) = 150.6
    myArray(10, 11) = 32.5
    myArray(10, 12) = 33.1
    myArray(10, 13) = 12.2

    myArray(11, 1) = 2004
    myArray(11, 2) = 12.5
    myArray(11, 3) = 42.3
    myArray(11, 4) = 67.3
    myArray(11, 5) = 61
    myArray(11, 6) = 121.8
    myArray(11, 7) = 421.5
    myArray(11, 8) = 318.9
    myArray(11, 9) = 247.6
    myArray(11, 10) = 139
    myArray(11, 11) = 2
    myArray(11, 12) = 34
    myArray(11, 13) = 38

    myArray(12, 1) = 2005
    myArray(12, 2) = 4.6
    myArray(12, 3) = 13.8
    myArray(12, 4) = 36.8
    myArray(12, 5) = 66.1
    myArray(12, 6) = 50.7
    myArray(12, 7) = 170
    myArray(12, 8) = 373.1
    myArray(12, 9) = 334.7
    myArray(12, 10) = 295.5
    myArray(12, 11) = 54.6
    myArray(12, 12) = 16
    myArray(12, 13) = 11.3

    myArray(13, 1) = 2006
    myArray(13, 2) = 20
    myArray(13, 3) = 28.9
    myArray(13, 4) = 8.2
    myArray(13, 5) = 89.3
    myArray(13, 6) = 119.4
    myArray(13, 7) = 115.5
    myArray(13, 8) = 508
    myArray(13, 9) = 52
    myArray(13, 10) = 18.4
    myArray(13, 11) = 21.3
    myArray(13, 12) = 83.4
    myArray(13, 13) = 16.7

    myArray(14, 1) = 2007
    myArray(14, 2) = 11.2
    myArray(14, 3) = 33.3
    myArray(14, 4) = 103.2
    myArray(14, 5) = 35.8
    myArray(14, 6) = 145.5
    myArray(14, 7) = 81.2
    myArray(14, 8) = 273.2
    myArray(14, 9) = 385.5
    myArray(14, 10) = 391.4
    myArray(14, 11) = 43.5
    myArray(14, 12) = 8.8
    myArray(14, 13) = 21.9

    myArray(15, 1) = 2008
    myArray(15, 2) = 29
    myArray(15, 3) = 7.7
    myArray(15, 4) = 29.4
    myArray(15, 5) = 27
    myArray(15, 6) = 64.5
    myArray(15, 7) = 112
    myArray(15, 8) = 296.6
    myArray(15, 9) = 195.5
    myArray(15, 10) = 92.6
    myArray(15, 11) = 13.1
    myArray(15, 12) = 10.5
    myArray(15, 13) = 14.4

    myArray(16, 1) = 2009
    myArray(16, 2) = 17.8
    myArray(16, 3) = 13.1
    myArray(16, 4) = 54.9
    myArray(16, 5) = 30.4
    myArray(16, 6) = 109.6
    myArray(16, 7) = 77.2
    myArray(16, 8) = 345.7
    myArray(16, 9) = 187.5
    myArray(16, 10) = 49.5
    myArray(16, 11) = 49.5
    myArray(16, 12) = 43.9
    myArray(16, 13) = 40.7

    myArray(17, 1) = 2010
    myArray(17, 2) = 37.8
    myArray(17, 3) = 69.2
    myArray(17, 4) = 99.8
    myArray(17, 5) = 70.5
    myArray(17, 6) = 110
    myArray(17, 7) = 42.6
    myArray(17, 8) = 224.1
    myArray(17, 9) = 433.2
    myArray(17, 10) = 278.6
    myArray(17, 11) = 17.1
    myArray(17, 12) = 15.7
    myArray(17, 13) = 23.8

    myArray(18, 1) = 2011
    myArray(18, 2) = 4.5
    myArray(18, 3) = 43.2
    myArray(18, 4) = 23.5
    myArray(18, 5) = 111.2
    myArray(18, 6) = 116.2
    myArray(18, 7) = 360.7
    myArray(18, 8) = 531.9
    myArray(18, 9) = 290.2
    myArray(18, 10) = 182.5
    myArray(18, 11) = 34.5
    myArray(18, 12) = 92.6
    myArray(18, 13) = 14.6

    myArray(19, 1) = 2012
    myArray(19, 2) = 17.8
    myArray(19, 3) = 3.7
    myArray(19, 4) = 65.1
    myArray(19, 5) = 106.8
    myArray(19, 6) = 31.2
    myArray(19, 7) = 93.7
    myArray(19, 8) = 257.4
    myArray(19, 9) = 479.5
    myArray(19, 10) = 162.5
    myArray(19, 11) = 61.2
    myArray(19, 12) = 52.1
    myArray(19, 13) = 56.6

    myArray(20, 1) = 2013
    myArray(20, 2) = 30.5
    myArray(20, 3) = 33.2
    myArray(20, 4) = 46.8
    myArray(20, 5) = 65
    myArray(20, 6) = 97.9
    myArray(20, 7) = 229.9
    myArray(20, 8) = 253.6
    myArray(20, 9) = 183.9
    myArray(20, 10) = 162.6
    myArray(20, 11) = 25
    myArray(20, 12) = 75
    myArray(20, 13) = 37.3

    myArray(21, 1) = 2014
    myArray(21, 2) = 5.9
    myArray(21, 3) = 6.8
    myArray(21, 4) = 51.1
    myArray(21, 5) = 43.7
    myArray(21, 6) = 35
    myArray(21, 7) = 92.6
    myArray(21, 8) = 125.1
    myArray(21, 9) = 197.5
    myArray(21, 10) = 147.5
    myArray(21, 11) = 151.1
    myArray(21, 12) = 24.8
    myArray(21, 13) = 32.6

    myArray(22, 1) = 2015
    myArray(22, 2) = 16
    myArray(22, 3) = 26.5
    myArray(22, 4) = 44.1
    myArray(22, 5) = 109.1
    myArray(22, 6) = 24.4
    myArray(22, 7) = 83.3
    myArray(22, 8) = 141.4
    myArray(22, 9) = 54.3
    myArray(22, 10) = 20.1
    myArray(22, 11) = 90.5
    myArray(22, 12) = 107.5
    myArray(22, 13) = 39.7

    myArray(23, 1) = 2016
    myArray(23, 2) = 5.7
    myArray(23, 3) = 45.5
    myArray(23, 4) = 13.2
    myArray(23, 5) = 132.1
    myArray(23, 6) = 84.4
    myArray(23, 7) = 39.9
    myArray(23, 8) = 320
    myArray(23, 9) = 69
    myArray(23, 10) = 78.1
    myArray(23, 11) = 83.6
    myArray(23, 12) = 26.4
    myArray(23, 13) = 40.1

    myArray(24, 1) = 2017
    myArray(24, 2) = 12
    myArray(24, 3) = 38.7
    myArray(24, 4) = 8.9
    myArray(24, 5) = 61.7
    myArray(24, 6) = 11.9
    myArray(24, 7) = 17.5
    myArray(24, 8) = 789.1
    myArray(24, 9) = 225.2
    myArray(24, 10) = 78.3
    myArray(24, 11) = 23.1
    myArray(24, 12) = 13.7
    myArray(24, 13) = 21.1

    myArray(25, 1) = 2018
    myArray(25, 2) = 17.6
    myArray(25, 3) = 30.6
    myArray(25, 4) = 81.7
    myArray(25, 5) = 133
    myArray(25, 6) = 92
    myArray(25, 7) = 63.3
    myArray(25, 8) = 324.9
    myArray(25, 9) = 247.9
    myArray(25, 10) = 204
    myArray(25, 11) = 112.2
    myArray(25, 12) = 45.9
    myArray(25, 13) = 28.5

    myArray(26, 1) = 2019
    myArray(26, 2) = 0.1
    myArray(26, 3) = 23
    myArray(26, 4) = 20.3
    myArray(26, 5) = 60.8
    myArray(26, 6) = 20.3
    myArray(26, 7) = 82.5
    myArray(26, 8) = 204.8
    myArray(26, 9) = 80.5
    myArray(26, 10) = 155.1
    myArray(26, 11) = 84.3
    myArray(26, 12) = 104.9
    myArray(26, 13) = 20.1

    myArray(27, 1) = 2020
    myArray(27, 2) = 62
    myArray(27, 3) = 62.7
    myArray(27, 4) = 22.9
    myArray(27, 5) = 15.7
    myArray(27, 6) = 65.3
    myArray(27, 7) = 145.9
    myArray(27, 8) = 386.6
    myArray(27, 9) = 385.8
    myArray(27, 10) = 160.6
    myArray(27, 11) = 5.8
    myArray(27, 12) = 41
    myArray(27, 13) = 4.3

    myArray(28, 1) = 2021
    myArray(28, 2) = 12.7
    myArray(28, 3) = 7.5
    myArray(28, 4) = 76.6
    myArray(28, 5) = 46.4
    myArray(28, 6) = 136.4
    myArray(28, 7) = 75.4
    myArray(28, 8) = 138.1
    myArray(28, 9) = 233.1
    myArray(28, 10) = 185
    myArray(28, 11) = 29.4
    myArray(28, 12) = 57.3
    myArray(28, 13) = 3.7

    myArray(29, 1) = 2022
    myArray(29, 2) = 1.4
    myArray(29, 3) = 2.4
    myArray(29, 4) = 59
    myArray(29, 5) = 45.2
    myArray(29, 6) = 9.1
    myArray(29, 7) = 129.6
    myArray(29, 8) = 171.7
    myArray(29, 9) = 519.4
    myArray(29, 10) = 116
    myArray(29, 11) = 105.9
    myArray(29, 12) = 56.7
    myArray(29, 13) = 20

    myArray(30, 1) = 2023
    myArray(30, 2) = 28
    myArray(30, 3) = 2.8
    myArray(30, 4) = 18.8
    myArray(30, 5) = 30.1
    myArray(30, 6) = 202.4
    myArray(30, 7) = 100.5
    myArray(30, 8) = 698.5
    myArray(30, 9) = 297.7
    myArray(30, 10) = 270.6
    myArray(30, 11) = 17.4
    myArray(30, 12) = 41.5
    myArray(30, 13) = 95.3

    data_CHEONGJU = myArray

End Function

Function data_CHUNGJU() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

    myArray(1, 1) = 1994
    myArray(1, 2) = 8.6
    myArray(1, 3) = 6.3
    myArray(1, 4) = 51.7
    myArray(1, 5) = 15.2
    myArray(1, 6) = 90
    myArray(1, 7) = 291.5
    myArray(1, 8) = 108.2
    myArray(1, 9) = 294.7
    myArray(1, 10) = 79.5
    myArray(1, 11) = 144.5
    myArray(1, 12) = 15.5
    myArray(1, 13) = 12.7

    myArray(2, 1) = 1995
    myArray(2, 2) = 11.9
    myArray(2, 3) = 6.8
    myArray(2, 4) = 29.3
    myArray(2, 5) = 46.2
    myArray(2, 6) = 45.5
    myArray(2, 7) = 19
    myArray(2, 8) = 290.8
    myArray(2, 9) = 802
    myArray(2, 10) = 16
    myArray(2, 11) = 23
    myArray(2, 12) = 23.9
    myArray(2, 13) = 4.2

    myArray(3, 1) = 1996
    myArray(3, 2) = 27.8
    myArray(3, 3) = 1.6
    myArray(3, 4) = 101.7
    myArray(3, 5) = 36.5
    myArray(3, 6) = 29.1
    myArray(3, 7) = 203.5
    myArray(3, 8) = 207
    myArray(3, 9) = 126
    myArray(3, 10) = 26.5
    myArray(3, 11) = 83
    myArray(3, 12) = 74.1
    myArray(3, 13) = 22.2

    myArray(4, 1) = 1997
    myArray(4, 2) = 5
    myArray(4, 3) = 44.3
    myArray(4, 4) = 23.2
    myArray(4, 5) = 60
    myArray(4, 6) = 193.5
    myArray(4, 7) = 147.8
    myArray(4, 8) = 308
    myArray(4, 9) = 179.9
    myArray(4, 10) = 52.5
    myArray(4, 11) = 35.6
    myArray(4, 12) = 128.5
    myArray(4, 13) = 41.4

    myArray(5, 1) = 1998
    myArray(5, 2) = 22.5
    myArray(5, 3) = 26.4
    myArray(5, 4) = 33.8
    myArray(5, 5) = 145.3
    myArray(5, 6) = 90
    myArray(5, 7) = 217.7
    myArray(5, 8) = 286.6
    myArray(5, 9) = 541.7
    myArray(5, 10) = 183
    myArray(5, 11) = 64.5
    myArray(5, 12) = 38.5
    myArray(5, 13) = 2.5

    myArray(6, 1) = 1999
    myArray(6, 2) = 3.1
    myArray(6, 3) = 1.9
    myArray(6, 4) = 54.2
    myArray(6, 5) = 96
    myArray(6, 6) = 103.2
    myArray(6, 7) = 168.5
    myArray(6, 8) = 112.7
    myArray(6, 9) = 299.6
    myArray(6, 10) = 239.6
    myArray(6, 11) = 195.1
    myArray(6, 12) = 34.5
    myArray(6, 13) = 7

    myArray(7, 1) = 2000
    myArray(7, 2) = 47.5
    myArray(7, 3) = 3.3
    myArray(7, 4) = 14.5
    myArray(7, 5) = 42.5
    myArray(7, 6) = 54
    myArray(7, 7) = 248.5
    myArray(7, 8) = 259.5
    myArray(7, 9) = 260.5
    myArray(7, 10) = 260.8
    myArray(7, 11) = 25.5
    myArray(7, 12) = 35.5
    myArray(7, 13) = 17.5

    myArray(8, 1) = 2001
    myArray(8, 2) = 47
    myArray(8, 3) = 52.2
    myArray(8, 4) = 8.3
    myArray(8, 5) = 12
    myArray(8, 6) = 4.6
    myArray(8, 7) = 238.9
    myArray(8, 8) = 241.6
    myArray(8, 9) = 82.6
    myArray(8, 10) = 13.8
    myArray(8, 11) = 82.1
    myArray(8, 12) = 4.4
    myArray(8, 13) = 10.6

    myArray(9, 1) = 2002
    myArray(9, 2) = 49.6
    myArray(9, 3) = 4.2
    myArray(9, 4) = 28.5
    myArray(9, 5) = 151.2
    myArray(9, 6) = 105
    myArray(9, 7) = 74.7
    myArray(9, 8) = 190.2
    myArray(9, 9) = 653
    myArray(9, 10) = 92.6
    myArray(9, 11) = 52.2
    myArray(9, 12) = 9.8
    myArray(9, 13) = 58.6

    myArray(10, 1) = 2003
    myArray(10, 2) = 17.2
    myArray(10, 3) = 59.2
    myArray(10, 4) = 58.2
    myArray(10, 5) = 170.8
    myArray(10, 6) = 117.8
    myArray(10, 7) = 152.1
    myArray(10, 8) = 382.8
    myArray(10, 9) = 314.7
    myArray(10, 10) = 268.1
    myArray(10, 11) = 27.5
    myArray(10, 12) = 55
    myArray(10, 13) = 17.8

    myArray(11, 1) = 2004
    myArray(11, 2) = 16.6
    myArray(11, 3) = 32
    myArray(11, 4) = 29.4
    myArray(11, 5) = 81
    myArray(11, 6) = 124.9
    myArray(11, 7) = 335
    myArray(11, 8) = 410.7
    myArray(11, 9) = 192.2
    myArray(11, 10) = 144.1
    myArray(11, 11) = 1.4
    myArray(11, 12) = 32.5
    myArray(11, 13) = 25.4

    myArray(12, 1) = 2005
    myArray(12, 2) = 2.8
    myArray(12, 3) = 20.8
    myArray(12, 4) = 43.1
    myArray(12, 5) = 63.1
    myArray(12, 6) = 53.9
    myArray(12, 7) = 178.7
    myArray(12, 8) = 381.6
    myArray(12, 9) = 226.1
    myArray(12, 10) = 320
    myArray(12, 11) = 63.4
    myArray(12, 12) = 15.7
    myArray(12, 13) = 11.7

    myArray(13, 1) = 2006
    myArray(13, 2) = 27.1
    myArray(13, 3) = 34.9
    myArray(13, 4) = 5.9
    myArray(13, 5) = 91.8
    myArray(13, 6) = 95.1
    myArray(13, 7) = 128.5
    myArray(13, 8) = 666.9
    myArray(13, 9) = 71.5
    myArray(13, 10) = 21.7
    myArray(13, 11) = 23.1
    myArray(13, 12) = 53.1
    myArray(13, 13) = 14.3

    myArray(14, 1) = 2007
    myArray(14, 2) = 5.5
    myArray(14, 3) = 38.5
    myArray(14, 4) = 112.7
    myArray(14, 5) = 18.3
    myArray(14, 6) = 116.5
    myArray(14, 7) = 90.1
    myArray(14, 8) = 282.7
    myArray(14, 9) = 366
    myArray(14, 10) = 332.7
    myArray(14, 11) = 32.8
    myArray(14, 12) = 22
    myArray(14, 13) = 21.4

    myArray(15, 1) = 2008
    myArray(15, 2) = 29.3
    myArray(15, 3) = 8.2
    myArray(15, 4) = 43.1
    myArray(15, 5) = 31.5
    myArray(15, 6) = 70.9
    myArray(15, 7) = 78.1
    myArray(15, 8) = 319.8
    myArray(15, 9) = 192.5
    myArray(15, 10) = 71.1
    myArray(15, 11) = 16
    myArray(15, 12) = 10.3
    myArray(15, 13) = 11.7

    myArray(16, 1) = 2009
    myArray(16, 2) = 16.7
    myArray(16, 3) = 15.8
    myArray(16, 4) = 52
    myArray(16, 5) = 30.7
    myArray(16, 6) = 97.1
    myArray(16, 7) = 89.5
    myArray(16, 8) = 316.2
    myArray(16, 9) = 142.5
    myArray(16, 10) = 70.6
    myArray(16, 11) = 45
    myArray(16, 12) = 31.2
    myArray(16, 13) = 29.5

    myArray(17, 1) = 2010
    myArray(17, 2) = 44.3
    myArray(17, 3) = 70.8
    myArray(17, 4) = 85.3
    myArray(17, 5) = 69.5
    myArray(17, 6) = 97
    myArray(17, 7) = 50.6
    myArray(17, 8) = 112.2
    myArray(17, 9) = 345.1
    myArray(17, 10) = 287.8
    myArray(17, 11) = 21
    myArray(17, 12) = 14.3
    myArray(17, 13) = 14.4

    myArray(18, 1) = 2011
    myArray(18, 2) = 2.7
    myArray(18, 3) = 45.9
    myArray(18, 4) = 30.6
    myArray(18, 5) = 157.8
    myArray(18, 6) = 187.7
    myArray(18, 7) = 452.6
    myArray(18, 8) = 603.9
    myArray(18, 9) = 289.4
    myArray(18, 10) = 158.6
    myArray(18, 11) = 61.5
    myArray(18, 12) = 67
    myArray(18, 13) = 15.6

    myArray(19, 1) = 2012
    myArray(19, 2) = 9.6
    myArray(19, 3) = 1.7
    myArray(19, 4) = 66.4
    myArray(19, 5) = 84.5
    myArray(19, 6) = 61
    myArray(19, 7) = 58.8
    myArray(19, 8) = 265.7
    myArray(19, 9) = 403.3
    myArray(19, 10) = 177.2
    myArray(19, 11) = 62
    myArray(19, 12) = 48
    myArray(19, 13) = 52.1

    myArray(20, 1) = 2013
    myArray(20, 2) = 40.5
    myArray(20, 3) = 36.9
    myArray(20, 4) = 48
    myArray(20, 5) = 84.7
    myArray(20, 6) = 92.5
    myArray(20, 7) = 126.6
    myArray(20, 8) = 240.7
    myArray(20, 9) = 222.2
    myArray(20, 10) = 122.2
    myArray(20, 11) = 12.1
    myArray(20, 12) = 44
    myArray(20, 13) = 32.2

    myArray(21, 1) = 2014
    myArray(21, 2) = 14
    myArray(21, 3) = 18.9
    myArray(21, 4) = 37.7
    myArray(21, 5) = 39.6
    myArray(21, 6) = 26.3
    myArray(21, 7) = 63.3
    myArray(21, 8) = 92.6
    myArray(21, 9) = 284.3
    myArray(21, 10) = 122.7
    myArray(21, 11) = 153.8
    myArray(21, 12) = 23.5
    myArray(21, 13) = 22.9

    myArray(22, 1) = 2015
    myArray(22, 2) = 15.6
    myArray(22, 3) = 22.8
    myArray(22, 4) = 31.7
    myArray(22, 5) = 88.9
    myArray(22, 6) = 23
    myArray(22, 7) = 75
    myArray(22, 8) = 181.6
    myArray(22, 9) = 71.8
    myArray(22, 10) = 33.8
    myArray(22, 11) = 60.2
    myArray(22, 12) = 89.9
    myArray(22, 13) = 37.5

    myArray(23, 1) = 2016
    myArray(23, 2) = 1.8
    myArray(23, 3) = 50.1
    myArray(23, 4) = 11.9
    myArray(23, 5) = 97.3
    myArray(23, 6) = 70
    myArray(23, 7) = 38.9
    myArray(23, 8) = 374.4
    myArray(23, 9) = 44
    myArray(23, 10) = 60.8
    myArray(23, 11) = 102.6
    myArray(23, 12) = 22.9
    myArray(23, 13) = 42.4

    myArray(24, 1) = 2017
    myArray(24, 2) = 18
    myArray(24, 3) = 36.2
    myArray(24, 4) = 22.5
    myArray(24, 5) = 71.4
    myArray(24, 6) = 32.2
    myArray(24, 7) = 43.7
    myArray(24, 8) = 464.3
    myArray(24, 9) = 257.9
    myArray(24, 10) = 62.4
    myArray(24, 11) = 21.2
    myArray(24, 12) = 17.6
    myArray(24, 13) = 25.5

    myArray(25, 1) = 2018
    myArray(25, 2) = 14.3
    myArray(25, 3) = 35.8
    myArray(25, 4) = 75.1
    myArray(25, 5) = 107.9
    myArray(25, 6) = 180
    myArray(25, 7) = 63.7
    myArray(25, 8) = 149.1
    myArray(25, 9) = 353.3
    myArray(25, 10) = 184.9
    myArray(25, 11) = 96
    myArray(25, 12) = 50
    myArray(25, 13) = 39

    myArray(26, 1) = 2019
    myArray(26, 2) = 4.1
    myArray(26, 3) = 29
    myArray(26, 4) = 27.9
    myArray(26, 5) = 58.5
    myArray(26, 6) = 15.4
    myArray(26, 7) = 59.6
    myArray(26, 8) = 161.4
    myArray(26, 9) = 102.6
    myArray(26, 10) = 165.9
    myArray(26, 11) = 59
    myArray(26, 12) = 84.6
    myArray(26, 13) = 27.5

    myArray(27, 1) = 2020
    myArray(27, 2) = 60.2
    myArray(27, 3) = 61.9
    myArray(27, 4) = 20.7
    myArray(27, 5) = 25.9
    myArray(27, 6) = 109.7
    myArray(27, 7) = 112.1
    myArray(27, 8) = 352.2
    myArray(27, 9) = 505.6
    myArray(27, 10) = 146.4
    myArray(27, 11) = 10.7
    myArray(27, 12) = 30
    myArray(27, 13) = 11.1

    myArray(28, 1) = 2021
    myArray(28, 2) = 13.6
    myArray(28, 3) = 12.3
    myArray(28, 4) = 80.5
    myArray(28, 5) = 63.4
    myArray(28, 6) = 178.4
    myArray(28, 7) = 130.4
    myArray(28, 8) = 310.7
    myArray(28, 9) = 239.9
    myArray(28, 10) = 240.3
    myArray(28, 11) = 45.5
    myArray(28, 12) = 44.9
    myArray(28, 13) = 5.6

    myArray(29, 1) = 2022
    myArray(29, 2) = 2
    myArray(29, 3) = 4.3
    myArray(29, 4) = 79.9
    myArray(29, 5) = 45.8
    myArray(29, 6) = 8.6
    myArray(29, 7) = 219
    myArray(29, 8) = 350.5
    myArray(29, 9) = 457.3
    myArray(29, 10) = 102
    myArray(29, 11) = 96.5
    myArray(29, 12) = 79.5
    myArray(29, 13) = 18.8

    myArray(30, 1) = 2023
    myArray(30, 2) = 23
    myArray(30, 3) = 3.3
    myArray(30, 4) = 16.7
    myArray(30, 5) = 32.1
    myArray(30, 6) = 123.8
    myArray(30, 7) = 239.6
    myArray(30, 8) = 554.2
    myArray(30, 9) = 248
    myArray(30, 10) = 242.5
    myArray(30, 11) = 31.3
    myArray(30, 12) = 50.8
    myArray(30, 13) = 89.9

    data_CHUNGJU = myArray

End Function


Function data_DAEJEON() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

    myArray(1, 1) = 1994
    myArray(1, 2) = 17.9
    myArray(1, 3) = 16.8
    myArray(1, 4) = 46.5
    myArray(1, 5) = 38.7
    myArray(1, 6) = 138.4
    myArray(1, 7) = 115.1
    myArray(1, 8) = 105.3
    myArray(1, 9) = 145.9
    myArray(1, 10) = 37.9
    myArray(1, 11) = 145.3
    myArray(1, 12) = 24.3
    myArray(1, 13) = 25.8

    myArray(2, 1) = 1995
    myArray(2, 2) = 23.5
    myArray(2, 3) = 16.9
    myArray(2, 4) = 33.8
    myArray(2, 5) = 54.7
    myArray(2, 6) = 62.2
    myArray(2, 7) = 33.6
    myArray(2, 8) = 155.4
    myArray(2, 9) = 641.9
    myArray(2, 10) = 53.4
    myArray(2, 11) = 36
    myArray(2, 12) = 17.5
    myArray(2, 13) = 7.3

    myArray(3, 1) = 1996
    myArray(3, 2) = 32.7
    myArray(3, 3) = 4.4
    myArray(3, 4) = 138
    myArray(3, 5) = 49.8
    myArray(3, 6) = 62.9
    myArray(3, 7) = 411.4
    myArray(3, 8) = 257.7
    myArray(3, 9) = 114.4
    myArray(3, 10) = 11.4
    myArray(3, 11) = 90.8
    myArray(3, 12) = 77.1
    myArray(3, 13) = 28.6

    myArray(4, 1) = 1997
    myArray(4, 2) = 15.6
    myArray(4, 3) = 51.1
    myArray(4, 4) = 37.1
    myArray(4, 5) = 55.4
    myArray(4, 6) = 200.9
    myArray(4, 7) = 267.5
    myArray(4, 8) = 424.2
    myArray(4, 9) = 463.5
    myArray(4, 10) = 30.2
    myArray(4, 11) = 7.7
    myArray(4, 12) = 168.2
    myArray(4, 13) = 44.5

    myArray(5, 1) = 1998
    myArray(5, 2) = 33.3
    myArray(5, 3) = 36.3
    myArray(5, 4) = 31.1
    myArray(5, 5) = 154.3
    myArray(5, 6) = 119.5
    myArray(5, 7) = 297.2
    myArray(5, 8) = 256.1
    myArray(5, 9) = 781.7
    myArray(5, 10) = 254.7
    myArray(5, 11) = 71.5
    myArray(5, 12) = 31.6
    myArray(5, 13) = 2.7

    myArray(6, 1) = 1999
    myArray(6, 2) = 1.8
    myArray(6, 3) = 12.2
    myArray(6, 4) = 79.4
    myArray(6, 5) = 103
    myArray(6, 6) = 116.8
    myArray(6, 7) = 245.7
    myArray(6, 8) = 137.8
    myArray(6, 9) = 203
    myArray(6, 10) = 359.5
    myArray(6, 11) = 171.6
    myArray(6, 12) = 16.5
    myArray(6, 13) = 7.9

    myArray(7, 1) = 2000
    myArray(7, 2) = 27.5
    myArray(7, 3) = 4.1
    myArray(7, 4) = 17.8
    myArray(7, 5) = 67.8
    myArray(7, 6) = 54.3
    myArray(7, 7) = 238.3
    myArray(7, 8) = 470.1
    myArray(7, 9) = 473.6
    myArray(7, 10) = 263.2
    myArray(7, 11) = 24.6
    myArray(7, 12) = 44.6
    myArray(7, 13) = 21.6

    myArray(8, 1) = 2001
    myArray(8, 2) = 61.2
    myArray(8, 3) = 70
    myArray(8, 4) = 16
    myArray(8, 5) = 20.4
    myArray(8, 6) = 30.2
    myArray(8, 7) = 234.2
    myArray(8, 8) = 171
    myArray(8, 9) = 78.1
    myArray(8, 10) = 25.2
    myArray(8, 11) = 91.2
    myArray(8, 12) = 10.8
    myArray(8, 13) = 20.4

    myArray(9, 1) = 2002
    myArray(9, 2) = 92.1
    myArray(9, 3) = 12
    myArray(9, 4) = 33.5
    myArray(9, 5) = 155.5
    myArray(9, 6) = 130.5
    myArray(9, 7) = 55.4
    myArray(9, 8) = 149.1
    myArray(9, 9) = 538.8
    myArray(9, 10) = 77
    myArray(9, 11) = 67.8
    myArray(9, 12) = 24
    myArray(9, 13) = 43

    myArray(10, 1) = 2003
    myArray(10, 2) = 11.2
    myArray(10, 3) = 59.2
    myArray(10, 4) = 44.2
    myArray(10, 5) = 217.5
    myArray(10, 6) = 119.5
    myArray(10, 7) = 186.4
    myArray(10, 8) = 576.3
    myArray(10, 9) = 254.9
    myArray(10, 10) = 208.5
    myArray(10, 11) = 21.5
    myArray(10, 12) = 32.6
    myArray(10, 13) = 17.1

    myArray(11, 1) = 2004
    myArray(11, 2) = 10.9
    myArray(11, 3) = 30.6
    myArray(11, 4) = 83.2
    myArray(11, 5) = 73.1
    myArray(11, 6) = 109
    myArray(11, 7) = 383.5
    myArray(11, 8) = 391
    myArray(11, 9) = 198.3
    myArray(11, 10) = 133.7
    myArray(11, 11) = 5
    myArray(11, 12) = 37.1
    myArray(11, 13) = 41.1

    myArray(12, 1) = 2005
    myArray(12, 2) = 6
    myArray(12, 3) = 37.5
    myArray(12, 4) = 38.8
    myArray(12, 5) = 48.5
    myArray(12, 6) = 60.5
    myArray(12, 7) = 209.6
    myArray(12, 8) = 463.3
    myArray(12, 9) = 499.5
    myArray(12, 10) = 226.4
    myArray(12, 11) = 30.5
    myArray(12, 12) = 20.3
    myArray(12, 13) = 15.2

    myArray(13, 1) = 2006
    myArray(13, 2) = 31.2
    myArray(13, 3) = 33.1
    myArray(13, 4) = 8.1
    myArray(13, 5) = 94.2
    myArray(13, 6) = 119.7
    myArray(13, 7) = 131
    myArray(13, 8) = 531
    myArray(13, 9) = 113.6
    myArray(13, 10) = 24.1
    myArray(13, 11) = 19.3
    myArray(13, 12) = 60
    myArray(13, 13) = 29.9

    myArray(14, 1) = 2007
    myArray(14, 2) = 14
    myArray(14, 3) = 45
    myArray(14, 4) = 117.5
    myArray(14, 5) = 28.6
    myArray(14, 6) = 130.1
    myArray(14, 7) = 133
    myArray(14, 8) = 275.7
    myArray(14, 9) = 373
    myArray(14, 10) = 549.9
    myArray(14, 11) = 47.4
    myArray(14, 12) = 9.8
    myArray(14, 13) = 26.9

    myArray(15, 1) = 2008
    myArray(15, 2) = 45.3
    myArray(15, 3) = 9.1
    myArray(15, 4) = 29.1
    myArray(15, 5) = 34.4
    myArray(15, 6) = 59.2
    myArray(15, 7) = 148.3
    myArray(15, 8) = 253.4
    myArray(15, 9) = 325.2
    myArray(15, 10) = 85.5
    myArray(15, 11) = 19.6
    myArray(15, 12) = 12.1
    myArray(15, 13) = 16.4

    myArray(16, 1) = 2009
    myArray(16, 2) = 15.4
    myArray(16, 3) = 27.5
    myArray(16, 4) = 60.3
    myArray(16, 5) = 34.5
    myArray(16, 6) = 124.3
    myArray(16, 7) = 87.3
    myArray(16, 8) = 429.2
    myArray(16, 9) = 148.3
    myArray(16, 10) = 46
    myArray(16, 11) = 24.7
    myArray(16, 12) = 54.7
    myArray(16, 13) = 38.2

    myArray(17, 1) = 2010
    myArray(17, 2) = 46.4
    myArray(17, 3) = 81.5
    myArray(17, 4) = 100.1
    myArray(17, 5) = 88.5
    myArray(17, 6) = 117.6
    myArray(17, 7) = 65.6
    myArray(17, 8) = 223.1
    myArray(17, 9) = 376.4
    myArray(17, 10) = 250.5
    myArray(17, 11) = 17.9
    myArray(17, 12) = 16.4
    myArray(17, 13) = 35.7

    myArray(18, 1) = 2011
    myArray(18, 2) = 4
    myArray(18, 3) = 44.8
    myArray(18, 4) = 19
    myArray(18, 5) = 71
    myArray(18, 6) = 162
    myArray(18, 7) = 391.6
    myArray(18, 8) = 587.3
    myArray(18, 9) = 420.3
    myArray(18, 10) = 91.7
    myArray(18, 11) = 37
    myArray(18, 12) = 103.2
    myArray(18, 13) = 11.5

    myArray(19, 1) = 2012
    myArray(19, 2) = 16.4
    myArray(19, 3) = 2.5
    myArray(19, 4) = 54.6
    myArray(19, 5) = 66.2
    myArray(19, 6) = 24
    myArray(19, 7) = 57.8
    myArray(19, 8) = 277.6
    myArray(19, 9) = 463.6
    myArray(19, 10) = 242.5
    myArray(19, 11) = 81.3
    myArray(19, 12) = 58.4
    myArray(19, 13) = 64.6

    myArray(20, 1) = 2013
    myArray(20, 2) = 46.2
    myArray(20, 3) = 54.2
    myArray(20, 4) = 52.8
    myArray(20, 5) = 86.8
    myArray(20, 6) = 110.4
    myArray(20, 7) = 162.6
    myArray(20, 8) = 218.7
    myArray(20, 9) = 126.6
    myArray(20, 10) = 146.4
    myArray(20, 11) = 19.6
    myArray(20, 12) = 63.1
    myArray(20, 13) = 32.8

    myArray(21, 1) = 2014
    myArray(21, 2) = 6.5
    myArray(21, 3) = 8.5
    myArray(21, 4) = 67.2
    myArray(21, 5) = 59.4
    myArray(21, 6) = 49.7
    myArray(21, 7) = 143.7
    myArray(21, 8) = 177.2
    myArray(21, 9) = 240.9
    myArray(21, 10) = 118
    myArray(21, 11) = 169.4
    myArray(21, 12) = 40.7
    myArray(21, 13) = 36.5

    myArray(22, 1) = 2015
    myArray(22, 2) = 31.5
    myArray(22, 3) = 27
    myArray(22, 4) = 44.7
    myArray(22, 5) = 95.2
    myArray(22, 6) = 28.9
    myArray(22, 7) = 119.8
    myArray(22, 8) = 145.6
    myArray(22, 9) = 51.6
    myArray(22, 10) = 18.5
    myArray(22, 11) = 94.1
    myArray(22, 12) = 126.1
    myArray(22, 13) = 39.7

    myArray(23, 1) = 2016
    myArray(23, 2) = 11.6
    myArray(23, 3) = 45.8
    myArray(23, 4) = 40.3
    myArray(23, 5) = 154.9
    myArray(23, 6) = 85.1
    myArray(23, 7) = 62.5
    myArray(23, 8) = 367.9
    myArray(23, 9) = 57.4
    myArray(23, 10) = 196
    myArray(23, 11) = 122.6
    myArray(23, 12) = 29.5
    myArray(23, 13) = 54.8

    myArray(24, 1) = 2017
    myArray(24, 2) = 15
    myArray(24, 3) = 42
    myArray(24, 4) = 11.6
    myArray(24, 5) = 77.7
    myArray(24, 6) = 29.3
    myArray(24, 7) = 35.3
    myArray(24, 8) = 434.5
    myArray(24, 9) = 293.8
    myArray(24, 10) = 111.4
    myArray(24, 11) = 28.3
    myArray(24, 12) = 15.1
    myArray(24, 13) = 33.5

    myArray(25, 1) = 2018
    myArray(25, 2) = 23.9
    myArray(25, 3) = 40.5
    myArray(25, 4) = 108.4
    myArray(25, 5) = 155.3
    myArray(25, 6) = 95.9
    myArray(25, 7) = 115.8
    myArray(25, 8) = 226.9
    myArray(25, 9) = 408.6
    myArray(25, 10) = 149.4
    myArray(25, 11) = 133.9
    myArray(25, 12) = 49.8
    myArray(25, 13) = 33.7

    myArray(26, 1) = 2019
    myArray(26, 2) = 1.7
    myArray(26, 3) = 46.3
    myArray(26, 4) = 33.7
    myArray(26, 5) = 91.6
    myArray(26, 6) = 35.6
    myArray(26, 7) = 77.9
    myArray(26, 8) = 199
    myArray(26, 9) = 104.3
    myArray(26, 10) = 167
    myArray(26, 11) = 106.1
    myArray(26, 12) = 94
    myArray(26, 13) = 27

    myArray(27, 1) = 2020
    myArray(27, 2) = 78.5
    myArray(27, 3) = 91.2
    myArray(27, 4) = 24.4
    myArray(27, 5) = 17.8
    myArray(27, 6) = 80.4
    myArray(27, 7) = 192.5
    myArray(27, 8) = 544.9
    myArray(27, 9) = 361.6
    myArray(27, 10) = 173.6
    myArray(27, 11) = 3.2
    myArray(27, 12) = 41.8
    myArray(27, 13) = 4.1

    myArray(28, 1) = 2021
    myArray(28, 2) = 23.6
    myArray(28, 3) = 14.1
    myArray(28, 4) = 95.2
    myArray(28, 5) = 47.4
    myArray(28, 6) = 134.2
    myArray(28, 7) = 105.9
    myArray(28, 8) = 151.8
    myArray(28, 9) = 289.2
    myArray(28, 10) = 161.2
    myArray(28, 11) = 40.8
    myArray(28, 12) = 41.7
    myArray(28, 13) = 4.4

    myArray(29, 1) = 2022
    myArray(29, 2) = 1.2
    myArray(29, 3) = 1.4
    myArray(29, 4) = 74
    myArray(29, 5) = 69.7
    myArray(29, 6) = 8.1
    myArray(29, 7) = 117.6
    myArray(29, 8) = 195
    myArray(29, 9) = 496.1
    myArray(29, 10) = 90.2
    myArray(29, 11) = 89.3
    myArray(29, 12) = 45.8
    myArray(29, 13) = 14.7

    myArray(30, 1) = 2023
    myArray(30, 2) = 28.4
    myArray(30, 3) = 5.4
    myArray(30, 4) = 23.8
    myArray(30, 5) = 54.5
    myArray(30, 6) = 192.9
    myArray(30, 7) = 147.5
    myArray(30, 8) = 776.3
    myArray(30, 9) = 326.9
    myArray(30, 10) = 310.2
    myArray(30, 11) = 12.2
    myArray(30, 12) = 40.3
    myArray(30, 13) = 122.3

    data_DAEJEON = myArray

End Function


Function data_GEUMSAN() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

    myArray(1, 1) = 1994
    myArray(1, 2) = 18.7
    myArray(1, 3) = 12.5
    myArray(1, 4) = 35
    myArray(1, 5) = 22.8
    myArray(1, 6) = 115
    myArray(1, 7) = 89
    myArray(1, 8) = 118.9
    myArray(1, 9) = 196
    myArray(1, 10) = 20
    myArray(1, 11) = 99
    myArray(1, 12) = 20.5
    myArray(1, 13) = 21.1

    myArray(2, 1) = 1995
    myArray(2, 2) = 23.2
    myArray(2, 3) = 17.1
    myArray(2, 4) = 46.9
    myArray(2, 5) = 65.5
    myArray(2, 6) = 35.5
    myArray(2, 7) = 54
    myArray(2, 8) = 83.5
    myArray(2, 9) = 579.5
    myArray(2, 10) = 47.5
    myArray(2, 11) = 23.5
    myArray(2, 12) = 31
    myArray(2, 13) = 4.6

    myArray(3, 1) = 1996
    myArray(3, 2) = 25.4
    myArray(3, 3) = 2.9
    myArray(3, 4) = 123
    myArray(3, 5) = 42.5
    myArray(3, 6) = 37.5
    myArray(3, 7) = 546
    myArray(3, 8) = 174
    myArray(3, 9) = 130
    myArray(3, 10) = 12.5
    myArray(3, 11) = 75.5
    myArray(3, 12) = 89.8
    myArray(3, 13) = 43

    myArray(4, 1) = 1997
    myArray(4, 2) = 21.3
    myArray(4, 3) = 48.2
    myArray(4, 4) = 34
    myArray(4, 5) = 58
    myArray(4, 6) = 170.5
    myArray(4, 7) = 238.5
    myArray(4, 8) = 444.5
    myArray(4, 9) = 246.5
    myArray(4, 10) = 89
    myArray(4, 11) = 9
    myArray(4, 12) = 160
    myArray(4, 13) = 49

    myArray(5, 1) = 1998
    myArray(5, 2) = 38.4
    myArray(5, 3) = 53.9
    myArray(5, 4) = 25.6
    myArray(5, 5) = 177.5
    myArray(5, 6) = 98.5
    myArray(5, 7) = 278.5
    myArray(5, 8) = 184
    myArray(5, 9) = 520
    myArray(5, 10) = 237.3
    myArray(5, 11) = 49
    myArray(5, 12) = 46.1
    myArray(5, 13) = 6.8

    myArray(6, 1) = 1999
    myArray(6, 2) = 5.3
    myArray(6, 3) = 22.9
    myArray(6, 4) = 73
    myArray(6, 5) = 91.5
    myArray(6, 6) = 117.5
    myArray(6, 7) = 198
    myArray(6, 8) = 114.5
    myArray(6, 9) = 167.5
    myArray(6, 10) = 289.5
    myArray(6, 11) = 125
    myArray(6, 12) = 16.4
    myArray(6, 13) = 10.3

    myArray(7, 1) = 2000
    myArray(7, 2) = 36.2
    myArray(7, 3) = 2.9
    myArray(7, 4) = 24.5
    myArray(7, 5) = 73.7
    myArray(7, 6) = 29
    myArray(7, 7) = 244.5
    myArray(7, 8) = 344
    myArray(7, 9) = 372
    myArray(7, 10) = 223
    myArray(7, 11) = 34.5
    myArray(7, 12) = 42
    myArray(7, 13) = 6.5

    myArray(8, 1) = 2001
    myArray(8, 2) = 63.2
    myArray(8, 3) = 76.5
    myArray(8, 4) = 17
    myArray(8, 5) = 22.5
    myArray(8, 6) = 22.5
    myArray(8, 7) = 212.5
    myArray(8, 8) = 203
    myArray(8, 9) = 43
    myArray(8, 10) = 87
    myArray(8, 11) = 96
    myArray(8, 12) = 12
    myArray(8, 13) = 24.1

    myArray(9, 1) = 2002
    myArray(9, 2) = 71.5
    myArray(9, 3) = 7.7
    myArray(9, 4) = 52
    myArray(9, 5) = 149.5
    myArray(9, 6) = 127.5
    myArray(9, 7) = 57
    myArray(9, 8) = 139.5
    myArray(9, 9) = 551
    myArray(9, 10) = 98.5
    myArray(9, 11) = 55.5
    myArray(9, 12) = 23.2
    myArray(9, 13) = 57.8

    myArray(10, 1) = 2003
    myArray(10, 2) = 22.4
    myArray(10, 3) = 66
    myArray(10, 4) = 44
    myArray(10, 5) = 202.5
    myArray(10, 6) = 164
    myArray(10, 7) = 138
    myArray(10, 8) = 575
    myArray(10, 9) = 280.5
    myArray(10, 10) = 192
    myArray(10, 11) = 22.5
    myArray(10, 12) = 42.5
    myArray(10, 13) = 17

    myArray(11, 1) = 2004
    myArray(11, 2) = 11.2
    myArray(11, 3) = 27.3
    myArray(11, 4) = 33
    myArray(11, 5) = 75.5
    myArray(11, 6) = 90.5
    myArray(11, 7) = 323.5
    myArray(11, 8) = 406
    myArray(11, 9) = 330.5
    myArray(11, 10) = 126
    myArray(11, 11) = 2.5
    myArray(11, 12) = 43
    myArray(11, 13) = 34.5

    myArray(12, 1) = 2005
    myArray(12, 2) = 9.4
    myArray(12, 3) = 34
    myArray(12, 4) = 51
    myArray(12, 5) = 31.5
    myArray(12, 6) = 65.5
    myArray(12, 7) = 191
    myArray(12, 8) = 411.5
    myArray(12, 9) = 387
    myArray(12, 10) = 118
    myArray(12, 11) = 23
    myArray(12, 12) = 30.5
    myArray(12, 13) = 22.6

    myArray(13, 1) = 2006
    myArray(13, 2) = 28
    myArray(13, 3) = 41.1
    myArray(13, 4) = 8.4
    myArray(13, 5) = 112
    myArray(13, 6) = 93.5
    myArray(13, 7) = 73
    myArray(13, 8) = 681.5
    myArray(13, 9) = 118
    myArray(13, 10) = 40.5
    myArray(13, 11) = 54
    myArray(13, 12) = 71
    myArray(13, 13) = 28.9

    myArray(14, 1) = 2007
    myArray(14, 2) = 13.7
    myArray(14, 3) = 57
    myArray(14, 4) = 129
    myArray(14, 5) = 27.5
    myArray(14, 6) = 104
    myArray(14, 7) = 180
    myArray(14, 8) = 252
    myArray(14, 9) = 343.5
    myArray(14, 10) = 398.5
    myArray(14, 11) = 32
    myArray(14, 12) = 13.5
    myArray(14, 13) = 35.4

    myArray(15, 1) = 2008
    myArray(15, 2) = 32.4
    myArray(15, 3) = 6.1
    myArray(15, 4) = 28.3
    myArray(15, 5) = 37.6
    myArray(15, 6) = 84.5
    myArray(15, 7) = 190.5
    myArray(15, 8) = 202
    myArray(15, 9) = 210
    myArray(15, 10) = 35.9
    myArray(15, 11) = 40.1
    myArray(15, 12) = 15.1
    myArray(15, 13) = 19.7

    myArray(16, 1) = 2009
    myArray(16, 2) = 13.2
    myArray(16, 3) = 41.5
    myArray(16, 4) = 43
    myArray(16, 5) = 36
    myArray(16, 6) = 120.3
    myArray(16, 7) = 116.3
    myArray(16, 8) = 515.5
    myArray(16, 9) = 97
    myArray(16, 10) = 54.5
    myArray(16, 11) = 24
    myArray(16, 12) = 29
    myArray(16, 13) = 38.3

    myArray(17, 1) = 2010
    myArray(17, 2) = 33.6
    myArray(17, 3) = 74.5
    myArray(17, 4) = 83.8
    myArray(17, 5) = 73.7
    myArray(17, 6) = 114.5
    myArray(17, 7) = 62.5
    myArray(17, 8) = 278.5
    myArray(17, 9) = 495.6
    myArray(17, 10) = 110.3
    myArray(17, 11) = 20.2
    myArray(17, 12) = 20
    myArray(17, 13) = 36.5

    myArray(18, 1) = 2011
    myArray(18, 2) = 2.2
    myArray(18, 3) = 63.5
    myArray(18, 4) = 21.5
    myArray(18, 5) = 132.9
    myArray(18, 6) = 130.6
    myArray(18, 7) = 237.8
    myArray(18, 8) = 571.2
    myArray(18, 9) = 403
    myArray(18, 10) = 77.8
    myArray(18, 11) = 52.2
    myArray(18, 12) = 98
    myArray(18, 13) = 7.8

    myArray(19, 1) = 2012
    myArray(19, 2) = 23.7
    myArray(19, 3) = 1.1
    myArray(19, 4) = 83.6
    myArray(19, 5) = 75.9
    myArray(19, 6) = 21.7
    myArray(19, 7) = 115.7
    myArray(19, 8) = 239.2
    myArray(19, 9) = 497.5
    myArray(19, 10) = 219.5
    myArray(19, 11) = 46.6
    myArray(19, 12) = 47.3
    myArray(19, 13) = 62.7

    myArray(20, 1) = 2013
    myArray(20, 2) = 37
    myArray(20, 3) = 43.8
    myArray(20, 4) = 64.6
    myArray(20, 5) = 86.4
    myArray(20, 6) = 79.5
    myArray(20, 7) = 117.7
    myArray(20, 8) = 216.9
    myArray(20, 9) = 159.5
    myArray(20, 10) = 80.8
    myArray(20, 11) = 32.6
    myArray(20, 12) = 53.9
    myArray(20, 13) = 24.1

    myArray(21, 1) = 2014
    myArray(21, 2) = 4.1
    myArray(21, 3) = 2.7
    myArray(21, 4) = 97.9
    myArray(21, 5) = 88.7
    myArray(21, 6) = 26
    myArray(21, 7) = 45.6
    myArray(21, 8) = 105.8
    myArray(21, 9) = 426.4
    myArray(21, 10) = 91.2
    myArray(21, 11) = 141.2
    myArray(21, 12) = 70.1
    myArray(21, 13) = 31.3

    myArray(22, 1) = 2015
    myArray(22, 2) = 37.6
    myArray(22, 3) = 23.4
    myArray(22, 4) = 46.6
    myArray(22, 5) = 93.5
    myArray(22, 6) = 29.5
    myArray(22, 7) = 143.7
    myArray(22, 8) = 162.3
    myArray(22, 9) = 83.6
    myArray(22, 10) = 18.6
    myArray(22, 11) = 93.5
    myArray(22, 12) = 109.6
    myArray(22, 13) = 35.7

    myArray(23, 1) = 2016
    myArray(23, 2) = 11.1
    myArray(23, 3) = 46
    myArray(23, 4) = 54.5
    myArray(23, 5) = 171.7
    myArray(23, 6) = 70.5
    myArray(23, 7) = 87.4
    myArray(23, 8) = 377.9
    myArray(23, 9) = 105.6
    myArray(23, 10) = 160.9
    myArray(23, 11) = 157.2
    myArray(23, 12) = 33.2
    myArray(23, 13) = 49.6

    myArray(24, 1) = 2017
    myArray(24, 2) = 13.6
    myArray(24, 3) = 54.6
    myArray(24, 4) = 29.8
    myArray(24, 5) = 76.1
    myArray(24, 6) = 31.8
    myArray(24, 7) = 48.3
    myArray(24, 8) = 305.5
    myArray(24, 9) = 222.3
    myArray(24, 10) = 105.6
    myArray(24, 11) = 35.1
    myArray(24, 12) = 15.6
    myArray(24, 13) = 29.3

    myArray(25, 1) = 2018
    myArray(25, 2) = 25.7
    myArray(25, 3) = 28.1
    myArray(25, 4) = 91.5
    myArray(25, 5) = 142.4
    myArray(25, 6) = 110.4
    myArray(25, 7) = 104.3
    myArray(25, 8) = 163.5
    myArray(25, 9) = 410.4
    myArray(25, 10) = 135.2
    myArray(25, 11) = 112.6
    myArray(25, 12) = 45.5
    myArray(25, 13) = 27.6

    myArray(26, 1) = 2019
    myArray(26, 2) = 6.4
    myArray(26, 3) = 41.5
    myArray(26, 4) = 33
    myArray(26, 5) = 93
    myArray(26, 6) = 44.2
    myArray(26, 7) = 101
    myArray(26, 8) = 141.1
    myArray(26, 9) = 105.8
    myArray(26, 10) = 236.4
    myArray(26, 11) = 99.3
    myArray(26, 12) = 47.9
    myArray(26, 13) = 33

    myArray(27, 1) = 2020
    myArray(27, 2) = 80.8
    myArray(27, 3) = 83.9
    myArray(27, 4) = 20.5
    myArray(27, 5) = 35.6
    myArray(27, 6) = 80.5
    myArray(27, 7) = 234
    myArray(27, 8) = 628
    myArray(27, 9) = 373.4
    myArray(27, 10) = 167.2
    myArray(27, 11) = 4.1
    myArray(27, 12) = 41.9
    myArray(27, 13) = 8.3

    myArray(28, 1) = 2021
    myArray(28, 2) = 23.5
    myArray(28, 3) = 19.3
    myArray(28, 4) = 88
    myArray(28, 5) = 39.3
    myArray(28, 6) = 162.7
    myArray(28, 7) = 105.6
    myArray(28, 8) = 300.8
    myArray(28, 9) = 297.2
    myArray(28, 10) = 151.9
    myArray(28, 11) = 44
    myArray(28, 12) = 50.7
    myArray(28, 13) = 7.1

    myArray(29, 1) = 2022
    myArray(29, 2) = 1.5
    myArray(29, 3) = 4.1
    myArray(29, 4) = 80.6
    myArray(29, 5) = 63.3
    myArray(29, 6) = 4.7
    myArray(29, 7) = 145.4
    myArray(29, 8) = 183.7
    myArray(29, 9) = 265.7
    myArray(29, 10) = 68.2
    myArray(29, 11) = 59.3
    myArray(29, 12) = 54.2
    myArray(29, 13) = 18

    myArray(30, 1) = 2023
    myArray(30, 2) = 28.7
    myArray(30, 3) = 8.7
    myArray(30, 4) = 22
    myArray(30, 5) = 46.1
    myArray(30, 6) = 211.5
    myArray(30, 7) = 196.9
    myArray(30, 8) = 624.5
    myArray(30, 9) = 248.7
    myArray(30, 10) = 218.5
    myArray(30, 11) = 7.6
    myArray(30, 12) = 60
    myArray(30, 13) = 121.6

    data_GEUMSAN = myArray

End Function

Function data_SEOSAN() As Variant

    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

    myArray(1, 1) = 1994
    myArray(1, 2) = 14.6
    myArray(1, 3) = 5.9
    myArray(1, 4) = 65.6
    myArray(1, 5) = 32.4
    myArray(1, 6) = 156
    myArray(1, 7) = 167.8
    myArray(1, 8) = 107.1
    myArray(1, 9) = 309.7
    myArray(1, 10) = 99.2
    myArray(1, 11) = 216.3
    myArray(1, 12) = 23.3
    myArray(1, 13) = 36.6

    myArray(2, 1) = 1995
    myArray(2, 2) = 22.7
    myArray(2, 3) = 7.2
    myArray(2, 4) = 37.3
    myArray(2, 5) = 48.2
    myArray(2, 6) = 67.1
    myArray(2, 7) = 24.5
    myArray(2, 8) = 144.1
    myArray(2, 9) = 992.7
    myArray(2, 10) = 20.2
    myArray(2, 11) = 19.3
    myArray(2, 12) = 49.9
    myArray(2, 13) = 15.1

    myArray(3, 1) = 1996
    myArray(3, 2) = 29.1
    myArray(3, 3) = 5.7
    myArray(3, 4) = 115.1
    myArray(3, 5) = 48.1
    myArray(3, 6) = 20
    myArray(3, 7) = 179.2
    myArray(3, 8) = 152.8
    myArray(3, 9) = 74.1
    myArray(3, 10) = 6.4
    myArray(3, 11) = 92.2
    myArray(3, 12) = 72.1
    myArray(3, 13) = 35.3

    myArray(4, 1) = 1997
    myArray(4, 2) = 20.5
    myArray(4, 3) = 32.5
    myArray(4, 4) = 29.6
    myArray(4, 5) = 69.5
    myArray(4, 6) = 232.8
    myArray(4, 7) = 204.4
    myArray(4, 8) = 298.7
    myArray(4, 9) = 87.2
    myArray(4, 10) = 16.1
    myArray(4, 11) = 8.7
    myArray(4, 12) = 116.7
    myArray(4, 13) = 40.2

    myArray(5, 1) = 1998
    myArray(5, 2) = 40.1
    myArray(5, 3) = 54.2
    myArray(5, 4) = 35
    myArray(5, 5) = 160.6
    myArray(5, 6) = 95.5
    myArray(5, 7) = 281.7
    myArray(5, 8) = 295.6
    myArray(5, 9) = 491.8
    myArray(5, 10) = 168
    myArray(5, 11) = 24.3
    myArray(5, 12) = 55.6
    myArray(5, 13) = 9.2

    myArray(6, 1) = 1999
    myArray(6, 2) = 8
    myArray(6, 3) = 7.8
    myArray(6, 4) = 59.9
    myArray(6, 5) = 90.1
    myArray(6, 6) = 178.8
    myArray(6, 7) = 105.1
    myArray(6, 8) = 175.6
    myArray(6, 9) = 497.4
    myArray(6, 10) = 532.6
    myArray(6, 11) = 111.3
    myArray(6, 12) = 36.6
    myArray(6, 13) = 23.4

    myArray(7, 1) = 2000
    myArray(7, 2) = 63
    myArray(7, 3) = 2.9
    myArray(7, 4) = 3.7
    myArray(7, 5) = 38.1
    myArray(7, 6) = 62.1
    myArray(7, 7) = 204.4
    myArray(7, 8) = 60.8
    myArray(7, 9) = 608.1
    myArray(7, 10) = 298.1
    myArray(7, 11) = 34.4
    myArray(7, 12) = 24.8
    myArray(7, 13) = 24.4

    myArray(8, 1) = 2001
    myArray(8, 2) = 66.9
    myArray(8, 3) = 40.4
    myArray(8, 4) = 12.7
    myArray(8, 5) = 18.7
    myArray(8, 6) = 17.8
    myArray(8, 7) = 200.2
    myArray(8, 8) = 402
    myArray(8, 9) = 136.6
    myArray(8, 10) = 15
    myArray(8, 11) = 47.5
    myArray(8, 12) = 8.2
    myArray(8, 13) = 20.8

    myArray(9, 1) = 2002
    myArray(9, 2) = 22.5
    myArray(9, 3) = 7
    myArray(9, 4) = 29.3
    myArray(9, 5) = 179.5
    myArray(9, 6) = 177.3
    myArray(9, 7) = 60.8
    myArray(9, 8) = 296.1
    myArray(9, 9) = 428.2
    myArray(9, 10) = 57.5
    myArray(9, 11) = 78.3
    myArray(9, 12) = 36.3
    myArray(9, 13) = 14.8

    myArray(10, 1) = 2003
    myArray(10, 2) = 20.9
    myArray(10, 3) = 39
    myArray(10, 4) = 22.5
    myArray(10, 5) = 180
    myArray(10, 6) = 105.5
    myArray(10, 7) = 221.8
    myArray(10, 8) = 290.2
    myArray(10, 9) = 257.9
    myArray(10, 10) = 201.9
    myArray(10, 11) = 23
    myArray(10, 12) = 53.6
    myArray(10, 13) = 17.1

    myArray(11, 1) = 2004
    myArray(11, 2) = 27.3
    myArray(11, 3) = 26.3
    myArray(11, 4) = 15.7
    myArray(11, 5) = 80.2
    myArray(11, 6) = 140.3
    myArray(11, 7) = 211.1
    myArray(11, 8) = 321.9
    myArray(11, 9) = 131.2
    myArray(11, 10) = 282.6
    myArray(11, 11) = 1.8
    myArray(11, 12) = 70.5
    myArray(11, 13) = 32

    myArray(12, 1) = 2005
    myArray(12, 2) = 10.4
    myArray(12, 3) = 34
    myArray(12, 4) = 36.1
    myArray(12, 5) = 77.2
    myArray(12, 6) = 56.1
    myArray(12, 7) = 147
    myArray(12, 8) = 386.1
    myArray(12, 9) = 270.5
    myArray(12, 10) = 228.7
    myArray(12, 11) = 30.9
    myArray(12, 12) = 19.6
    myArray(12, 13) = 37.6

    myArray(13, 1) = 2006
    myArray(13, 2) = 29.7
    myArray(13, 3) = 18.9
    myArray(13, 4) = 5
    myArray(13, 5) = 77.3
    myArray(13, 6) = 133.5
    myArray(13, 7) = 226.8
    myArray(13, 8) = 494.5
    myArray(13, 9) = 58.2
    myArray(13, 10) = 10.1
    myArray(13, 11) = 10.5
    myArray(13, 12) = 55
    myArray(13, 13) = 19.7

    myArray(14, 1) = 2007
    myArray(14, 2) = 13
    myArray(14, 3) = 25.5
    myArray(14, 4) = 127.2
    myArray(14, 5) = 28.1
    myArray(14, 6) = 108.5
    myArray(14, 7) = 123.5
    myArray(14, 8) = 257
    myArray(14, 9) = 414.6
    myArray(14, 10) = 305.8
    myArray(14, 11) = 30.7
    myArray(14, 12) = 14.4
    myArray(14, 13) = 22.8

    myArray(15, 1) = 2008
    myArray(15, 2) = 15
    myArray(15, 3) = 7
    myArray(15, 4) = 26
    myArray(15, 5) = 46.1
    myArray(15, 6) = 88.5
    myArray(15, 7) = 118.1
    myArray(15, 8) = 335.5
    myArray(15, 9) = 114.2
    myArray(15, 10) = 62.7
    myArray(15, 11) = 34
    myArray(15, 12) = 34.6
    myArray(15, 13) = 27.9

    myArray(16, 1) = 2009
    myArray(16, 2) = 15.2
    myArray(16, 3) = 26.5
    myArray(16, 4) = 67
    myArray(16, 5) = 43
    myArray(16, 6) = 117.9
    myArray(16, 7) = 74.9
    myArray(16, 8) = 364.9
    myArray(16, 9) = 196.3
    myArray(16, 10) = 16
    myArray(16, 11) = 49.2
    myArray(16, 12) = 59.1
    myArray(16, 13) = 44.3

    myArray(17, 1) = 2010
    myArray(17, 2) = 55.5
    myArray(17, 3) = 58.4
    myArray(17, 4) = 79.2
    myArray(17, 5) = 52.2
    myArray(17, 6) = 168
    myArray(17, 7) = 94.9
    myArray(17, 8) = 447.1
    myArray(17, 9) = 707
    myArray(17, 10) = 402
    myArray(17, 11) = 29.1
    myArray(17, 12) = 12
    myArray(17, 13) = 36.4

    myArray(18, 1) = 2011
    myArray(18, 2) = 8.8
    myArray(18, 3) = 55.8
    myArray(18, 4) = 34.5
    myArray(18, 5) = 96.2
    myArray(18, 6) = 107.9
    myArray(18, 7) = 462.6
    myArray(18, 8) = 656.5
    myArray(18, 9) = 151.2
    myArray(18, 10) = 50.3
    myArray(18, 11) = 18.1
    myArray(18, 12) = 48.9
    myArray(18, 13) = 13.6

    myArray(19, 1) = 2012
    myArray(19, 2) = 15.1
    myArray(19, 3) = 2.4
    myArray(19, 4) = 41.6
    myArray(19, 5) = 113.5
    myArray(19, 6) = 14.5
    myArray(19, 7) = 91.1
    myArray(19, 8) = 266.8
    myArray(19, 9) = 647.9
    myArray(19, 10) = 201.5
    myArray(19, 11) = 100.7
    myArray(19, 12) = 82.1
    myArray(19, 13) = 65.4

    myArray(20, 1) = 2013
    myArray(20, 2) = 36.8
    myArray(20, 3) = 64.8
    myArray(20, 4) = 60.8
    myArray(20, 5) = 61.8
    myArray(20, 6) = 114.9
    myArray(20, 7) = 94.4
    myArray(20, 8) = 213.8
    myArray(20, 9) = 120.6
    myArray(20, 10) = 147.4
    myArray(20, 11) = 5.7
    myArray(20, 12) = 64.9
    myArray(20, 13) = 32.8

    myArray(21, 1) = 2014
    myArray(21, 2) = 7
    myArray(21, 3) = 17
    myArray(21, 4) = 31.2
    myArray(21, 5) = 85.6
    myArray(21, 6) = 52.7
    myArray(21, 7) = 69.3
    myArray(21, 8) = 151.7
    myArray(21, 9) = 242.3
    myArray(21, 10) = 106.7
    myArray(21, 11) = 117.2
    myArray(21, 12) = 37.8
    myArray(21, 13) = 81.6

    myArray(22, 1) = 2015
    myArray(22, 2) = 20.7
    myArray(22, 3) = 23.1
    myArray(22, 4) = 20.6
    myArray(22, 5) = 116.8
    myArray(22, 6) = 40.6
    myArray(22, 7) = 64.1
    myArray(22, 8) = 158.5
    myArray(22, 9) = 63.1
    myArray(22, 10) = 15.1
    myArray(22, 11) = 73.1
    myArray(22, 12) = 156.6
    myArray(22, 13) = 63.6

    myArray(23, 1) = 2016
    myArray(23, 2) = 21.9
    myArray(23, 3) = 61.7
    myArray(23, 4) = 24.3
    myArray(23, 5) = 87
    myArray(23, 6) = 153.7
    myArray(23, 7) = 36.8
    myArray(23, 8) = 295.6
    myArray(23, 9) = 34
    myArray(23, 10) = 53.1
    myArray(23, 11) = 73.8
    myArray(23, 12) = 17.5
    myArray(23, 13) = 62.7

    myArray(24, 1) = 2017
    myArray(24, 2) = 21.3
    myArray(24, 3) = 31.4
    myArray(24, 4) = 4.8
    myArray(24, 5) = 38.9
    myArray(24, 6) = 27.9
    myArray(24, 7) = 23.3
    myArray(24, 8) = 327.8
    myArray(24, 9) = 231.3
    myArray(24, 10) = 37.6
    myArray(24, 11) = 25.5
    myArray(24, 12) = 24.7
    myArray(24, 13) = 35.9

    myArray(25, 1) = 2018
    myArray(25, 2) = 21
    myArray(25, 3) = 40.5
    myArray(25, 4) = 76.6
    myArray(25, 5) = 132.8
    myArray(25, 6) = 147.7
    myArray(25, 7) = 162.3
    myArray(25, 8) = 152.9
    myArray(25, 9) = 156.8
    myArray(25, 10) = 82.7
    myArray(25, 11) = 153.2
    myArray(25, 12) = 73.9
    myArray(25, 13) = 26.8

    myArray(26, 1) = 2019
    myArray(26, 2) = 1.1
    myArray(26, 3) = 30.2
    myArray(26, 4) = 43.7
    myArray(26, 5) = 43.9
    myArray(26, 6) = 20.1
    myArray(26, 7) = 56.3
    myArray(26, 8) = 174.5
    myArray(26, 9) = 121.1
    myArray(26, 10) = 181.1
    myArray(26, 11) = 81
    myArray(26, 12) = 124.6
    myArray(26, 13) = 37.4

    myArray(27, 1) = 2020
    myArray(27, 2) = 46
    myArray(27, 3) = 72.3
    myArray(27, 4) = 23
    myArray(27, 5) = 20.7
    myArray(27, 6) = 101.3
    myArray(27, 7) = 144
    myArray(27, 8) = 329.4
    myArray(27, 9) = 400
    myArray(27, 10) = 257.7
    myArray(27, 11) = 12.6
    myArray(27, 12) = 72
    myArray(27, 13) = 9.7

    myArray(28, 1) = 2021
    myArray(28, 2) = 25.3
    myArray(28, 3) = 9.6
    myArray(28, 4) = 112.8
    myArray(28, 5) = 110.6
    myArray(28, 6) = 132.3
    myArray(28, 7) = 70.9
    myArray(28, 8) = 121.6
    myArray(28, 9) = 217.8
    myArray(28, 10) = 206
    myArray(28, 11) = 55.9
    myArray(28, 12) = 126.2
    myArray(28, 13) = 18.3

    myArray(29, 1) = 2022
    myArray(29, 2) = 8.6
    myArray(29, 3) = 4.7
    myArray(29, 4) = 72.1
    myArray(29, 5) = 52.2
    myArray(29, 6) = 2.9
    myArray(29, 7) = 352.4
    myArray(29, 8) = 178.4
    myArray(29, 9) = 468.7
    myArray(29, 10) = 165.9
    myArray(29, 11) = 160
    myArray(29, 12) = 72.9
    myArray(29, 13) = 31.9

    myArray(30, 1) = 2023
    myArray(30, 2) = 30.5
    myArray(30, 3) = 0.1
    myArray(30, 4) = 6.4
    myArray(30, 5) = 54.6
    myArray(30, 6) = 132.9
    myArray(30, 7) = 138.1
    myArray(30, 8) = 507
    myArray(30, 9) = 225
    myArray(30, 10) = 166.1
    myArray(30, 11) = 39.6
    myArray(30, 12) = 122.9
    myArray(30, 13) = 103.5

    data_SEOSAN = myArray

End Function








Option Explicit

Option Explicit

Function data_SEOUL() As Variant
    Dim myArray() As Variant
    ReDim myArray(1 To 30, 1 To 13)

    myArray(1, 1) = 1994
    myArray(1, 2) = 6.5
    myArray(1, 3) = 14.8
    myArray(1, 4) = 31.7
    myArray(1, 5) = 44.9
    myArray(1, 6) = 152.4
    myArray(1, 7) = 85
    myArray(1, 8) = 139.5
    myArray(1, 9) = 232.7
    myArray(1, 10) = 60.7
    myArray(1, 11) = 214.5
    myArray(1, 12) = 49.6
    myArray(1, 13) = 23.5

    myArray(2, 1) = 1995
    myArray(2, 2) = 11.6
    myArray(2, 3) = 5.2
    myArray(2, 4) = 60.6
    myArray(2, 5) = 44.4
    myArray(2, 6) = 60.6
    myArray(2, 7) = 70.7
    myArray(2, 8) = 436.1
    myArray(2, 9) = 786.6
    myArray(2, 10) = 47.2
    myArray(2, 11) = 39.3
    myArray(2, 12) = 32.9
    myArray(2, 13) = 3.4

    myArray(3, 1) = 1996
    myArray(3, 2) = 16.3
    myArray(3, 3) = 1
    myArray(3, 4) = 77.9
    myArray(3, 5) = 62
    myArray(3, 6) = 29.3
    myArray(3, 7) = 249.7
    myArray(3, 8) = 512.8
    myArray(3, 9) = 132.4
    myArray(3, 10) = 11
    myArray(3, 11) = 90.3
    myArray(3, 12) = 62.9
    myArray(3, 13) = 11

    myArray(4, 1) = 1997
    myArray(4, 2) = 16.8
    myArray(4, 3) = 39.6
    myArray(4, 4) = 25.3
    myArray(4, 5) = 56.1
    myArray(4, 6) = 291.3
    myArray(4, 7) = 110
    myArray(4, 8) = 299.6
    myArray(4, 9) = 117.2
    myArray(4, 10) = 76.9
    myArray(4, 11) = 45.5
    myArray(4, 12) = 93.8
    myArray(4, 13) = 38.1

    myArray(5, 1) = 1998
    myArray(5, 2) = 10.4
    myArray(5, 3) = 32.3
    myArray(5, 4) = 45.1
    myArray(5, 5) = 120.2
    myArray(5, 6) = 121.5
    myArray(5, 7) = 234.1
    myArray(5, 8) = 311.8
    myArray(5, 9) = 1237.8
    myArray(5, 10) = 177.9
    myArray(5, 11) = 27.4
    myArray(5, 12) = 26.9
    myArray(5, 13) = 3.7

    myArray(6, 1) = 1999
    myArray(6, 2) = 10.2
    myArray(6, 3) = 2.9
    myArray(6, 4) = 55
    myArray(6, 5) = 97.2
    myArray(6, 6) = 109.7
    myArray(6, 7) = 131.8
    myArray(6, 8) = 230.4
    myArray(6, 9) = 600.5
    myArray(6, 10) = 377.3
    myArray(6, 11) = 81.6
    myArray(6, 12) = 19.5
    myArray(6, 13) = 17

    myArray(7, 1) = 2000
    myArray(7, 2) = 42.8
    myArray(7, 3) = 2.1
    myArray(7, 4) = 3.1
    myArray(7, 5) = 30.7
    myArray(7, 6) = 75.2
    myArray(7, 7) = 68.1
    myArray(7, 8) = 114.7
    myArray(7, 9) = 599.4
    myArray(7, 10) = 178.5
    myArray(7, 11) = 18.1
    myArray(7, 12) = 27.1
    myArray(7, 13) = 27

    myArray(8, 1) = 2001
    myArray(8, 2) = 39.4
    myArray(8, 3) = 45.7
    myArray(8, 4) = 18.1
    myArray(8, 5) = 12.3
    myArray(8, 6) = 16.5
    myArray(8, 7) = 157.4
    myArray(8, 8) = 698.4
    myArray(8, 9) = 252
    myArray(8, 10) = 49.3
    myArray(8, 11) = 68.2
    myArray(8, 12) = 13
    myArray(8, 13) = 15.7

    myArray(9, 1) = 2002
    myArray(9, 2) = 37.4
    myArray(9, 3) = 2.4
    myArray(9, 4) = 31.5
    myArray(9, 5) = 155.1
    myArray(9, 6) = 58
    myArray(9, 7) = 61.4
    myArray(9, 8) = 220.6
    myArray(9, 9) = 688
    myArray(9, 10) = 61.1
    myArray(9, 11) = 45
    myArray(9, 12) = 12.5
    myArray(9, 13) = 15

    myArray(10, 1) = 2003
    myArray(10, 2) = 14.1
    myArray(10, 3) = 39.6
    myArray(10, 4) = 26.8
    myArray(10, 5) = 139.6
    myArray(10, 6) = 106
    myArray(10, 7) = 156
    myArray(10, 8) = 469.8
    myArray(10, 9) = 684.2
    myArray(10, 10) = 258.2
    myArray(10, 11) = 41.5
    myArray(10, 12) = 69.3
    myArray(10, 13) = 6.9

    myArray(11, 1) = 2004
    myArray(11, 2) = 19.8
    myArray(11, 3) = 54.6
    myArray(11, 4) = 27.6
    myArray(11, 5) = 74.1
    myArray(11, 6) = 168.5
    myArray(11, 7) = 138.1
    myArray(11, 8) = 510.7
    myArray(11, 9) = 193.3
    myArray(11, 10) = 198.7
    myArray(11, 11) = 6.5
    myArray(11, 12) = 80
    myArray(11, 13) = 27.2

    myArray(12, 1) = 2005
    myArray(12, 2) = 4.5
    myArray(12, 3) = 17.2
    myArray(12, 4) = 12.5
    myArray(12, 5) = 94.7
    myArray(12, 6) = 85.8
    myArray(12, 7) = 168.5
    myArray(12, 8) = 269.4
    myArray(12, 9) = 285
    myArray(12, 10) = 313.3
    myArray(12, 11) = 52.6
    myArray(12, 12) = 44.6
    myArray(12, 13) = 10.3

    myArray(13, 1) = 2006
    myArray(13, 2) = 34.3
    myArray(13, 3) = 15.7
    myArray(13, 4) = 14
    myArray(13, 5) = 51.8
    myArray(13, 6) = 156.2
    myArray(13, 7) = 168.5
    myArray(13, 8) = 1014
    myArray(13, 9) = 121.2
    myArray(13, 10) = 11.1
    myArray(13, 11) = 30.2
    myArray(13, 12) = 47.6
    myArray(13, 13) = 17.3

    myArray(14, 1) = 2007
    myArray(14, 2) = 10.8
    myArray(14, 3) = 12.6
    myArray(14, 4) = 123.5
    myArray(14, 5) = 41.1
    myArray(14, 6) = 137.6
    myArray(14, 7) = 54.5
    myArray(14, 8) = 274.1
    myArray(14, 9) = 237.6
    myArray(14, 10) = 241.9
    myArray(14, 11) = 39.5
    myArray(14, 12) = 26.4
    myArray(14, 13) = 12.7

    myArray(15, 1) = 2008
    myArray(15, 2) = 17.7
    myArray(15, 3) = 15
    myArray(15, 4) = 53.9
    myArray(15, 5) = 38.5
    myArray(15, 6) = 97.7
    myArray(15, 7) = 165
    myArray(15, 8) = 530.8
    myArray(15, 9) = 251.2
    myArray(15, 10) = 99.2
    myArray(15, 11) = 41.8
    myArray(15, 12) = 19.6
    myArray(15, 13) = 25.9

    myArray(16, 1) = 2009
    myArray(16, 2) = 5.7
    myArray(16, 3) = 36.9
    myArray(16, 4) = 63.9
    myArray(16, 5) = 66.5
    myArray(16, 6) = 109
    myArray(16, 7) = 132
    myArray(16, 8) = 659.4
    myArray(16, 9) = 285.3
    myArray(16, 10) = 64.5
    myArray(16, 11) = 66.9
    myArray(16, 12) = 52.4
    myArray(16, 13) = 21.5

    myArray(17, 1) = 2010
    myArray(17, 2) = 29.3
    myArray(17, 3) = 55.3
    myArray(17, 4) = 82.5
    myArray(17, 5) = 62.8
    myArray(17, 6) = 124
    myArray(17, 7) = 127.6
    myArray(17, 8) = 239.2
    myArray(17, 9) = 598.7
    myArray(17, 10) = 671.5
    myArray(17, 11) = 25.6
    myArray(17, 12) = 10.9
    myArray(17, 13) = 16.1

    myArray(18, 1) = 2011
    myArray(18, 2) = 8.9
    myArray(18, 3) = 29.1
    myArray(18, 4) = 14.6
    myArray(18, 5) = 110.1
    myArray(18, 6) = 53.4
    myArray(18, 7) = 404.5
    myArray(18, 8) = 1131
    myArray(18, 9) = 166.8
    myArray(18, 10) = 25.6
    myArray(18, 11) = 32
    myArray(18, 12) = 56.2
    myArray(18, 13) = 7.1

    myArray(19, 1) = 2012
    myArray(19, 2) = 6.7
    myArray(19, 3) = 0.8
    myArray(19, 4) = 47.4
    myArray(19, 5) = 157
    myArray(19, 6) = 8.2
    myArray(19, 7) = 91.9
    myArray(19, 8) = 448.9
    myArray(19, 9) = 464.9
    myArray(19, 10) = 212
    myArray(19, 11) = 99.3
    myArray(19, 12) = 67.8
    myArray(19, 13) = 41.4

    myArray(20, 1) = 2013
    myArray(20, 2) = 22.1
    myArray(20, 3) = 74.1
    myArray(20, 4) = 27.3
    myArray(20, 5) = 71.7
    myArray(20, 6) = 132
    myArray(20, 7) = 28.3
    myArray(20, 8) = 676.2
    myArray(20, 9) = 148.6
    myArray(20, 10) = 138.5
    myArray(20, 11) = 13.5
    myArray(20, 12) = 46.8
    myArray(20, 13) = 24.7

    myArray(21, 1) = 2014
    myArray(21, 2) = 13
    myArray(21, 3) = 16.2
    myArray(21, 4) = 7.2
    myArray(21, 5) = 31
    myArray(21, 6) = 63
    myArray(21, 7) = 98.1
    myArray(21, 8) = 207.9
    myArray(21, 9) = 172.8
    myArray(21, 10) = 88.1
    myArray(21, 11) = 52.2
    myArray(21, 12) = 41.5
    myArray(21, 13) = 17.9

    myArray(22, 1) = 2015
    myArray(22, 2) = 11.3
    myArray(22, 3) = 22.7
    myArray(22, 4) = 9.6
    myArray(22, 5) = 80.5
    myArray(22, 6) = 28.9
    myArray(22, 7) = 99
    myArray(22, 8) = 226
    myArray(22, 9) = 72.9
    myArray(22, 10) = 26
    myArray(22, 11) = 81.5
    myArray(22, 12) = 104.6
    myArray(22, 13) = 29.1

    myArray(23, 1) = 2016
    myArray(23, 2) = 1
    myArray(23, 3) = 47.6
    myArray(23, 4) = 40.5
    myArray(23, 5) = 76.8
    myArray(23, 6) = 160.5
    myArray(23, 7) = 54.4
    myArray(23, 8) = 358.2
    myArray(23, 9) = 67.1
    myArray(23, 10) = 33
    myArray(23, 11) = 74.8
    myArray(23, 12) = 16.7
    myArray(23, 13) = 61.1

    myArray(24, 1) = 2017
    myArray(24, 2) = 14.9
    myArray(24, 3) = 11.1
    myArray(24, 4) = 7.9
    myArray(24, 5) = 61.6
    myArray(24, 6) = 16.1
    myArray(24, 7) = 66.6
    myArray(24, 8) = 621
    myArray(24, 9) = 297
    myArray(24, 10) = 35
    myArray(24, 11) = 26.5
    myArray(24, 12) = 40.7
    myArray(24, 13) = 34.8

    myArray(25, 1) = 2018
    myArray(25, 2) = 8.5
    myArray(25, 3) = 29.6
    myArray(25, 4) = 49.5
    myArray(25, 5) = 130.3
    myArray(25, 6) = 222
    myArray(25, 7) = 171.5
    myArray(25, 8) = 185.6
    myArray(25, 9) = 202.6
    myArray(25, 10) = 68.5
    myArray(25, 11) = 120.5
    myArray(25, 12) = 79.1
    myArray(25, 13) = 16.4

    myArray(26, 1) = 2019
    myArray(26, 2) = 0
    myArray(26, 3) = 23.8
    myArray(26, 4) = 26.8
    myArray(26, 5) = 47.3
    myArray(26, 6) = 37.8
    myArray(26, 7) = 74
    myArray(26, 8) = 194.4
    myArray(26, 9) = 190.5
    myArray(26, 10) = 139.8
    myArray(26, 11) = 55.5
    myArray(26, 12) = 78.8
    myArray(26, 13) = 22.6

    myArray(27, 1) = 2020
    myArray(27, 2) = 60.5
    myArray(27, 3) = 53.1
    myArray(27, 4) = 16.3
    myArray(27, 5) = 16.9
    myArray(27, 6) = 112.4
    myArray(27, 7) = 139.6
    myArray(27, 8) = 270.4
    myArray(27, 9) = 675.7
    myArray(27, 10) = 181.5
    myArray(27, 11) = 0
    myArray(27, 12) = 120.1
    myArray(27, 13) = 4.6

    myArray(28, 1) = 2021
    myArray(28, 2) = 18.9
    myArray(28, 3) = 7.1
    myArray(28, 4) = 110.9
    myArray(28, 5) = 124.1
    myArray(28, 6) = 183.1
    myArray(28, 7) = 104.6
    myArray(28, 8) = 168.3
    myArray(28, 9) = 211.2
    myArray(28, 10) = 131
    myArray(28, 11) = 57
    myArray(28, 12) = 62.4
    myArray(28, 13) = 7.9

    myArray(29, 1) = 2022
    myArray(29, 2) = 5.5
    myArray(29, 3) = 4.7
    myArray(29, 4) = 102.6
    myArray(29, 5) = 20.4
    myArray(29, 6) = 7.5
    myArray(29, 7) = 393.8
    myArray(29, 8) = 252.3
    myArray(29, 9) = 564.8
    myArray(29, 10) = 201.5
    myArray(29, 11) = 124.1
    myArray(29, 12) = 84.5
    myArray(29, 13) = 13.6

    myArray(30, 1) = 2023
    myArray(30, 2) = 47.9
    myArray(30, 3) = 1
    myArray(30, 4) = 10.5
    myArray(30, 5) = 96.9
    myArray(30, 6) = 155.6
    myArray(30, 7) = 195.6
    myArray(30, 8) = 459.9
    myArray(30, 9) = 298.1
    myArray(30, 10) = 134.5
    myArray(30, 11) = 31
    myArray(30, 12) = 81.9
    myArray(30, 13) = 85.9

    data_SEOUL = myArray
End Function






Sub GitSave()
    
    DeleteAndMake
    ExportModules
    PrintAllCode
    PrintAllContainers
    
End Sub

Sub DeleteAndMake()
        
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim parentFolder As String: parentFolder = ThisWorkbook.Path & "\VBA"
    Dim childA As String: childA = parentFolder & "\VBA-Code_Together"
    Dim childB As String: childB = parentFolder & "\VBA-Code_By_Modules"
        
    On Error Resume Next
    fso.DeleteFolder parentFolder
    On Error GoTo 0
    
    MkDir parentFolder
    MkDir childA
    MkDir childB
    
End Sub

Sub PrintAllCode()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.codeModule.Lines(1, item.codeModule.CountOfLines)
        Debug.Print lineToPrint
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_Together\"
    If Dir(pathToExport) <> "" Then Kill pathToExport & "*.*"
    SaveTextToFile textToPrint, pathToExport & "all_code.vb"
    
End Sub

Sub PrintAllContainers()
    
    Dim item  As Variant
    Dim textToPrint As String
    Dim lineToPrint As String
    
    For Each item In ThisWorkbook.VBProject.VBComponents
        lineToPrint = item.Name
        Debug.Print lineToPrint
        textToPrint = textToPrint & vbCrLf & lineToPrint
    Next item
    
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_Together\"
    SaveTextToFile textToPrint, pathToExport & "all_modules.vb"
    
End Sub

Sub ExportModules()
       
    Dim pathToExport As String: pathToExport = ThisWorkbook.Path & "\VBA\VBA-Code_By_Modules\"
    
    If Dir(pathToExport) <> "" Then
        Kill pathToExport & "*.*"
    End If
     
    Dim wkb As Workbook: Set wkb = Excel.Workbooks(ThisWorkbook.Name)
    
    Dim unitsCount As Long
    Dim filePath As String
    Dim component As VBIDE.VBComponent
    Dim tryExport As Boolean

    For Each component In wkb.VBProject.VBComponents
        tryExport = True
        filePath = component.Name
       
        Select Case component.Type
            Case vbext_ct_ClassModule
                filePath = filePath & ".cls"
            Case vbext_ct_MSForm
                filePath = filePath & ".frm"
            Case vbext_ct_StdModule
                filePath = filePath & ".bas"
            Case vbext_ct_Document
                tryExport = False
        End Select
        
        If tryExport Then
            Debug.Print unitsCount & " exporting " & filePath
            component.Export pathToExport & "\" & filePath
        End If
    Next

    Debug.Print "Exported at " & pathToExport
    
End Sub

Sub SaveTextToFile(dataToPrint As String, pathToExport As String)
    
    Dim fileSystem As Object
    Dim textObject As Object
    Dim fileName As String
    Dim newFile  As String
    Dim shellPath  As String
    
    If Dir(ThisWorkbook.Path & newFile, vbDirectory) = vbNullString Then MkDir ThisWorkbook.Path & newFile
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set textObject = fileSystem.CreateTextFile(pathToExport, True)
    
    textObject.WriteLine dataToPrint
    textObject.Close
        
    On Error GoTo 0
    Exit Sub

CreateLogFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CreateLogFile of Sub mod_TDD_Export"

End Sub
