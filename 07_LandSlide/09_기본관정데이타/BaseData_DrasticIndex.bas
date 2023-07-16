Attribute VB_Name = "BaseData_DrasticIndex"
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

