Attribute VB_Name = "BaseData_DrasticIndex"
'기본관정데이타 - 드라스틱인덱스


Dim Dr, Rr As Single


Private Sub TurnOffStuff()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    

End Sub

Private Sub TurnOnStuff()

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    
End Sub

Private Sub ShiftNewYear()

   
    Range("B7:N35").Select
    ActiveWindow.SmallScroll Down:=-9
    Selection.Copy
    
    Range("B6").Select
    ActiveSheet.PasteSpecial Format:=3, Link:=1, DisplayAsIcon:=False, _
                             IconFileName:=False
        
    Range("B35:N35").Select
    Selection.ClearContents
    
    ActiveWindow.SmallScroll Down:=18
    Range("B45:N53").Select
    Selection.Copy
    Range("B44").Select
    ActiveSheet.PasteSpecial Format:=3, Link:=1, DisplayAsIcon:=False, _
                             IconFileName:=False
    Range("B53:N53").Select
    Selection.ClearContents
    
End Sub

'Drastic Index 를 계산 해주기 위한 함수 ...
' 2017/11/21 화요일


' 1, 지하수위에 대한 등급의 계산
Private Function Rating_UnderGroundWater(ByVal water_level As Single) As Integer

    Dim result As Integer

   
    If (water_level < 1.52) Then
        result = 10
    ElseIf (water_level < 4.57) Then
        result = 9
    ElseIf (water_level < 9.14) Then
        result = 7
    ElseIf (water_level < 15.24) Then
        result = 5
    ElseIf (water_level < 22.86) Then
        result = 3
    ElseIf (water_level < 30.48) Then
        result = 2
    Else
        result = 1
    End If
              
                        
    Rating_UnderGroundWater = result
                          
End Function

'2, 강수의 지하함양량
Private Function Rating_NetRecharge(ByVal value As Single) As Integer

    Dim result As Integer
    
    If (value < 5.08) Then
        result = 1
    ElseIf (value < 10.16) Then
        result = 3
    ElseIf (value < 17.78) Then
        result = 6
    ElseIf (value < 25.4) Then
        result = 8
    Else
        result = 9
    End If
              
    Rating_NetRecharge = result

End Function

'3, 대수층

Private Function Rating_AqMedia(ByVal value As String) As Integer
        
    
    If StrComp(value, "Massive Shale") = 0 Then
        Rating_AqMedia = 2
        
    End If
    
    If StrComp(value, "Metamorphic/Igneous") = 0 Then
        Rating_AqMedia = 3
        
    End If
    
    If StrComp(value, "Weathered Metamorphic / Igneous") = 0 Then
        Rating_AqMedia = 4
        
    End If
    
    If StrComp(value, "Glacial Till") = 0 Then
        Rating_AqMedia = 5
        
    End If
    
    If StrComp(value, "Bedded SandStone") = 0 Then
        Rating_AqMedia = 6
        
    End If
    
    If StrComp(value, "Massive Sandstone") = 0 Then
        Rating_AqMedia = 6
        
    End If
    
   
    If StrComp(value, "Massive Limestone") = 0 Then
        Rating_AqMedia = 6
        
    End If
       
    If StrComp(value, "Sand and Gravel") = 0 Then
        Rating_AqMedia = 8
        
    End If
       
    If StrComp(value, "Basalt") = 0 Then
        Rating_AqMedia = 9
        
    End If
       
    If StrComp(value, "Karst Limestone") = 0 Then
        Rating_AqMedia = 10
        
    End If
    
    
    
End Function

'4 토양특성ㅇ에 대한 등급

Private Function Rating_SoilMedia(ByVal value As String) As Integer


    If StrComp(value, "Thin or Absecnt") = 0 Then
        Rating_SoilMedia = 10
        
    End If
    
    If StrComp(value, "Gravel") = 0 Then
        Rating_SoilMedia = 10
        
    End If
    
    If StrComp(value, "Sand") = 0 Then
        Rating_SoilMedia = 9
        
    End If
    
    If StrComp(value, "Peat") = 0 Then
        Rating_SoilMedia = 8
        
    End If
    
    If StrComp(value, "Shringing or Aggregated Clay") = 0 Then
        Rating_SoilMedia = 7
        
    End If
    
    If StrComp(value, "Sandy Loam") = 0 Then
        Rating_SoilMedia = 6
        
    End If
    
   
    If StrComp(value, "Loam") = 0 Then
        Rating_SoilMedia = 5
        
    End If
       
    If StrComp(value, "Silty Loam") = 0 Then
        Rating_SoilMedia = 4
        
    End If
       
    If StrComp(value, "Clay Loam") = 0 Then
        Rating_SoilMedia = 3
        
    End If
       
    If StrComp(value, "Mud") = 0 Then
        Rating_SoilMedia = 2
        
    End If
    
    If StrComp(value, "Nonshrinking and Nonaggregated Clay") = 0 Then
        Rating_SoilMedia = 1
        
    End If
    

End Function

' 5, 지형구배
Private Function Rating_Topo(ByVal value As Single) As Integer
    
    Dim result As Integer
    
    If (value < 2) Then
        result = 10
    ElseIf (value < 6) Then
        result = 9
    ElseIf (value < 12) Then
        result = 5
    ElseIf (value < 18) Then
        result = 3
    Else
        result = 1
    End If
              
    Rating_Topo = result


End Function

'6 비포화대의 영향에 대한 등급 Ir

Private Function Rating_Vadose(ByVal value As String) As Integer

    If StrComp(value, "Confining Layer") = 0 Then
        Rating_Vadose = 1
    End If

    If StrComp(value, "Silt/Clay") = 0 Then
        Rating_Vadose = 3
    End If

    If StrComp(value, "Shale") = 0 Then
        Rating_Vadose = 3
    End If

    If StrComp(value, "Limestone") = 0 Then
        Rating_Vadose = 6
    End If
    
    If StrComp(value, "Sandstone") = 0 Then
        Rating_Vadose = 6
    End If

    If StrComp(value, "Bedded Limestone, Sandstone, Shale") = 0 Then
        Rating_Vadose = 6
    End If

    If StrComp(value, "Sand and Gravel with Significant Silt and Clay") = 0 Then
        Rating_Vadose = 6
    End If

    If StrComp(value, "Metamorphic/Igneous") = 0 Then
        Rating_Vadose = 4
    End If

    If StrComp(value, "Sand and Gravel") = 0 Then
        Rating_Vadose = 8
    End If

    If StrComp(value, "Basalt") = 0 Then
        Rating_Vadose = 9
    End If

    If StrComp(value, "Karst Limestone") = 0 Then
        Rating_Vadose = 10
    End If

End Function

' 7, 대수층의 수리전도도에 대한 등급 : Cr
Private Function Rating_EC(ByVal value As Double) As Integer

    Dim result As Integer
    
    If (value < 0.0000472) Then
        result = 1
    ElseIf (value < 0.000142) Then
        result = 2
    ElseIf (value < 0.00033) Then
        result = 4
    ElseIf (value < 0.000472) Then
        result = 6
    ElseIf (value < 0.000944) Then
        result = 8
    Else
        result = 10
    End If
              
    Rating_EC = result
    

End Function

Public Sub find_average()
    ' 2019/10/18 일 작성함
    ' get 투수량계수, 대수층, 유향, 동수경사의 평균을 구해야한다.

    Dim n_sheets As Integer
    Dim i As Integer
    
    Dim nTooSoo As Single: nTooSoo = 0
    Dim nDaeSoo As Single: nDaeSoo = 0
    Dim nDirection As Single: nDirection = 0
    Dim nGradient As Single: nGradient = 0
    
    
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
    
    Range("k3").Formula = "=round(j3,4)"
    Range("k4").Formula = "=round(j4,1)"
    Range("k5").Formula = "=round(j5,1)"
    Range("k6").Formula = "=round(j6,4)"
    
    Call make_frame
    
    
End Sub

Public Sub find_average2(ByVal sheet As Integer, ByVal nof_well As Integer)
    ' 2019/10/18 일 작성함
    ' get 투수량계수, 대수층, 유향, 동수경사의 평균을 구해야한다.

    Dim n_sheets As Integer
    Dim i As Integer
    
    Dim nTooSoo As Single: nTooSoo = 0
    Dim nDaeSoo As Single: nDaeSoo = 0
    Dim nDirection As Single: nDirection = 0
    Dim nGradient As Single: nGradient = 0
    
    
    
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
    
    Range("k3").Formula = "=round(j3,4)"
    Range("k4").Formula = "=round(j4,1)"
    Range("k5").Formula = "=round(j5,1)"
    Range("k6").Formula = "=round(j6,4)"
    
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
    
    Dim n_sheets As Integer
    Dim i As Integer
    

    ' 쉬트의 갯수 ..., 검사할 공의 갯수
    
    n_sheets = sheets_count()
      
    For i = 1 To n_sheets
    
        'Debug.Print "Loop counter : " & i
            
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
            
            
        ' Debug.Print " D : " & Soil
            
    Next i


End Sub


Function check_drasticindex() As String

    Dim value As Integer
    Dim result As String

    value = Range("k29").value
    
     If (value <= 100) Then
        result = "매우 낮음"
    ElseIf (value <= 120) Then
        result = "낮음"
    ElseIf (value <= 140) Then
        result = "비교적 낮음"
    ElseIf (value <= 160) Then
        result = "중간 정도"
    ElseIf (value <= 180) Then
        result = "높음"
    Else
        result = "매우 높음"
    End If
              
    check_drasticindex = result

End Function

Public Sub print_drastic_string()
    Dim n_sheets As Integer
    Dim i As Integer
    
    n_sheets = sheets_count()
      
    For i = 1 To n_sheets
         Worksheets(CStr(i)).Activate
         Range("k26").value = check_drasticindex()
    Next i

End Sub

Public Sub make_wellstyle()
    Dim n_sheets As Integer
    Dim i As Integer
    
    n_sheets = sheets_count()
    
    Call TurnOffStuff
    
    For i = 1 To n_sheets
         Worksheets(CStr(i)).Activate
         Call initialize_wellstyle
    Next i
    
    Call TurnOnStuff
    
End Sub

Private Function ConvertToLongInteger(ByVal stValue As String) As Long
    On Error GoTo ConversionFailureHandler
    ConvertToLongInteger = CLng(stValue)         'TRY to convert to an Integer value
    Exit Function                                'If we reach this point, then we succeeded so exit

ConversionFailureHandler:
    'IF we've reached this point, then we did not succeed in conversion
    'If the error is type-mismatch, clear the error and return numeric 0 from the function
    'Otherwise, disable the error handler, and re-run the code to allow the system to
    'display the error
    If Err.Number = 13 Then                      'error # 13 is Type mismatch
        Err.Clear
        ConvertToLongInteger = 0
        Exit Function
    Else
        On Error GoTo 0
        Resume
    End If

End Function

Public Function sheets_count() As Long

    Dim i, nSheetsCount, nWell  As Integer
    Dim strSheetsName(50) As String
    
    
    nSheetsCount = ThisWorkbook.Sheets.count
    nWell = 0
      
    
    For i = 1 To nSheetsCount
        strSheetsName(i) = ThisWorkbook.Sheets(i).Name
        'MsgBox (strSheetsName(i))
        If (ConvertToLongInteger(strSheetsName(i)) <> 0) Then
            nWell = nWell + 1
        End If
    Next
    
    'MsgBox (CStr(nWell))
    sheets_count = nWell

End Function








