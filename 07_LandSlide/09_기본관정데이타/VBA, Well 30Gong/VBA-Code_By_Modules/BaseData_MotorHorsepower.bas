Attribute VB_Name = "BaseData_MotorHorsepower"
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
        .name = "Arial"
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
    
    mychar = ColumnNumberToLetter(n + 1)
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
    
    mychar = ColumnNumberToLetter(i + 1)
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
        .name = "Arial"
        .Size = 12
        .Bold = True
        .Italic = True
    End With
End Sub



