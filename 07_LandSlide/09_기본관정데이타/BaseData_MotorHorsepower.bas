Attribute VB_Name = "BaseData_MotorHorsepower"
Option Explicit

Public IP  As Long

Private Function nColorsInArray(ByRef array_tabcolor() As Variant, ByVal check As Variant) As Integer
    ' array_tabcolor :
    ' check : color값
    ' 관정에 지정하는 색갈은 모두 달라야 한다.
    ' 컬러값이 관정에 몇개가 있는지를 리턴
    
    Dim i, limit    As Integer
    Dim count       As Integer: count = 0
    
    limit = UBound(array_tabcolor, 1)
    
    For i = 1 To limit
        If array_tabcolor(i) = check Then
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
        C.Add Item:=CStr(uc), key:=CStr(i)
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

Private Function get_efficiency_A(ByVal Q As Variant) As Variant
    Dim result      As Variant
    
    If (Q < 57.6) Then
        result = 40
    ElseIf (Q < 72) Then
        result = 40
    ElseIf (Q < 86.4) Then
        result = 42
    ElseIf (Q < 115.2) Then
        result = 45
    ElseIf (Q < 144) Then
        result = 48
    ElseIf (Q < 216) Then
        result = 50
    ElseIf (Q < 288) Then
        result = 52
    ElseIf (Q < 432) Then
        result = 54
    ElseIf (Q < 576) Then
        result = 57
    ElseIf (Q < 720) Then
        result = 59
    ElseIf (Q < 864) Then
        result = 61
    ElseIf (Q < 1152) Then
        result = 62
    ElseIf (Q < 1440) Then
        result = 64
    Else
        result = 65
    End If
    
    get_efficiency_A = result
End Function

Private Function get_efficiency_B(ByVal Q As Variant) As Variant
    Dim result      As Variant
    
    If (Q < 72) Then
        result = 34
    ElseIf (Q < 86.4) Then
        result = 36
    ElseIf (Q < 115.2) Then
        result = 38
    ElseIf (Q < 144) Then
        result = 41
    ElseIf (Q < 216) Then
        result = 42
    ElseIf (Q < 288) Then
        result = 44
    ElseIf (Q < 432) Then
        result = 46
    ElseIf (Q < 576) Then
        result = 58
    ElseIf (Q < 720) Then
        result = 50
    ElseIf (Q < 864) Then
        result = 52
    ElseIf (Q < 1152) Then
        result = 53
    ElseIf (Q < 1440) Then
        result = 54
    Else
        result = 55
    End If
    
    get_efficiency_B = result
End Function

Private Function get_efficiency_dongho(ByVal Q As Variant) As Variant
    Dim result      As Variant
    
    If (Q < 72) Then
        result = 38
    ElseIf (Q < 86.4) Then
        result = (42 + 45 + 36 + 38) / 4
    ElseIf (Q < 115.2) Then
        result = (45 + 48 + 38 + 41) / 4
    ElseIf (Q < 144) Then
        result = (48 + 50 + 41 + 42) / 4
    ElseIf (Q < 216) Then
        result = (50 + 52 + 42 + 44) / 4
    ElseIf (Q < 288) Then
        result = (52 + 54 + 44 + 46) / 4
    ElseIf (Q < 432) Then
        result = (54 + 57 + 46 + 48) / 4
    ElseIf (Q < 576) Then
        result = (57 + 59 + 48 + 50) / 4
    ElseIf (Q < 720) Then
        result = (59 + 61 + 50 + 52) / 4
    ElseIf (Q < 864) Then
        result = (61 + 62 + 52 + 53) / 4
    ElseIf (Q < 1152) Then
        result = (62 + 64 + 53 + 54) / 4
    ElseIf (Q < 1440) Then
        result = (64 + 65 + 54 + 55) / 4
    Else
        result = 60
    End If
    
    get_efficiency_dongho = result
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
    Debug.Print mychar
    
    height = "=" & mychar & CStr(h1) & "+" & mychar & CStr(h1 + 1)
    height1 = "=round(" & mychar & CStr(h2 + 4) & "/10,1)"
    
    eq = "=round((" & mychar & CStr(h2 + 3) & "*" & mychar & CStr(h2 + 6) & ")/(6572.5*" & mychar & CStr(h2 + 7) & "),4)"
    round_hp = "=roundup(" & mychar & CStr(h2 + 9) & ",0)"
    theory_hp = "=round((" & mychar & CStr(h2 + 11) & "*" & mychar & CStr(h2 + 7) & "*6572.5)" & "/" & mychar & CStr(h2 + 6) & ",1)"
    
    Range(mychar & CStr(h2 + 5)).Formula = height1        '양정고
    Range(mychar & CStr(h2 + 6)).Formula = height        '합계
    
    Range(mychar & CStr(h2 + 9)).Formula = eq
    Range(mychar & CStr(h2 + 10)).Formula = round_hp
    Range(mychar & CStr(h2 + 12)).Formula = theory_hp
    
    Debug.Print height
    Debug.Print eq
    Debug.Print round_hp
    Debug.Print theory_hp
End Sub

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
    
    IP = lastRow() + 4
    
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
    
    Sheet4.Activate
    
    Call draw_motor_frame(nof_sheets, IP)
    
    For i = 1 To nof_sheets
        Call insert_basic_entry(title(i), simdo(i), pump_q(i), motor_depth(i), efficiency(i), hp(i), i, IP)
        Call insert_cell_function(i, IP)
    Next i
    
    Application.ScreenUpdating = True
End Sub

Private Sub insert_basic_entry(title As Variant, simdo As Variant, Q As Variant, motor_depth As Variant, e As Variant, hp As Variant, _
        ByVal i As Integer, ByVal po As Variant)
    Dim mychar
    
    mychar = Chr(65 + i)
    
    Range(mychar & CStr(po + 1)).value = title
    Range(mychar & CStr(po + 2)).value = simdo
    Range(mychar & CStr(po + 3)).value = Q
    Range(mychar & CStr(po + 4)).value = motor_depth
    Range(mychar & CStr(po + 7)).value = e / 100
    Range(mychar & CStr(po + 11)).value = hp
End Sub

