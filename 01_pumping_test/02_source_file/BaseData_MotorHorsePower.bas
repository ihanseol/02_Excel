Attribute VB_Name = "BaseData_MotorHorsePower"
Public IP As Long

Private Sub Range_End_Method()
    'Finds the last non-blank cell in a single row or column

    Dim lRow As Long
    Dim lCol As Long
    
    'Find the last non-blank cell in column A(1)
    lRow = Cells(Rows.count, 1).End(xlUp).Row
    
    'Find the last non-blank cell in row 1
    lCol = Cells(1, Columns.count).End(xlToLeft).Column
    
    MsgBox "Last Row: " & lRow & vbNewLine & _
           "Last Column: " & lCol
  
End Sub

Public Function lastRow() As Long

    Dim lRow As Long
    lRow = Cells(Rows.count, 1).End(xlUp).Row
    
    lastRow = lRow

End Function

Public Function Contains(Col As Collection, key As Variant) As Boolean
    On Error Resume Next
    Col (key)                                    ' Just try it. If it fails, Err.Number will be nonzero.
    Contains = (Err.Number = 0)
    Err.Clear
End Function

Function RemoveDupesDict(MyArray As Variant) As Variant

    'DESCRIPTION: Removes duplicates from your array using the dictionary method.
    'NOTES: (1.a) You must add a reference to the Microsoft Scripting Runtime library via
    ' the Tools > References menu.
    ' (1.b) This is necessary because I use Early Binding in this function.
    ' Early Binding greatly enhances the speed of the function.
    ' (2) The scripting dictionary will not work on the Mac OS.
    'SOURCE: https://wellsr.com
    '-----------------------------------------------------------------------
    Dim i As Long
    Dim d As Scripting.Dictionary
    Set d = New Scripting.Dictionary
    
    With d
        For i = LBound(MyArray) To UBound(MyArray)
            If IsMissing(MyArray(i)) = False Then
                .Item(MyArray(i)) = 1
            End If
        Next
        RemoveDupesDict = .Keys
    End With
    
    
End Function

Public Function GetLength(a As Variant) As Integer
    If IsEmpty(a) Then
        GetLength = 0
    Else
        GetLength = UBound(a) - LBound(a) + 1
    End If
End Function

Private Function getUnique(ByRef array_tabcolor As Variant) As Variant
  
    Dim array_size As Variant
    Dim new_array As Variant
    
    new_array = RemoveDupesDict(array_tabcolor)
    getUnique = new_array
   
    
End Function

Private Function nColorsInArray(ByRef array_tabcolor() As Variant, ByVal check As Variant) As Integer

    '관정에 지정하는 색갈은 모두 달라야 한다.


    Dim i, limit As Integer
    Dim count As Integer: count = 0

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
    Dim uc As Integer
   
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

    Dim n_sheets, i, j, limit As Integer
    Dim arr_tabcolors(), new_tabcolors() As Variant
   
    'uc : unique colors
    Dim uc As Integer
   
    'c colors code
    Dim c As Collection
    Set c = New Collection
    
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
        c.Add Item:=CStr(uc), key:=CStr(i)
    Next i

    
    Set getkey_tabcolors = c
   
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

Private Function get_efficiency_A(ByVal q As Variant) As Variant

    Dim result As Variant
    
    If (q < 72) Then
        result = 40
    ElseIf (q < 86.4) Then
        result = 42
    ElseIf (q < 115.2) Then
        result = 45
    ElseIf (q < 144) Then
        result = 48
    ElseIf (q < 216) Then
        result = 50
    ElseIf (q < 288) Then
        result = 52
    ElseIf (q < 432) Then
        result = 54
    ElseIf (q < 576) Then
        result = 57
    ElseIf (q < 720) Then
        result = 59
    ElseIf (q < 864) Then
        result = 61
    ElseIf (q < 1152) Then
        result = 62
    ElseIf (q < 1440) Then
        result = 64
    Else
        result = 65
    End If
              
    get_efficiency_A = result
    
End Function

Private Function get_efficiency_B(ByVal q As Variant) As Variant

    Dim result As Variant
    
    If (q < 72) Then
        result = 34
    ElseIf (q < 86.4) Then
        result = 36
    ElseIf (q < 115.2) Then
        result = 38
    ElseIf (q < 144) Then
        result = 41
    ElseIf (q < 216) Then
        result = 42
    ElseIf (q < 288) Then
        result = 44
    ElseIf (q < 432) Then
        result = 46
    ElseIf (q < 576) Then
        result = 58
    ElseIf (q < 720) Then
        result = 50
    ElseIf (q < 864) Then
        result = 52
    ElseIf (q < 1152) Then
        result = 53
    ElseIf (q < 1440) Then
        result = 54
    Else
        result = 55
    End If
              
    get_efficiency_B = result
    
End Function

Private Function get_efficiency_dongho(ByVal q As Variant) As Variant

    Dim result As Variant
    
    If (q < 72) Then
        result = 38
    ElseIf (q < 86.4) Then
        result = 40.25
    ElseIf (q < 115.2) Then
        result = 43
    ElseIf (q < 144) Then
        result = 45.25
    ElseIf (q < 216) Then
        result = 47
    ElseIf (q < 288) Then
        result = 49
    ElseIf (q < 432) Then
        result = 51.25
    ElseIf (q < 576) Then
        result = 53.5
    ElseIf (q < 720) Then
        result = 55.5
    ElseIf (q < 864) Then
        result = 57
    ElseIf (q < 1152) Then
        result = 58.25
    ElseIf (q < 1440) Then
        result = 59.5
    Else
        result = 60
    End If
              
    get_efficiency_dongho = result
    
End Function

Private Sub insert_cell_function(ByVal n As Integer, ByVal position As Integer)

    Dim mychar
    Dim height, eq, round_hp, theory_hp As String
    Dim h1, h2 As Integer
    
    h1 = position + 4
    h2 = position
    
    mychar = Chr(65 + n)
    Debug.Print mychar
    
    height = "=" & mychar & CStr(h1) & "+" & mychar & CStr(h1 + 1)
    eq = "=round((" & mychar & CStr(h2 + 3) & "*" & mychar & CStr(h2 + 6) & ")/(6572.5*" & mychar & CStr(h2 + 7) & "),4)"
    round_hp = "=roundup(" & mychar & CStr(h2 + 9) & ",0)"
    theory_hp = "=round((" & mychar & CStr(h2 + 11) & "*" & mychar & CStr(h2 + 7) & "*6572.5)" & "/" & mychar & CStr(h2 + 6) & ",1)"
    
    
    Range(mychar & CStr(h2 + 6)).Formula = height
    Range(mychar & CStr(h2 + 9)).Formula = eq
    Range(mychar & CStr(h2 + 10)).Formula = round_hp
    Range(mychar & CStr(h2 + 12)).Formula = theory_hp
 
 
    Debug.Print height
    Debug.Print eq
    Debug.Print round_hp
    Debug.Print theory_hp
    
    
    
End Sub

Public Sub getMotorPower()
    
    Dim r_ans() As Variant
    Dim rc As Collection                         'return collection
    Dim nof_sheets As Integer
    Dim nof_unique_tab As Integer
    Dim i, sheet As Integer
    
    Dim title() As Variant
    Dim simdo() As Variant
    Dim pump_q() As Variant
    Dim motor_depth() As Variant
    Dim efficiency() As Variant
    Dim hp() As Variant
    
    
    
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
        pump_q(i) = Range("c16").value
        motor_depth(i) = Range("c18").value
        efficiency(i) = get_efficiency_dongho(pump_q(i))
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

Private Sub insert_basic_entry(title As Variant, simdo As Variant, q As Variant, motor_depth As Variant, e As Variant, hp As Variant, _
                       ByVal i As Integer, ByVal po As Variant)

   
    Dim mychar
    
    mychar = Chr(65 + i)
    
    Range(mychar & CStr(po + 1)).value = title
    Range(mychar & CStr(po + 2)).value = simdo
    Range(mychar & CStr(po + 3)).value = q
    Range(mychar & CStr(po + 4)).value = motor_depth
    Range(mychar & CStr(po + 7)).value = e / 100
    Range(mychar & CStr(po + 11)).value = hp

End Sub

Sub getWhpaData_AllWell()
   
    Dim nof_sheets As Integer
    Dim nof_unique_tab As Integer
    Dim i, sheet As Integer
        
    
    Call get_tabsize(nof_sheets, nof_unique_tab)
    
    Application.ScreenUpdating = False
    
    For i = 1 To nof_sheets
        
        Call find_average2(i, 1)
        
    Next i
    
    Application.ScreenUpdating = True
    
End Sub

Sub getWhpaData_EachWell()
    
    Dim r_ans() As Variant
    Dim rc As Collection                         'return collection
    Dim nof_sheets As Integer
    Dim nof_unique_tab As Integer
    Dim i, sheet As Integer
        
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








