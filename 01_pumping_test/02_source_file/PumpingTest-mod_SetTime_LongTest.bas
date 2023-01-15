Attribute VB_Name = "mod_SetTime_LongTest"
Public MY_TIME As Integer



'10-77 : 2880 (68) - longterm pumping test
'78-101: recover (24) - recover test

Sub set_daydifference()

    Dim n_passed_time() As Integer
    Dim i As Integer
    Dim day1, day2 As Integer
    
    ReDim n_passed_time(1 To 92)
     
    For i = 1 To 92
        n_passed_time(i) = Cells(i + 9, "D").Value
        If (i > 68) Then
            n_passed_time(i) = Cells(i + 9, "D").Value + 2880
        End If
    Next i
    
    For i = 1 To 92
        Cells(i + 9, "h").Value = Range("c10").Value + n_passed_time(i) / 1440
    Next i
    
    Range("H10:H101").Select
    Selection.NumberFormatLocal = "yyyy""년"" m""월"" d""일"";@"
    Range("A1").Select
    Application.CutCopyMode = False

    
    Application.ScreenUpdating = False
    day1 = Day(Cells(10, "h").Value)

    For i = 2 To 92
        day2 = Day(Cells(i + 9, "h").Value)
        If (day2 = day1) Then
            Cells(i + 9, "h").Value = ""
        End If
        day1 = day2
    Next i
    
    Range("h77").Value = "양수종료"
    Range("h78").Value = "회복수위측정"
    Application.ScreenUpdating = True

End Sub

Function find_stable_time() As Integer

    Dim i As Integer
    
    
    For i = 30 To 50
            
        If Range("AC" & CStr(i)).Value = Range("AC" & CStr(i + 1)) Then
            'MsgBox "found " & "AB" & CStr(i) & " time : " & Range("Z" & CStr(i)).Value
            
            find_stable_time = i
            Exit For
        End If
        
    Next i
 
End Function

Function initialize_myTime() As Integer

    'Range("G17").Value = 840 + 60 * (i - 35)

    initialize_myTime = (shSkinFactor.Range("g17").Value - 840) / 60 + 35

   
End Function

Sub OptionButton_Setting(i As Integer)


    Select Case i
        Case 38:
            shLongTermTest.Frame1.Controls("OptionButton11").Value = True
            MY_TIME = 38
        Case 39:
            shLongTermTest.Frame1.Controls("OptionButton12").Value = True
            MY_TIME = 39
        Case 40:
            shLongTermTest.Frame1.Controls("OptionButton13").Value = True
            MY_TIME = 40
        Case 41:
            shLongTermTest.Frame1.Controls("OptionButton14").Value = True
            MY_TIME = 41
        Case 42:
            shLongTermTest.Frame1.Controls("OptionButton15").Value = True
            MY_TIME = 42
        Case 43:
            shLongTermTest.Frame1.Controls("OptionButton16").Value = True
            MY_TIME = 43
        Case 44:
            shLongTermTest.Frame1.Controls("OptionButton17").Value = True
            MY_TIME = 44
        Case 45:
            shLongTermTest.Frame1.Controls("OptionButton18").Value = True
            MY_TIME = 45
        Case 46:
            shLongTermTest.Frame1.Controls("OptionButton19").Value = True
            MY_TIME = 46
            
        Case Else:
            shLongTermTest.Frame1.Controls("OptionButton14").Value = True
            MY_TIME = 41
    End Select


End Sub

Sub TimeSetting()
    Dim stable_time, h1, h2, my_random_time As Integer
    Dim myRange As String
           
    stable_time = find_stable_time()
    
    If MY_TIME = 0 Then
    
        MY_TIME = initialize_myTime
        my_random_time = MY_TIME
        OptionButton_Setting (MY_TIME)
        'Frame1.Controls("OptionButton14").Value = True
    Else
        my_random_time = MY_TIME
    End If
    
    If stable_time < my_random_time Then
        h1 = stable_time
        h2 = my_random_time
        Range("ac" & CStr(h1)).Select
        myRange = "AC" & CStr(h1) & ":AC" & CStr(h2)
        
    ElseIf stable_time > my_random_time Then
        h1 = my_random_time
        h2 = stable_time
        Range("ac" & CStr(h2 + 1)).Select
        myRange = "AC" & CStr(h1 + 1) & ":AC" & CStr(h2 + 1)
    Else
        Exit Sub
    End If
              
    
    Selection.AutoFill Destination:=Range(myRange), Type:=xlFillDefault
    setSkinTime (MY_TIME)
    
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    
End Sub

Sub setSkinTime(i As Integer)

    Application.ScreenUpdating = False
    
    shSkinFactor.Activate
    Range("G17").Value = 840 + 60 * (i - 35)
    shLongTermTest.Activate
    
    Application.ScreenUpdating = True

End Sub

Sub setForRandomTime(i As Integer)

    Select Case i
        Case 38:
            shLongTermTest.Frame1.Controls("OptionButton11").Value = True
            MY_TIME = 38
        Case 39:
            shLongTermTest.Frame1.Controls("OptionButton12").Value = True
            MY_TIME = 39
        Case 40:
            shLongTermTest.Frame1.Controls("OptionButton13").Value = True
            MY_TIME = 40
        Case 41:
            shLongTermTest.Frame1.Controls("OptionButton14").Value = True
            MY_TIME = 41
        Case 42:
            shLongTermTest.Frame1.Controls("OptionButton15").Value = True
            MY_TIME = 42
        Case 43:
            shLongTermTest.Frame1.Controls("OptionButton16").Value = True
            MY_TIME = 43
        Case 44:
            shLongTermTest.Frame1.Controls("OptionButton17").Value = True
            MY_TIME = 44
        Case 45:
            shLongTermTest.Frame1.Controls("OptionButton18").Value = True
            MY_TIME = 45
        Case 46:
            shLongTermTest.Frame1.Controls("OptionButton19").Value = True
            MY_TIME = 46
            
        Case Else:
            shLongTermTest.Frame1.Controls("OptionButton14").Value = True
            MY_TIME = 41
    End Select


    Call setSkinTime(i)
    

End Sub

Sub RandomTimeSetting()
    Dim my_random_time As Integer
    Dim stable_time, h1, h2 As Integer
    Dim myRange As String
           
    Randomize                                    'Initialize the Rnd function
     
    my_random_time = CInt(38 + Rnd * 6)                'Generate a random number between 5-100
    'MsgBox CStr(my_random_time)
    
    stable_time = find_stable_time()
    
    If stable_time < my_random_time Then
        h1 = stable_time
        h2 = my_random_time
        Range("ac" & CStr(h1)).Select
        myRange = "AC" & CStr(h1) & ":AC" & CStr(h2)
        
    ElseIf stable_time > my_random_time Then
        h1 = my_random_time
        h2 = stable_time
        Range("ac" & CStr(h2 + 1)).Select
        myRange = "AC" & CStr(h1 + 1) & ":AC" & CStr(h2 + 1)
    Else
        Exit Sub
    End If
              
    Selection.AutoFill Destination:=Range(myRange), Type:=xlFillDefault
    
    Call setForRandomTime(my_random_time)
    
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Selection.End(xlUp).Select
    
End Sub

Sub cellRED(ByVal strcell As String)

    Range(strcell).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13209
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    
End Sub

Sub cellBLACK(ByVal strcell As String)

    Range(strcell).Select
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0.499984740745262
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    
End Sub

Sub resetValue()
    
    Range("p3").ClearContents
    Range("t1").Value = 0.1
    Range("l6").Value = 0.2
        
    Range("o3:o14").ClearContents
   

End Sub


Function isPositive(ByVal data As Double) As Double

    If data < 0 Then
        isPositive = False
    Else
        isPositive = True
    End If
    
End Function


Function CellReverse(ByVal data As Double) As Double

If data < 0 Then
    CellReverse = Abs(data)
Else
    CellReverse = -data
End If


End Function

Sub findAnswer_LongTest()
    
    If (Range("p3").Value > 0) Then Exit Sub
    
    Range("l10").GoalSeek goal:=0, ChangingCell:=Range("t1")
    
    Range("p3").Value = CellReverse(Range("k10").Value)
    
    If Range("l8").Value < 0 Then
        cellRED ("l8")
    Else
        cellBLACK ("l8")
    End If
    
    shSkinFactor.Range("d5").Value = Round(Range("t1").Value, 4)
    
End Sub

Sub check_LongTest()

    Dim igoal, k0, k1 As Double
    
    k1 = Range("l8").Value
    k0 = Range("l6").Value
    
    If k0 = k1 Then Exit Sub
    If k1 > 0 Then Exit Sub
    
    If k0 <> "" Then
        igoal = k0
    Else
        igoal = 0.3
    End If
    
    Range("l8").GoalSeek goal:=igoal, ChangingCell:=Range("o3")
     
    If Range("l8").Value < 0 Then
        cellRED ("l8")
    Else
        cellBLACK ("l8")
    End If
    

End Sub

Sub findAnswer_StepTest()
   
    Range("Q4:Q13").ClearContents
    Range("T4").Value = 0.1
    Range("G12").GoalSeek goal:=1#, ChangingCell:=Range("T4")
    
    If Range("J11").Value < 0 Then
        Call cellRED("J11")
    Else
        Call cellBLACK("J11")
    End If
    
End Sub

Sub check_StepTest()

    Dim igoal, nj As Double
    
    igoal = 0.12
    
    Do While (Range("J11").Value < 0 Or Range("j11").Value >= 50)
        Range("J11").GoalSeek goal:=igoal, ChangingCell:=Range("Q4")
        igoal = igoal + 0.1
    Loop
    
    If Range("J11").Value < 0 Then
        cellRED ("J11")
    Else
        cellBLACK ("J11")
    End If
    
    

End Sub



