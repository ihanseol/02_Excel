Attribute VB_Name = "mod_W1LongtermTEST"
Public MY_TIME      As Integer

Public gDicStableTime As Scripting.Dictionary
Public gDicMyTime   As Scripting.Dictionary

Sub initDictionary()
    Set gDicStableTime = New Scripting.Dictionary
    Set gDicMyTime = New Scripting.Dictionary
    
    gDicStableTime.Add Key:=60, item:=17
    gDicStableTime.Add Key:=75, item:=18
    gDicStableTime.Add Key:=90, item:=19
    gDicStableTime.Add Key:=105, item:=20
    gDicStableTime.Add Key:=120, item:=21
    gDicStableTime.Add Key:=140, item:=22
    gDicStableTime.Add Key:=160, item:=23
    gDicStableTime.Add Key:=180, item:=24
    gDicStableTime.Add Key:=240, item:=25
    gDicStableTime.Add Key:=300, item:=26
    gDicStableTime.Add Key:=360, item:=27
    gDicStableTime.Add Key:=420, item:=28
    gDicStableTime.Add Key:=480, item:=29
    gDicStableTime.Add Key:=540, item:=30
    gDicStableTime.Add Key:=600, item:=31
    gDicStableTime.Add Key:=660, item:=32
    gDicStableTime.Add Key:=720, item:=33
    gDicStableTime.Add Key:=780, item:=34
    gDicStableTime.Add Key:=840, item:=35
    gDicStableTime.Add Key:=900, item:=36
    gDicStableTime.Add Key:=960, item:=37
    gDicStableTime.Add Key:=1020, item:=38
    gDicStableTime.Add Key:=1080, item:=39
    gDicStableTime.Add Key:=1140, item:=40
    gDicStableTime.Add Key:=1200, item:=41
    gDicStableTime.Add Key:=1260, item:=42
    gDicStableTime.Add Key:=1320, item:=43
    gDicStableTime.Add Key:=1380, item:=44
    gDicStableTime.Add Key:=1440, item:=45
    gDicStableTime.Add Key:=1500, item:=46
    
    gDicMyTime.Add Key:=17, item:=60
    gDicMyTime.Add Key:=18, item:=75
    gDicMyTime.Add Key:=19, item:=90
    gDicMyTime.Add Key:=20, item:=105
    gDicMyTime.Add Key:=21, item:=120
    gDicMyTime.Add Key:=22, item:=140
    gDicMyTime.Add Key:=23, item:=160
    gDicMyTime.Add Key:=24, item:=180
    gDicMyTime.Add Key:=25, item:=240
    gDicMyTime.Add Key:=26, item:=300
    gDicMyTime.Add Key:=27, item:=360
    gDicMyTime.Add Key:=28, item:=420
    gDicMyTime.Add Key:=29, item:=480
    gDicMyTime.Add Key:=30, item:=540
    gDicMyTime.Add Key:=31, item:=600
    gDicMyTime.Add Key:=32, item:=660
    gDicMyTime.Add Key:=33, item:=720
    gDicMyTime.Add Key:=34, item:=780
    gDicMyTime.Add Key:=35, item:=840
    gDicMyTime.Add Key:=36, item:=900
    gDicMyTime.Add Key:=37, item:=960
    gDicMyTime.Add Key:=38, item:=1020
    gDicMyTime.Add Key:=39, item:=1080
    gDicMyTime.Add Key:=40, item:=1140
    gDicMyTime.Add Key:=41, item:=1200
    gDicMyTime.Add Key:=42, item:=1260
    gDicMyTime.Add Key:=43, item:=1320
    gDicMyTime.Add Key:=44, item:=1380
    gDicMyTime.Add Key:=45, item:=1440
    gDicMyTime.Add Key:=46, item:=1500
End Sub

'10-77 : 2880 (68) - longterm pumping test
'78-101: recover (24) - recover test

Sub set_daydifference()
    Dim n_passed_time() As Integer
    Dim i           As Integer
    Dim day1, day2  As Integer
    
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
    Dim i           As Integer
    
    For i = 10 To 50
        If Range("AC" & CStr(i)).Value = Range("AC" & CStr(i + 1)) Then
            'MsgBox "found " & "AB" & CStr(i) & " time : " & Range("Z" & CStr(i)).Value
            find_stable_time = i
            Exit For
        End If
    Next i
End Function

Function initialize_myTime() As Integer
    initialize_myTime = gDicStableTime(shW_aSkinFactor.Range("g16").Value)
End Function

Sub TimeSetting()
    Dim stable_time, h1, h2, my_random_time As Integer
    Dim myRange     As String
    
    stable_time = find_stable_time()
    
    If MY_TIME = 0 Then
        MY_TIME = initialize_myTime
        my_random_time = MY_TIME
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
    
    Range("A27").Select
End Sub

Sub setSkinTime(i As Integer)
    Application.ScreenUpdating = False
    
    shW_aSkinFactor.Activate
    Range("G16").Value = gDicMyTime(i)
    shW_LongTEST.Activate
    
    Application.ScreenUpdating = True
End Sub

Sub cellRED(ByVal strcell As String)
    Range(strcell).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 13209
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
    
    '2023/4/3
    If Not IsNumeric(Range("GoalSeekTarget").Value) Then
        Range("GoalSeekTarget").Value = 0
    End If
    
    Range("o3:o14").ClearContents
    
    ' ActiveSheet.OptionButton1.Value = True
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
    ' 2023/4/3
    Dim GoalSeekTarget As Double
    
    
    ' 2023/4/3
    If Not IsNumeric(Range("GoalSeekTarget").Value) Then
        GoalSeekTarget = 0
    Else
        GoalSeekTarget = Range("GoalSeekTarget").Value
    End If
    
    If (Range("p3").Value <> 0) Then
        Exit Sub
    End If
    
    
  
    Range("L10").GoalSeek goal:=GoalSeekTarget, ChangingCell:=Range("T1")
    Range("p3").Value = CellReverse(Range("k10").Value)
    
    If Range("l8").Value < 0 Then
        cellRED ("l8")
    Else
        cellBLACK ("l8")
    End If
    
    shW_aSkinFactor.Range("d5").Value = Round(Range("t1").Value, 4)
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
    
    Call drawdown_level_check
End Sub

' 2023/4/4
Sub drawdown_level_check()

    Dim check_cell(1 To 5) As String
    Dim i As Integer
    
    check_cell(1) = "G10"
    check_cell(2) = "K11"
    check_cell(3) = "K12"
    check_cell(4) = "K13"
    check_cell(5) = "K14"
    
    For i = 1 To 4
        If Range(check_cell(i)).Value < Range(check_cell(i + 1)).Value Then
            MsgBox "Water Level is Trouble ....", vbOKOnly
            Exit Sub
        End If
    Next i

End Sub



Sub findAnswer_StepTest()
    Range("Q4:Q13").ClearContents
    Range("T4").Value = 0.1
    Range("G12").GoalSeek goal:=1#, ChangingCell:=Range("T4")
    
    If Range("J11").Value < 0 Then
        cellRED ("J11")
    Else
        cellBLACK ("J11")
    End If
End Sub

Sub check_StepTest()
    Dim igoal, nj   As Double
    
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
