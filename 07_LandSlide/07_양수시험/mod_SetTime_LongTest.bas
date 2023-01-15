Attribute VB_Name = "mod_SetTime_LongTest"
Public MY_TIME      As Integer

Public gDicStableTime As Scripting.Dictionary
Public gDicMyTime   As Scripting.Dictionary

Sub initDictionary()
    Set gDicStableTime = New Scripting.Dictionary
    Set gDicMyTime = New Scripting.Dictionary
    
    gDicStableTime.Add Key:=60, Item:=17
    gDicStableTime.Add Key:=75, Item:=18
    gDicStableTime.Add Key:=90, Item:=19
    gDicStableTime.Add Key:=105, Item:=20
    gDicStableTime.Add Key:=120, Item:=21
    gDicStableTime.Add Key:=140, Item:=22
    gDicStableTime.Add Key:=160, Item:=23
    gDicStableTime.Add Key:=180, Item:=24
    gDicStableTime.Add Key:=240, Item:=25
    gDicStableTime.Add Key:=300, Item:=26
    gDicStableTime.Add Key:=360, Item:=27
    gDicStableTime.Add Key:=420, Item:=28
    gDicStableTime.Add Key:=480, Item:=29
    gDicStableTime.Add Key:=540, Item:=30
    gDicStableTime.Add Key:=600, Item:=31
    gDicStableTime.Add Key:=660, Item:=32
    gDicStableTime.Add Key:=720, Item:=33
    gDicStableTime.Add Key:=780, Item:=34
    gDicStableTime.Add Key:=840, Item:=35
    gDicStableTime.Add Key:=900, Item:=36
    gDicStableTime.Add Key:=960, Item:=37
    gDicStableTime.Add Key:=1020, Item:=38
    gDicStableTime.Add Key:=1080, Item:=39
    gDicStableTime.Add Key:=1140, Item:=40
    gDicStableTime.Add Key:=1200, Item:=41
    gDicStableTime.Add Key:=1260, Item:=42
    gDicStableTime.Add Key:=1320, Item:=43
    gDicStableTime.Add Key:=1380, Item:=44
    gDicStableTime.Add Key:=1440, Item:=45
    gDicStableTime.Add Key:=1500, Item:=46
    
    gDicMyTime.Add Key:=17, Item:=60
    gDicMyTime.Add Key:=18, Item:=75
    gDicMyTime.Add Key:=19, Item:=90
    gDicMyTime.Add Key:=20, Item:=105
    gDicMyTime.Add Key:=21, Item:=120
    gDicMyTime.Add Key:=22, Item:=140
    gDicMyTime.Add Key:=23, Item:=160
    gDicMyTime.Add Key:=24, Item:=180
    gDicMyTime.Add Key:=25, Item:=240
    gDicMyTime.Add Key:=26, Item:=300
    gDicMyTime.Add Key:=27, Item:=360
    gDicMyTime.Add Key:=28, Item:=420
    gDicMyTime.Add Key:=29, Item:=480
    gDicMyTime.Add Key:=30, Item:=540
    gDicMyTime.Add Key:=31, Item:=600
    gDicMyTime.Add Key:=32, Item:=660
    gDicMyTime.Add Key:=33, Item:=720
    gDicMyTime.Add Key:=34, Item:=780
    gDicMyTime.Add Key:=35, Item:=840
    gDicMyTime.Add Key:=36, Item:=900
    gDicMyTime.Add Key:=37, Item:=960
    gDicMyTime.Add Key:=38, Item:=1020
    gDicMyTime.Add Key:=39, Item:=1080
    gDicMyTime.Add Key:=40, Item:=1140
    gDicMyTime.Add Key:=41, Item:=1200
    gDicMyTime.Add Key:=42, Item:=1260
    gDicMyTime.Add Key:=43, Item:=1320
    gDicMyTime.Add Key:=44, Item:=1380
    gDicMyTime.Add Key:=45, Item:=1440
    gDicMyTime.Add Key:=46, Item:=1500
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
    initialize_myTime = gDicStableTime(shSkinFactor.Range("g16").Value)
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
    
    shSkinFactor.Activate
    Range("G16").Value = gDicMyTime(i)
    shLongTermTest.Activate
    
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
