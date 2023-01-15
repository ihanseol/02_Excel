Attribute VB_Name = "달력"
Option Explicit
Sub Calendar()
'연도별 달력 만들기

    Dim i As Integer
    Dim Days As Integer, StartDay As Integer
    Dim Row As Integer, cnt As Integer
    Dim InputYear As Integer
    Dim theDay As Date
    Dim formSheet As Worksheet

    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    
        '기존 워크시트 삭제
        .DisplayAlerts = False
            For i = 3 To Sheets.Count
                Sheets(3).Delete
            Next i
        .DisplayAlerts = True
    End With
    
    '에러나면 매크로 종료
    On Error GoTo j
    
    Set formSheet = Sheets("서식")  '서식시트
    formSheet.Visible = True   '서식 시트 보이게
    InputYear = CInt(InputBox("몇 년도 달력을 만드세요??", "달력 만들기", year(Date)))
    
    For i = 1 To 12
        
        '서식시트에서 시트복사 후 시트명 변경
        formSheet.Copy after:=Sheets(Sheets.Count)
        Sheets(Sheets.Count).Name = i & "월"
        
        '월별 달력 만들기
        Range("B1") = InputYear '연도 출력
        Range("C1") = i '월
        Days = DateSerial(InputYear, i + 1, 1) - DateSerial(InputYear, i, 1)    '그 달의 총 날짜
        StartDay = Weekday(DateSerial(InputYear, i, 1), vbSunday) + 1   '달력 시작하는 날짜
        Row = 4
        cnt = 1
        
        '달력 입력
        Do
            theDay = DateSerial(InputYear, i, cnt)
            With Cells(Row, StartDay)
                .Value = theDay
        
                '주말, 공휴일은 빨간색으로
                With .Font
        
                    '주말
                    Select Case Weekday(theDay)
                        Case 1, 7: .Color = vbRed
                    End Select
            
                    '양력공휴일
                    Select Case Format(theDay, "m.d")
                        Case "1.1", "3.1", "5.5", "6.6", "8.15", "10.3", "10.9", "12.25"
                        .Color = vbRed
                    End Select
                    
                    '음력공휴일
                    Select Case Sol2Lun(InputYear, i, cnt)
                        Case "1.1", "1.2", "4.8", "8.14", "8.15", "8.16"
                        .Color = vbRed
                    End Select
                    If Sol2Lun(year(theDay + 1), Month(theDay + 1), Day(theDay + 1)) = "1.1" Then: .Color = vbRed '구정 전날은 전년도 12.31이므로 별도 설정
                    
                    '대체공휴일
                    Select Case theDay
                        Case #9/29/2015#, #2/10/2016#, #1/30/2017#, #9/26/2018#, #5/7/2018#, #5/6/2019#, #1/27/2020#, #9/12/2022#, #1/24/2023#, #2/12/2024#, #5/6/2024#, #10/8/2025#, #2/9/2027#, #9/24/2029#, #5/7/2029#
                            .Color = vbRed
                    End Select
                End With
            End With
        
            StartDay = StartDay + 1
            cnt = cnt + 1
            If StartDay = 9 Then
                StartDay = 2
                Row = Row + 2
            End If
        Loop While cnt <= Days
        
    Next i

    Sheets(3).Activate
    
j:
    formSheet.Visible = False
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

Sub ExportSchedule()

    RefChkBox.Show

End Sub

Sub ClearCalendar()
'달력 초기화

    Dim i As Integer
    
    Application.ScreenUpdating = False
    
    If Sheets.Count > 2 Then
        
        For i = 3 To Sheets.Count
            With Sheets(i)
                .Rows(5).ClearContents
                .Rows(7).ClearContents
                .Rows(9).ClearContents
                .Rows(11).ClearContents
                .Rows(13).ClearContents
                .Rows(15).ClearContents
            End With
            
        Next
    
    End If
    
    Application.ScreenUpdating = True
    
End Sub
