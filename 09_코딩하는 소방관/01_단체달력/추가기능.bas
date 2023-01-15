Attribute VB_Name = "추가기능"
Sub Auto_Open()
 
    Auto_Close
    
    On Error Resume Next
    
    With Application.CommandBars("Tools").Controls
    
        With .Add(Type:=msoControlButton)
            .FaceId = 8
            .Caption = "달력 만들기"
            .OnAction = "Calendar"
        End With
    
        With .Add(Type:=msoControlButton)
            .FaceId = 9644
            .Caption = "일정 추출하기"
            .OnAction = "ExportSchedule"
        End With
    
        With .Add(Type:=msoControlButton)
            .FaceId = 47
            .Caption = "달력 초기화"
            .OnAction = "ClearCalendar"
        End With
        
        With .Add(Type:=msoControlButton)
            .FaceId = 3
            .Caption = "월별 시트 저장"
            .OnAction = "mkFile"
        End With
    
    End With
    
    On Error GoTo 0
    
End Sub
 
Sub Auto_Close()
 
    Application.CommandBars("Tools").Reset
    
End Sub
