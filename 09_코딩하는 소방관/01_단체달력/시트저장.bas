Attribute VB_Name = "시트저장"
Sub mkFile()
 
    Dim inTemp As Variant
    Dim vrTemp As Variant
    Dim oldBook As Workbook
    Dim newBook As Workbook
    Dim i As Integer
    
    '처리 속도 높이기
    With Application
        .Calculation = xlCalculationManual
        .EnableEvents = False
        .ScreenUpdating = False
    End With
    
    On Error GoTo j '에러나면 매크로 종료
    
    '현재 워크시트 정의
    Set oldBook = ActiveWorkbook
    
    'inputbox를 이용하여 저장할 시트번호 받음
    inTemp = InputBox("저장할 달을 입력하세요" & vbCr & vbCr & "예) 1,2,3")
    If TypeName(inTemp) = "Boolean" Or inTemp = "" Then GoTo j
    
    '콤마를 기준으로 받은 값 분리
    vrTemp = Split(inTemp, ",")
    
    '시트번호를 받아 새파일로 저장
    For i = 1 To UBound(vrTemp) + 1
        If i = 1 Then
            Sheets(Val(vrTemp(i - 1)) + 2).Copy
            Set newBook = ActiveWorkbook
            
        Else
            oldBook.Sheets(Val(vrTemp(i - 1)) + 2).Copy after:=newBook.Sheets(newBook.Sheets.Count)
        End If
    Next i
    
j:
    '처리속도 복구 후 새파일 선택
    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
    End With
    
End Sub

