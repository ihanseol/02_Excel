Attribute VB_Name = "z_Others"
Option Explicit

'---------------------------------------------------
' 오늘날짜 확인 및 API 사용량 초기화
'---------------------------------------------------
Sub Auto_Open()
'마지막으로 로그인 한 날짜가 오늘 날짜와 다를 경우 API 사용량을 초기화합니다.
If Sheet2.Range("U8").Value <> Date Then
    Sheet2.Range("U8").Value = Date
    Sheet2.Range("U6").Value = 0
End If
End Sub

'---------------------------------------------------
' 클립보드 복사/붙여넣기 후 쿼리 초기화 명령문
'---------------------------------------------------
Sub RefreshAll()
'클립보드에 복사된 값을 [원문] 시트 안에 텍스트 형태로 붙여넣기 한 뒤
'쿼리를 업데이트 합니다.
Dim Rng As Range
Application.ScreenUpdating = False

With Sheet1
    Set Rng = .UsedRange
    If Rng.Rows.Count > 1 Then .Range("2:" & Rng.Rows.Count).EntireRow.Delete
    .Activate
    .Range("A2").Select
    If Application.LanguageSettings.LanguageID(msoLanguageIDUI) = 1042 Then
        .PasteSpecial Format:="유니코드 텍스트"
    Else
        .PasteSpecial Format:="Unicode Text"
    End If
End With
Sheet2.Activate
Application.ScreenUpdating = True
ThisWorkbook.RefreshAll

End Sub

'---------------------------------------------------
' 음성변환 관련 명령문
'---------------------------------------------------
Sub vocStop()
' 실행중인 음성변환을 중단합니다.
VocSpeak blnStop:=True

End Sub
Sub vocSpeak_To()
' 번역된 텍스트를 음성으로 변환합니다.
Dim x As Long: Dim i As Long
Dim s As String
With Sheet2
    x = .Range("P14").End(xlDown).Row
    If x = .Rows.Count Then x = 14
    For i = 14 To x
        s = s & .Range("P" & i).Value
    Next
    VocSpeak s, Application.WorksheetFunction.VLookup(Sheet2.Range("P9").Value, Sheet3.Range("A:C"), 3, 0)
End With
End Sub
Sub vocSpeak_From()
' 번역되기 전 텍스트를 음성으로 변환합니다.
Dim x As Long: Dim i As Long
Dim s As String
With Sheet2
    x = .Range("F14").End(xlDown).Row
    If x = .Rows.Count Then x = 14
    For i = 14 To x
        s = s & .Range("F" & i).Value
    Next
    VocSpeak s, Application.WorksheetFunction.VLookup(Sheet2.Range("F9").Value, Sheet3.Range("A:C"), 3, 0)
End With
End Sub
