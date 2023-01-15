Attribute VB_Name = "Email_Module"
Option Explicit

'##########################################################
' 오빠두엑셀 퀵 VBA 7강, 메일 보내기 자동화 완성파일
' 명령문에 대한 자세한 설명은 아래 링크에서 확인하세요.
' https://www.oppadu.com/엑셀-메일-보내기-아웃룩-매크로/
'##########################################################

Sub Test()

Dim FileName As String
Dim SavePath As String

FileName = Sheet4.Range("E3").Value & "_" & Sheet4.Range("C4").Value & "_" & Sheet4.Range("H3").Value & "년 " & Sheet4.Range("H4").Value & "월"
SavePath = GetDesktopPath

Rng_To_Pdf Sheet4.Range("B2:E18"), FileName, SavePath, OpenPdf:=False, AddSequence:=False
'// Rng_To_Pdf 명령문의 첫번째 인수를 Selection으로 변경하면 선택한 범위를 PDF파일로 추출하여 메일에 첨부합니다.

Sheet4.Range("B2:E18").Select
'// 선택된 범위가 아닌 원하는 부분을 지정해서 메일에 첨부하는 방법이 궁금하신분은 Send_Email 명령문 관련 포스트를 참고하세요.

Send_Email "test@oppadu.com", _
            FileName, _
            "<p><span style=""font-family: NanumGothic, 나눔고딕, sans-serif; font-size: 9pt;""><b><u><span style=""font-size: 9pt;"">오빠두 대리</span></u></b> 님께&nbsp;</span></p><p><br></p><p><span style=""font-family: NanumGothic, 나눔고딕, sans-serif; font-size: 9pt;"">귀하의 <b><u><span style=""font-size: 9pt;"">2019년 10월 급여명세서</span></u></b>를 송부드립니다.&nbsp;</span></p><p><span style=""font-family: NanumGothic, 나눔고딕, sans-serif; font-size: 9pt;"">오빠두엑셀을 위한 귀하의 노고에 깊은 감사드리며 더욱 발전된 모습으로 귀하에 노고에 보답하겠습니다.</span></p>", _
            True, _
            "", , _
            SavePath & FileName & ".pdf" & "|" & SavePath & FileName & ".pdf"

End Sub

'######################################################################
' 명령문    : Send_Email
' 설명      : 아웃룩과 연동하여 메일보내기를 자동화하는 모듈입니다.
' 명령문에 대한 자세한 설명은 오빠두엑셀 홈페이지를 참고하세요.
' https://www.oppadu.com
'######################################################################

Sub Send_Email(MailTo As String, _
                Subject As String, _
                HTMLString As String, _
                Optional PasteSelection As Boolean = False, _
                Optional CCTo As String = "", _
                Optional BCCTo As String = "", _
                Optional AttachFilePath As String = "", _
                Optional PathDelimiter As String = "|")

Dim AppOutlook As Outlook.Application       '// 아웃룻 프로그램
Dim newEmail As Outlook.MailItem            '// 아웃룻 새로 메일을 보내기 위해 생성한 메일
Dim pageInspector As Outlook.Inspector      '// 아웃룩 워드에디터 가져오기위한 항목
Dim pageEditor As Object                    '// 아웃룩 이메일 편집창
Dim varFilePath As Variant                  '// 파일경로를 배열형태로 만들어준 변수
Dim FileCount As Long                       '// 첨부파일의 개수
Dim i As Long                               '// For문 반복문의 변수
Dim wdPasteDefault As Variant

Set AppOutlook = New Outlook.Application
Set newEmail = AppOutlook.CreateItem(olMailItem)


If AttachFilePath <> "" Then
        varFilePath = Split(AttachFilePath, PathDelimiter)
End If


With newEmail
    .To = MailTo
    .CC = CCTo
    .BCC = BCCTo
    .Subject = Subject
    If AttachFilePath <> "" Then
        For i = 1 To UBound(varFilePath) + 1
            .Attachments.Add varFilePath(i - 1), 1, i
        Next
    End If
    .HTMLBody = HTMLString
    
    '.DeferredDeliveryTime = DateAdd("n", 5, Now)
    .DeferredDeliveryTime = DateSerial(2030, 1, 1) + TimeSerial(8, 0, 0)
    
    If PasteSelection = True Then
        .Display
        Set pageInspector = newEmail.GetInspector
        Set pageEditor = pageInspector.WordEditor
        pageEditor.Application.Selection.Start = Len(.Body)
        
        Selection.Copy
        pageEditor.Application.Selection.PasteAndFormat wdPasteDefault
    Else
        .Display
    End If
    
'.Send    '// 메일을 보내려면 주석처리를 해제하세요.

End With

Set pageEditor = Nothing
Set pageInspector = Nothing
Set newEmail = Nothing
Set AppOutlook = Nothing

End Sub
