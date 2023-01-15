Attribute VB_Name = "Main"
Option Explicit

Sub DoTranslate()

'-----------------------------------------
' 1. 변수 선언 (Rng, r)
'-----------------------------------------
Dim Rng As Range            '원본 언어가 입력된 범위
Dim r As Range                  'Rng 를 하나씩 돌아가면서 참조할 범위

'-----------------------------------------
' 2. 기존 번역되기 전/후 값이 있을 경우 초기화
'-----------------------------------------
On Error Resume Next
Sheet2.Range("tblResult").Delete
On Error GoTo 0

'-----------------------------------------
' 3. 원본 문장이 입력된 범위 설정
'-----------------------------------------
Set Rng = Sheet2.Range("tblOriginal")

'-----------------------------------------
' 4. 범위를 하나씩 돌아가며 Papago 번역 실행 -> 결과셀로 반환
'-----------------------------------------
For Each r In Rng
    Sheet2.Range("P" & r.Row).Value = PapagoTranslate(r.Value)
Next

'-----------------------------------------
' 5. 파파고 API 사용량 업데이트
'-----------------------------------------
Sheet2.Range("U6").Value = Sheet2.Range("U6").Value + Sheet2.Range("F11").Value

MsgBox "문장 번역이 완료되었습니다!"

End Sub

Sub asdlf()



End Sub
'=VLOOKUP(찾을값,범위,열번호...)
'=PapagoTranslate(문자)
Function PapagoTranslate(OriginalText) As String

'-----------------------------------------
'1. 변수 선언하기
'-----------------------------------------
Dim sID As String           ' API 키, 보안키
Dim sSecret As String
Dim sFrom As String         ' 원본 언어
Dim sTo As String           ' 목적 언어

Dim URL As String           ' API URL ("https://openapi.naver.com/v1/papago/n2mt")
Dim Query As String         ' 쿼리 스트링 ("source=ko&target=en&text=안녕하세요. 만나서 반갑습니다.")
Dim vArray As Variant       ' HTTP 요청을 보낼 Request Header
Dim objHTML As HTMLDocument ' HTTP 요청으로 반환된 HTML 문서
Dim sResult As String       ' HTTML 요청으로 반환된 HTML 텍스트

'-----------------------------------------
'2. 변수 설정
'-----------------------------------------
'API키와 보안키를 설정
sID = Sheet3.Range("B1").Value
sSecret = Sheet3.Range("B2").Value

'원본언어와 목적언어의 코드 찾기 (VLOOKUP 함수)
sFrom = Application.WorksheetFunction.VLookup(Sheet2.Range("F9").Value, Sheet3.Range("A:B"), 2, 0)
sTo = Application.WorksheetFunction.VLookup(Sheet2.Range("P9").Value, Sheet3.Range("A:B"), 2, 0)

''HTTP 요청을 보낼 URL과 Request Header 설정하기
URL = "https://openapi.naver.com/v1/papago/n2mt"
Query = "source=" & sFrom & "&target=" & sTo & "&text=" & OriginalText

'Request Header는 총 4개, Uger-Agent, Content-Type, Client-Id, Client-Secret
ReDim vArray(0 To 3)
vArray(0) = Array("User-Agent", "curl/749.1")
vArray(1) = Array("Content-Type", "application/x-www-form-urlencoded; charset=UTF-8")
vArray(2) = Array("X-Naver-Client-Id", CStr(sID))
vArray(3) = Array("X-Naver-Client-Secret", CStr(sSecret))

'' 쿼리는 아래와 같이 입력되겠죠? :)
''"source=ko&target=en&text=안녕하세요. 만나서 반갑습니다."
''위와 같이 쿼리가 들어가도록 명령문을 작성합니다.
'-----------------------------------------
'3. HTTP 요청으로 API 결과 받아오기
'-----------------------------------------
Set objHTML = GetHttp(URL, Query, False, vArray)
sResult = objHTML.body.innerHTML

'-----------------------------------------
'4. Splitter로 원하는 결과만 추출하기
'-----------------------------------------
sResult = Splitter(sResult, "translatedText"":""", """,""engineType")

'오류가 발생할 경우 오류메시지를 반환
If sResult = "" Then sResult = Splitter(sResult, "errorMessage"":""", """,""errorCode")

'-----------------------------------------
'5. PapagoTranslate 함수로 결과 반환 후 명령문 종료
'-----------------------------------------
PapagoTranslate = sResult

End Function
