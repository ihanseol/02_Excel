Attribute VB_Name = "z_Scrape"
Function GetHttp(URL As String, Optional formText As String, Optional isWinHttp As Boolean = False, Optional RequestHeader As Variant) As Object
 
'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ GetHttp 함수
'▶ 웹에서 데이터를 받아옵니다.
'▶ 인수 설명
'_____________URL                      : 데이터를 스크랩할 웹 페이지 주소입니다.
'_____________formText              : Encoding 된 FormText 형식으로 보내야 할 경우, Send String에 쿼리문을 추가합니다.
'_____________isWinHttp             : WinHTTP 로 요청할지 여부입니다. Redirect가 필요할 경우 True로 입력하여 WinHttp 요청을 전송합니다.
'_____________RequestHeader     : RequestHeader를 배열로 입력합니다. 반드시 짝수(한 쌍씩 이루어진) 개수로 입력되어야 합니다.
'▶ 사용 예제
'Dim HtmlResult As Object
'Set htmlResult = GetHttp("https://www.naver.com")
'msgbox htmlResult.body.innerHTML
'###############################################################
 
Dim oHTMLDoc As Object: Dim objHTTP As Object
Dim i As Long: Dim blnAgent As Boolean: blnAgent = False
Dim sUserAgent As String: sUserAgent = "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.183 Mobile Safari/537.36"
 
Application.DisplayAlerts = False
 
If Left(URL, 4) <> "http" Then URL = "http://" & URL
 
Set oHTMLDoc = CreateObject("HtmlFile")
 
If isWinHttp = False Then
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
Else
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
End If
 
objHTTP.setTimeouts 3000, 3000, 3000, 3000
objHTTP.Open "POST", URL, False
If Not IsMissing(RequestHeader) Then
    Dim vRequestHeader As Variant
    For Each vRequestHeader In RequestHeader
        Dim uHeader As Long: Dim Lheader As Long: Dim steps As Long
        uHeader = UBound(vRequestHeader): Lheader = LBound(vRequestHeader)
        If (uHeader - Lheader) Mod 2 = 0 Then GetHttp = CVErr(xlValue): Exit Function
        For i = Lheader To uHeader Step 2
            If vRequestHeader(i) = "User-Agent" Then blnAgent = True
            objHTTP.setRequestHeader vRequestHeader(i), vRequestHeader(i + 1)
        Next
    Next
End If
If blnAgent = False Then objHTTP.setRequestHeader "User-Agent", sUserAgent

objHTTP.send formText
 
With oHTMLDoc
    .Open
    .Write objHTTP.responsetext
    .Close
End With
 
Set GetHttp = oHTMLDoc
Set oHTMLDoc = Nothing
Set objHTTP = Nothing
 
Application.DisplayAlerts = True
 
End Function

Function Splitter(v As Variant, Cutter As String, Optional Trimmer As String)
 
'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ Splitter 함수
'▶ Cutter ~ Timmer 사이의 문자를 추출합니다. (Timmer가 빈칸일 경우 Cutter 이후 문자열을 추출합니다.)
'▶ 인수 설명
'_____________v       : 문자열입니다.
'_________Cutter      : 문자열 절삭을 시작할 텍스트입니다.
'_________Trimmer  : 문자열 절삭을 종료할 텍스트입니다. (선택인수)
'▶ 사용 예제
'Dim s As String
's = "{sa;b132@drama#weekend;aabbcc"
's = Splitter(s, "@", "#")
'msgbox s   '--> "drama"를 반환합니다.
'###############################################################
 
Dim vaArr As Variant

On Error GoTo EH:

vaArr = Split(v, Cutter)(1)
If Not IsMissing(Trimmer) Then vaArr = Split(vaArr, Trimmer)(0)
 
Splitter = vaArr

Exit Function

EH:
    Splitter = ""
    
End Function
