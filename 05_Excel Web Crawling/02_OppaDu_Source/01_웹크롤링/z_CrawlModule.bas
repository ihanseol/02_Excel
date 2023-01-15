Attribute VB_Name = "z_CrawlModule"

Public Function HasContent(strName As String) As Boolean
    HasContent = (Len(Trim(strName)) > 0)
End Function



Function GetLastRow(WS As Worksheet) As Long

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ GetLastRow 함수
'▶ 시트에서 사용된 마지막 행을 반환합니다.
'▶ 인수 설명
'_____________WS        : 마지막 행을 조회할 대상 시트입니다.
'###############################################################

With WS
GetLastRow = .UsedRange.Rows.Count + .UsedRange.Row - 1
End With

End Function

Sub InsertWebImage(targetRng As Range, ImgLink As String, Optional imgWidth As Double, Optional imgHeight As Double)

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ InsertWebImage 함수
'▶ 웹 이미지를 셀 위에 삽입합니다.
'▶ 인수 설명
'_____________targetRng                              : 이미지를 삽입할 셀입니다.
'_____________imgLink                                 : 삽입할 이미지 링크입니다.
'_____________imgWidth                             : 이미지 너비입니다. 기본값은 셀의 너비입니다. (선택인수)
'_____________imgHeight                            : 이미지 높이입니다. 기본값은 셀의 높이입니다. (선택인수)
'▶ 사용 예제
'InsertWebImage(sheet1.Range("A1"),"https://www.google.com/images/branding/googlelogo/1x/googlelogo_color_272x92dp.png")
'###############################################################

Dim WS As Worksheet
Dim shp As Shape
Dim W As Double: Dim H As Double

Set WS = targetRng.Parent
If IsMissing(imgWidth) Then W = imgWidth Else W = targetRng.Width - 6
If IsMissing(imgHeight) Then W = imgHeight Else H = targetRng.Height - 6

Set shp = WS.Shapes.AddPicture(ImgLink, msoFalse, msoTrue, targetRng.Left + 3, targetRng.Top + 3, W, H)

End Sub

Function GetHttp(url As String, ParamArray RequestHeader() As Variant) As Object

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ GetHttp 함수
'▶ 웹에서 데이터를 받아옵니다.
'▶ 인수 설명
'_____________URL                               : 데이터를 스크랩할 웹 페이지 주소입니다.
'_____________RequestHeader            : RequestHeader를 배열로 입력합니다. 반드시 짝수(한 쌍씩 이루어진) 개수로 입력되어야 합니다.
'▶ 사용 예제
'Dim HtmlResult As Object
'Set htmlResult = GetHttp("https://www.naver.com")
'msgbox htmlResult.body.innerHTML
'###############################################################

Dim oHTMLDoc As Object: Dim objHTTP As Object

Dim sUserAgent As String: sUserAgent = "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.183 Mobile Safari/537.36"

Application.DisplayAlerts = False

Set oHTMLDoc = CreateObject("HtmlFile")
Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")

objHTTP.setTimeouts 3000, 3000, 3000, 3000
objHTTP.Open "GET", url, False
objHTTP.setRequestHeader "User-Agent", sUserAgent
If Not IsMissing(RequestHeader) Then
    Dim vRequestHeader As Variant
    For Each vRequestHeader In RequestHeader
        Dim uHeader As Long: Dim Lheader As Long: Dim steps As Long
        uHeader = UBound(vRequestHeader): Lheader = LBound(vRequestHeader)
        If (uHeader - Lheader) Mod 2 = 0 Then GET_HttpRequest = CVErr(xlValue): Exit Function
        For i = Lheader To uHeader Step 2
            objHTTP.setRequestHeader vRequestHeader(i), vRequestHeader(i + 1)
        Next
    Next
End If
objHTTP.send
        
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

Sub ExportText(InnerStrings As String, Optional fileName As String = "텍스트추출", Optional Path As String)

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ Export_Text 함수
'▶ 문자열을 텍스트파일로 추출합니다.
'▶ 인수 설명
'_____________InnerStrings      : 텍스트파일로 추출할 문자열입니다.
'_____________fileName           : 텍스트 파일 이름입니다. 기본값은 "텍스트추출" 입니다. (선택인수)
'_____________path                   : 텍스트 파일을 생성할 경로입니다. 기본값은 바탕화면입니다. (선택인수)
'▶ 사용 예제
'ExportText "추출할 텍스트"
'###############################################################

Dim TextFile As Integer
Dim FilePath As String

If Not IsMissing(Path) Then Path = Environ("USERPROFILE") & "\Desktop\"
FilePath = Path & fileName & ".txt"

TextFile = FreeFile

Open FilePath For Output As TextFile
Print #TextFile, InnerStrings
Close TextFile

End Sub

Sub ArrayToRng(startRng As Range, Arr As Variant)

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ ArrayToRng 함수
'▶ 배열을 범위 위로 반환합니다.
'▶ 인수 설명
'_____________startRng      : 배열을 반환할 기준 범위(셀) 입니다.
'_____________Arr               : 반환할 배열입니다.
'▶ 사용 예제
'Dim v As Variant
'ReDim v(0 to 1)
''v(0) = "a" : v(1) = "b"
'ArrayToRng Sheet1.Range("A1"), v
'##############################################################

On Error GoTo SingleDimension:
startRng.Cells(1, 1).Resize(UBound(Arr, 1) - LBound(Arr, 1) + 1, UBound(Arr, 2) - LBound(Arr, 2) + 1) = Arr

Exit Sub
SingleDimension:
Dim tempArr As Variant: Dim i As Long
ReDim tempArr(LBound(Arr, 1) To UBound(Arr, 1), 1 To 1)
For i = LBound(Arr, 1) To UBound(Arr, 1)
    tempArr(i, 1) = Arr(i)
Next
startRng.Cells(1, 1).Resize(UBound(Arr, 1) - LBound(Arr, 1) + 1, 1) = tempArr

End Sub

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

vaArr = Split(v, Cutter)(1)
If Not IsMissing(Trimmer) Then vaArr = Split(vaArr, Trimmer)(0)

Splitter = vaArr

End Function

Function ParseJSON(strJSON, strToParse, Optional strID, Optional strToRemove) As Variant

'###############################################################
'오빠두엑셀 VBA 사용자지정함수 (https://www.oppadu.com)
'▶ ParseJSON 함수
'▶ strJSON 데이터에서 선택한 데이터 값만 추출합니다.
'▶ 인수 설명
'_____________strJSON            : JSON 데이터입니다.
'_____________strToParse        : JSON 추출할 데이터 필드명입니다. 쉼표(,)로 구분하여 입력합니다.
'_____________strID                 : 추출한 데이터 배열의 열에 추가할 ID 입니다. (선택인수)
'_____________strToRemove   : 추출한 데이터에서 제거할 문자열입니다. 쉼표(,)로 구분하여 입력합니다.
'▶ 사용 예제
'Dim v As Variant
'v = ParseJSON(JsonData, "Date, Name, Item")
'###############################################################

'----------------------------------------------------
'변수 설정
'----------------------------------------------------
Dim vaToParse As Variant: Dim vToParse As Variant
Dim vaToRemove As Variant: Dim vToRemove As Variant
Dim lngStart As Long
Dim objItm As Variant: Dim strItm As String: Dim tmpItm As String
Dim itmCnt As Long
Dim i As Long: Dim r As Long: Dim c As Long
Dim dicItm As Object
Dim vaItm As Variant: Dim vaItems As Variant: Dim vaReturn As Variant
Dim iCol As Long: Dim maxCol As Long: Dim j As Long

Set dicItm = CreateObject("Scripting.Dictionary")

'----------------------------------------------------
'JSON 쿼리 분할
'----------------------------------------------------
vaToParse = Split(strToParse, ",")
If Not IsMissing(strToRemove) Then vaToRemove = Split(strToRemove, ",")

lngStart = InStr(1, strJSON, "[")
strJSON = Right(strJSON, Len(strJSON) - lngStart)

objItm = Split(strJSON, "{""")
itmCnt = UBound(objItm)

For i = 1 To itmCnt
    strItm = Split(objItm(i), """}")(0)
     iCol = Len(strItm) - Len(Replace(strItm, ":", ""))
    If iCol > maxCol Then maxCol = iCol
    
    If Not IsMissing(strID) Then
        ReDim vaItm(0 To iCol)
        vaItm(0) = strID: j = 1
    Else
        ReDim vaItm(0 To iCol - 1)
        j = 0
    End If
    
    On Error Resume Next
    For Each vToParse In vaToParse
        If InStr(strItm, Trim(vToParse) & """:") > 0 Then
            tmpItm = Split(strItm, Trim(vToParse) & """:")(1)
            tmpItm = Split(tmpItm, ",""")(0)
            If Left(tmpItm, 1) = """" Then tmpItm = Right(tmpItm, Len(tmpItm) - 1)
            If Right(tmpItm, 1) = """" Then tmpItm = Left(tmpItm, Len(tmpItm) - 1)
            tmpItm = Replace(tmpItm, "< ", "")
            If Not IsMissing(strToRemove) Then
                For Each vToRemove In vaToRemove
                    tmpItm = Replace(tmpItm, Trim(vToRemove), "")
                Next
            End If
            vaItm(j) = CStr(tmpItm)
        End If
        j = j + 1
    Next
    On Error GoTo 0
    dicItm.Add i, Array(vaItm, 1)
Next

'----------------------------------------------------
'Dictionary -> 배열 변환
'----------------------------------------------------
r = dicItm.Count
c = UBound(vaToParse) + 1
If Not IsMissing(strID) Then c = c + 1

If r = 0 Then ParseJSON = ""

vaItems = dicItm.Items

ReDim vaReturn(1 To r, 1 To c)

On Error Resume Next
For i = 0 To r - 1
    For j = 0 To c - 1
        tmpItm = vaItems(i)(0)(j)
        If IsNumeric(tmpItm) And Left(tmpItm, 1) <> 0 Then vaReturn(i + 1, j + 1) = CDbl(tmpItm) Else vaReturn(i + 1, j + 1) = tmpItm
    Next
Next
On Error GoTo 0

'----------------------------------------------------
'결과값 리턴
'----------------------------------------------------
ParseJSON = vaReturn

End Function
