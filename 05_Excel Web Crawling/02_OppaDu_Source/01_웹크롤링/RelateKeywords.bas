Attribute VB_Name = "RelateKeywords"
Sub deleteSearchKeywords()

    Sheet2.Range("f2") = ""

End Sub


Sub SearchRelateKeywords()

Dim url As String
Dim htmlResult As Object
Dim strResult As String
Dim strSearch As String

Dim v As Variant
Dim i As Long



'strSearch = "화수"
strSearch = Sheet2.Range("f2")
If Not HasContent(strSearch) Then strSearch = "사랑"

url = "https://ac.search.naver.com/nx/ac?q=" & strSearch & "&frm=nv&st=100"

Set htmlResult = GetHttp(url)

strResult = htmlResult.body.innerHTML

'ExportText strResult

'splitter - in json format
'{ "query" : ["화수"], "items" : [ [["화수분"],["화순군청"],["화순"],["화순 코로나"],["화순 맛집"],["화순날씨"],["화순전남대병원"],["화순 요양병원"],["화순 자연의미학"],["화순 요양병원 코로나"]] ] }
'start : <"items" : [ [> , end : <] ] }>
'잘라낸다 시작과 끝을 제이슨 포맷스트링의

Debug.Print strResult
strResult = Splitter(strResult, """items"" : [ [", "] ] }")

v = Split(strResult, ",")


On Error GoTo EmptyStrResult
For i = LBound(v) To UBound(v)
    v(i) = Replace(v(i), "[""", "")
    v(i) = Replace(v(i), """]", "")
    Debug.Print (v(i))
Next i


ArrayToRng Sheet2.Range("A1"), v

EmptyStrResult:

End Sub



