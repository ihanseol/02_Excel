Attribute VB_Name = "NaverModule"
Option Explicit


Sub NaverLandCrawl()

'##########################################################
' 변수 설정
Dim htmlResult As Object
Dim strResult As String
Dim url As String
Dim city As String


'##########################################################
'1. 네이버 부동산 메인 페이지 -> 지역 검색
'                                               -> 가용 매물 목록을 받아오기 위한 위/경도, 그 외 변수 추출
'[            https://meyerweb.com/eric/tools/dencoder/                ]  '-> URL 디코딩/인코딩
'[            http://json.parser.online.fr/                                             ]  '-> JSON 파서

city = Sheet1.Range("D5")
url = "https://m.land.naver.com/search/result/" & city

'Debug.Print url

Set htmlResult = GetHttp(url)
strResult = htmlResult.body.innerHTML

strResult = Splitter(strResult, "filter: {", "},")

'Debug.Print strResult
'ExportText strResult


'----------------------------------------
' 매물유형 예) 아파트: APT, 빌라 : VL, 오피스텔: OPST, ...적절히 수정
' 건물유형 A1, B1, ... 적절히 수정가능
'----------------------------------------



'---------------- 해당 지역(메인) 위/경도 + 그 외 변수 생성
Dim lat As String: lat = Splitter(strResult, "lat: '", "',")
Dim lon As String: lon = Splitter(strResult, "lon: '", "',")
Dim z As String: z = Splitter(strResult, "z: '", "',")
Dim cortarNo As String: cortarNo = Splitter(strResult, "cortarNo: '", "',")
Dim searchType As String: searchType = "APT"
Dim buildingType As String: buildingType = "A1:B1:B2"

'#########################################################
' 2. 네이버 부동산 지도 페이지 -> 받아온 위/경도, 그 외 변수로 검색
'                                                -> 해당 지역 가용매물 목록의 상세 위/경도 추출


url = "https://m.land.naver.com/cluster/clusterList?view=actl&cortarNo=" & cortarNo & "&rletTpCd=" & searchType & "&tradTpCd=" & buildingType & _
            "&z=" & z & "&lat=" & lat & "&lon=" & lon & "&addon=COMPLEX&bAddon=COMPLEX&isOnlyIsale=false"

Set htmlResult = Nothing
Set htmlResult = GetHttp(url)
strResult = htmlResult.body.innerHTML

Dim forceCOMPLEX As Boolean
If InStr(1, strResult, "COMPLEX") > 0 Then strResult = Splitter(strResult, "COMPLEX"): forceCOMPLEX = True
Dim v As Variant
v = ParseJSON(strResult, "lgeo,lat,lon,count")

'##########################################################
' 3. 네이버 부동산 매물페이지  -> 상세 위 경도로 매물 정보 검색
'                                                -> 각 매물별 상세정보 추출
' 매물페이지는 페이지당 20개씩만 출력.. 그래서 추가 작업 필요!
' Set htmlResult = Nothing
' Set htmlResult = GetHttp(URL)
'                strResult = htmlResult.body.innerHTML
'                vReturn = ParseJSON(strResult, "hscpNm,hscpNo,scpTypeCd,hscpTypeNm,totDongCnt,totHsehCnt,genHsehCnt,useAprvYmd,repImgUrl,dealCnt,leaseCnt,rentCnt," & _
                                             "strmRentCnt,totalAtclCnt,minSpc,maxSpc,dealPrcMin,dealPrcMax,leasePrcMin,leasePrcMax,isalePrcMin,isalePrcMax,isaleNotifSeq,isaleScheLabel,isaleScheLabelPre", city, ",")

Dim i As Long: Dim iPage As Long: Dim j As Long
Dim vReturn As Variant
Dim x As Long: x = GetLastRow(Sheet1) + 1
Dim initR As Long
initR = x

For i = LBound(v, 1) To UBound(v, 1)
    If v(i, 1) <> "" Then
        iPage = Application.WorksheetFunction.RoundUp(v(i, 4) / 20, 0)
        For j = 1 To iPage
            If forceCOMPLEX = False Then
                url = "https://m.land.naver.com/cluster/ajax/articleList?itemId=" & v(i, 1) & "&lgeo=" & v(i, 1) & _
                                "&rletTpCd=" & searchType & "&tradTpCd=" & buildingType & "&z=" & z & "&lat=" & v(i, 2) & "&lon=" & v(i, 3) & "&cortarNo=" & cortarNo & _
                                "&isOnlyIsale=false&sort=readRank&page=" & j
                Debug.Print url
                Set htmlResult = Nothing
                Set htmlResult = GetHttp(url)
                strResult = htmlResult.body.innerHTML
                If InStr(1, strResult, "atclNo") > 0 Then
                    vReturn = ParseJSON(strResult, "atclNm,atclNo,tradTpCd,rletTpNm,totDongCnt_tmp,totHsehCnt_tmp,genHsehCnt_tmp,atclCfmYmd,repImgUrl,tradTpNm,flrInfo,atclTetrDesc," & _
                                    "strmRentCnt_tmp,totalAtclCnt_tmp,spc1,spc2,sameAddrMinPrc,sameAddrMaxPrc,minMviFee,maxMviFee,cpid,cpNm,rltrNm,isaleScheLabel_tmp,isaleScheLabelPre_tmp", city, ",")
                    ArrayToRng Sheet1.Cells(x, 4), vReturn
                    x = x + UBound(vReturn, 1)
                End If
            Else
                url = "https://m.land.naver.com/cluster/ajax/complexList?itemId=" & v(i, 1) & "&lgeo=" & v(i, 1) & _
                        "&rletTpCd=" & searchType & "&tradTpCd=" & buildingType & "&z=" & z & "&lat=" & v(i, 2) & "&lon=" & v(i, 3) & "&cortarNo=" & cortarNo & "&isOnlyIsale=false&sort=readRank&page=" & j
                Set htmlResult = Nothing
                Set htmlResult = GetHttp(url)
                strResult = htmlResult.body.innerHTML
                If InStr(1, strResult, "hscpNo") > 0 Then
                    vReturn = ParseJSON(strResult, "hscpNm,hscpNo,scpTypeCd,hscpTypeNm,totDongCnt,totHsehCnt,genHsehCnt,useAprvYmd,repImgUrl,dealCnt,leaseCnt,rentCnt," & _
                                    "strmRentCnt,totalAtclCnt,minSpc,maxSpc,dealPrcMin,dealPrcMax,leasePrcMin,leasePrcMax,isalePrcMin,isalePrcMax,isaleNotifSeq,isaleScheLabel,isaleScheLabelPre", city, ",")
                    ArrayToRng Sheet1.Cells(x, 4), vReturn
                    x = x + UBound(vReturn, 1)
                End If
            End If
        Next j
    End If
Next i

'#########################################################
' 4. 대표 이미지 삽입

'    Dim shpImg As Shape: Dim shpRng As Range
'    For j = initR To Sheet1.Cells(Sheet1.Rows.Count, 4).End(xlUp).Row + 1
'        Set shpRng = Sheet1.Cells(j, 13)
'        If shpRng.Value <> "" Then
'            shpRng.EntireRow.RowHeight = 80
'            InsertWebImage shpRng, "https://landthumb-phinf.pstatic.net" & shpRng.Value
'        Else
'            shpRng.EntireRow.RowHeight = 18
'        End If
'    Next
'

End Sub



