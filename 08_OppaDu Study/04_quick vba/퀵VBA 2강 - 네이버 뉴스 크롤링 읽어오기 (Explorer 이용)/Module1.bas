Attribute VB_Name = "Module1"
Option Explicit

Sub Web_scarping()

Dim MyBrowser As InternetExplorer
Dim HTMLDoc As HTMLDocument
Dim iArticle As IHTMLElement
Dim i As Long

'// 익스플로어(XML변수 설정)
Set MyBrowser = Sheet1.WebBrowser1

'// 검색 (XML 요청)
With MyBrowser
    .Navigate Sheet1.Range("C3").Value
    
    Wait_Browser MyBrowser

    '// HTML 추출
    Set HTMLDoc = .Document
    
    i = 7
    
    '// 개체검색
    For Each iArticle In HTMLDoc.getElementsByClassName("_sp_each_title")
        '// 요소별 값 추출
        Sheet1.Cells(i, 2) = iArticle.Title
        Sheet1.Cells(i, 3) = iArticle.getAttribute("href")
        Sheet1.Hyperlinks.Add Sheet1.Cells(i, 3), Sheet1.Cells(i, 3)
        i = i + 1
    Next
    
End With

MsgBox "검색하신 단어 [" & Sheet1.Range("C2") & "] 에 대한 네이버 뉴스기사 스크랩핑을 완료하였습니다."


End Sub

Sub Wait_Browser(Browser As InternetExplorer, Optional t As Integer = 1)

While Browser.Busy Or Browser.ReadyState <> 4
DoEvents
Wend

Application.Wait DateAdd("s", t, Now)

End Sub
