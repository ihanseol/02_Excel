Attribute VB_Name = "WebScrapingExam"
Option Explicit

Function RemoveHTML(text As String) As String
    Dim regexObject As Object
    Set regexObject = CreateObject("vbscript.regexp")

    With regexObject
        .Pattern = "<!*[^<>]*>"    'html tags and comments
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
    End With

    RemoveHTML = regexObject.Replace(text, "")
End Function


Sub getjuso()

    'source :    https://youtu.be/Y5W-_j_ZnoA?list=PL4NqC0DTQ1yRaGTFKPxiYtp6yMWkmqv2E

    Dim ie As InternetExplorer
    Dim strURL As String
    Dim i, nMax As Integer
    Dim elm As Object
    
    
    Set ie = New InternetExplorer
    strURL = "http://www.juso.go.kr/support/AddressMainSearch.do?searchType=TOTAL"
    
    ie.navigate strURL
    ie.Visible = True
    
    
    Do While (ie.readyState <> READYSTATE_COMPLETE Or ie.Busy = True)
        DoEvents
    Loop
    
    ie.Document.getElementsByName("searchKeyword")(0).value = "½ÅºÀ1·Î 27"
    
    ' this script input
    Call ie.Document.parentWindow.execScript("headerSearch('seachList');", "JavaScript")
    
    Do While (ie.readyState <> READYSTATE_COMPLETE Or ie.Busy = True)
        DoEvents
    Loop
    
      
    Set elm = ie.Document.getElementsByClassName("list")
    
    For i = 0 To elm.Length
        If (elm(i).parentElement.parentElement.parentElement.className <> "section-search") Then
            nMax = i
            Exit For
        End If
    Next i
    
    For i = 1 To nMax
        Debug.Print RemoveHTML(ie.Document.getElementById("rnAddr" & i).value)
        Debug.Print ie.Document.getElementById("bsiZonNo" & i).value
    Next i
    
    ie.Quit
    Set ie = Nothing

End Sub




Sub exam_getcurrency()

    Dim ie As InternetExplorer
    Dim strURL As String
    Dim i As Integer
    
    strURL = "https://www1.oanda.com/currency/converter/"
    
    Set ie = CreateObject("InternetExplorer.Application")

    ie.navigate strURL
    
    ie.Visible = True
    
    Do While (ie.readyState <> READYSTATE_COMPLETE Or ie.Busy = True)
        DoEvents
    Loop

    ' left currency country
    ie.Document.getElementById("form_quote_currency_hidden").value = "USD"
    
    ' right right country
    ie.Document.getElementById("form_base_currency_hidden").value = "EUR"
    
    ie.Document.getElementById("quote_amount_input").value = "1.2"
    
    'form_end_date_hidden
    ie.Document.getElementById("form_end_date_hidden").value = "2019-11-18"
    
    'flipper
    ie.Document.getElementById("flipper").Click
    
    Range("r8") = ie.Document.getElementById("bidAskAskAvg").innerHTML
    
    Application.Wait (Now + TimeValue("00:00:01"))
    
    
    
    ie.Quit
    Set ie = Nothing
    

End Sub

