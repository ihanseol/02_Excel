Attribute VB_Name = "mod_Selenium"
Option Explicit


Function get_weather_data(ByVal nYear As Integer, ByVal nArea As Integer) As String()
 
    Dim i, j  As Integer
        
    Dim htmlResult As Object
    Dim strResult As String
    Dim strTable As String
    Dim url As String
    Dim arr As Variant
    
    Dim out(0 To 30, 0 To 12) As String
    
    
    nYear = nYear - 29
    
    For j = 0 To 29
    
        Range("u5") = "Working " & j & " ---->  ( " & nYear & " )"
        url = "https://www.weather.go.kr/weather/climate/past_table.jsp?stn=" & nArea & "&yy=" & nYear & "&obs=21&x=22&y=12"
        
        Set htmlResult = GetHttp(url)
        strResult = htmlResult.body.innerHTML
        'ExportText strResult
                            
        strResult = Splitter(strResult, "<TD scope=row>합계</TD>", "</TR>")
        strTable = RemoveHTML(strResult)
               
        arr = StringToDoubleArray(strTable)
        
        For i = 1 To UBound(arr)
            out(j, i) = CDbl(arr(i))
        Next i
              
        nYear = nYear + 1
    Next j
     
    get_weather_data = out
    
End Function


' 2021/7/20 parsing table by selenium
' https://youtu.be/lr7CFZEI2YA

Function get_weather_data_bySelenium(ByVal nYear As Integer, ByVal nArea As Integer) As String()
 
    Dim i, j As Long
        
    Dim bot As New ChromeDriver
    Dim td  As Selenium.WebElement
    Dim trs As Selenium.WebElements
    Dim tds As Selenium.WebElements
    
    Dim url As String
    
    Dim out(0 To 30, 0 To 12) As String
 
    nYear = nYear - 29
    
    For j = 0 To 29
        Range("u5") = "Working " & j & " ---->  ( " & nYear & " )"
        
        url = "https://www.weather.go.kr/weather/climate/past_table.jsp?stn=" & nArea & "&yy=" & nYear & "&obs=21&x=22&y=12"
        
        With bot
            .AddArgument "--headless"   ''This is the fix
            .Get url
        End With
        
       Set trs = bot.FindElementByClass("table_develop").FindElementByTag("tbody").FindElementsByCss("tr")
       Set tds = trs(32).FindElementsByCss("td")
                   
        i = 1
        For Each td In tds
            If td.text <> "합계" Then
                out(j, i) = CStr(td.text)
                i = i + 1
            End If
        Next td
        
        nYear = nYear + 1
    Next j
     
    get_weather_data_bySelenium = out
    
    bot.Close
    Set bot = Nothing
    
End Function


Function get_weather_data_bySeleniumV2(ByVal nYear As Integer, ByVal nArea As Integer) As String()
 
    Dim i, j As Long
        
    Dim bot As New ChromeDriver
    Dim trs As Selenium.WebElements
    Dim tds As Selenium.WebElements
    
    Dim url As String
    Dim out(0 To 30, 0 To 12) As String
 
    nYear = nYear - 29
    
    For j = 0 To 29
        Range("u5") = "Working " & j & " ---->  ( " & nYear & " )"
        url = "https://www.weather.go.kr/weather/climate/past_table.jsp?stn=" & nArea & "&yy=" & nYear & "&obs=21&x=22&y=12"
        
        With bot
            .AddArgument "--headless"   ''This is the fix
            .Get url
        End With
        
       Set trs = bot.FindElementByClass("table_develop").FindElementByTag("tbody").FindElementsByCss("tr")
       Set tds = trs(32).FindElementsByCss("td")
                           
        For i = 2 To 13
            out(j, i - 1) = CStr(tds(i).text)
        Next
        
        nYear = nYear + 1
    Next j
     
    get_weather_data_bySeleniumV2 = out
    
    bot.Close
    Set bot = Nothing
    
End Function




