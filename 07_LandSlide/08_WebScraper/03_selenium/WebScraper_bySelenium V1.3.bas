Attribute VB_Name = "WebScraper_bySelenium"
Option Explicit

' microsoft internet controls
' microsoft HTML object Library
' sheet 30 has an error code
' because omitted cell error and i want to delete error


Public Function StringToIntArray(str As String) As Variant
    Dim temp As String, i As Long, L As Long
    Dim CH As String
    Dim wf As WorksheetFunction
    Set wf = Application.WorksheetFunction

    temp = ""
    L = Len(str)
    For i = 1 To L
        CH = Mid(str, i, 1)
        If CH Like "[0-9]" Then
            temp = temp & CH
        Else
            temp = temp & " "
        End If
    Next i

    StringToIntArray = Split(wf.Trim(temp), " ")

End Function

Public Function StringToDoubleArray(str As String) As Variant
    Dim wf As WorksheetFunction
    Set wf = Application.WorksheetFunction
    
    Dim trimString As String

    trimString = LTrim(RTrim(str))
   
    StringToDoubleArray = Split(trimString, vbLf)
End Function


Sub delete_ignore_error()
    
    Dim rg1 As Range
    
    For Each rg1 In Range("o6:o35")
            
            If rg1.Errors.Item(xlOmittedCells).Ignore = False Then
                rg1.Errors.Item(xlOmittedCells).Ignore = True
            End If
        
    Next rg1

    For Each rg1 In Range("o44:o53")
            
            If rg1.Errors.Item(xlOmittedCells).Ignore = False Then
                rg1.Errors.Item(xlOmittedCells).Ignore = True
            End If
        
    Next rg1

End Sub


Sub clear_30year_data()

    Range("b6:n35").ClearContents

End Sub


'find all form control, and select one radio button
'activex button control : msoOLEControlObject
'2019/11/20
'2020/3/24 : add areacode for etc

Function get_area_code() As Integer
    
    Dim vctrl As Shape
    Dim area_code As Integer
    Dim strA, strCaption
    
    For Each vctrl In ActiveSheet.Shapes
           
        If vctrl.Type = msoOLEControlObject Then
                With vctrl.DrawingObject.Object
                    strCaption = .Caption
                    If (.Value = True) Then
                            Select Case Left(.Caption, 1)
                                Case "e" 'etc
                                    get_area_code = Range("r14").Value
                                    Exit Function
                            
                                Case "D" 'DaeJeon
                                    get_area_code = 133
                                    Exit Function
                                    
                                Case "S" 'SeoSan
                                    get_area_code = 129
                                    Exit Function
                                    
                                Case "B"
                                    strA = Left(strCaption, 2)
                                    
                                    If strA = "Bo" Then
                                        get_area_code = 235  'Boryung
                                        Exit Function
                                    Else
                                        get_area_code = 236  'BuYeo
                                        Exit Function
                                    End If
                                        
                                Case "K" 'KeumSan
                                    get_area_code = 238
                                    Exit Function
                                    
                                Case "C" 'CheonAn
                                    get_area_code = 232
                                    Exit Function
                                
                                Case "H" 'HongSung
                                    get_area_code = 177
                                    Exit Function
                                
                                Case Else 'Default DaeJeon
                                    get_area_code = 133
                                    Exit Function
                            End Select
                    End If
                End With
            End If
            
    Next vctrl
    
End Function


'nArea = 235 -- 보령(무)
'nArea = 133 -- 대전(유)
'nArea = 129 -- 서산(무)

'in here is a Current Year


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
    
        Range("t7") = "Working " & j & " ---->  ( " & nYear & " )"
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
        Range("t7") = "Working " & j & " ---->  ( " & nYear & " )"
        
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
        Range("t7") = "Working " & j & " ---->  ( " & nYear & " )"
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


Function get_weather_data_bySeleniumXpath(ByVal nYear As Integer, ByVal nArea As Integer) As String()
 
    Dim i, j  As Long
        
    Dim bot As New ChromeDriver
    Dim td As Selenium.WebElement
    
    Dim url As String
    Dim SearchString As String
    Dim out(0 To 30, 0 To 12) As String
 
    nYear = nYear - 29
    
    For j = 0 To 29
        Range("t7") = "Working " & j & " ---->  ( " & nYear & " )"
        
        url = "https://www.weather.go.kr/weather/climate/past_table.jsp?stn=" & nArea & "&yy=" & nYear & "&obs=21&x=22&y=12"
        
        With bot
            .AddArgument "--headless"   ''This is the fix
            .Get url
        End With
        
        For i = 2 To 13
            SearchString = "//*[@id=""content_weather""]/table/tbody/tr[32]/td[" & i & "]"
            Set td = bot.FindElementByXPath(SearchString)
            out(j, i - 1) = CStr(td.text)
        Next i
        
        nYear = nYear + 1
    Next j
     
    get_weather_data_bySeleniumXpath = out
    
    bot.Close
    Set bot = Nothing
    
End Function

'in here is a Current Year

Function get_weather_data_byIE(ByVal nYear As Integer, ByVal nArea As Integer) As String()
 
    Dim ie As Object
    Dim strURL As String
    Dim i, j  As Integer
    Dim elmCollection As Object
    
    Dim out(0 To 30, 0 To 12) As String
    Set ie = CreateObject("InternetExplorer.Application")

    ie.Visible = False
    
    nYear = nYear - 29
    
    For j = 0 To 29
    
        Range("t7") = "Working " & j & " ---->  ( " & nYear & " )"
        ie.navigate "https://www.weather.go.kr/weather/climate/past_table.jsp?stn=" & nArea & "&yy=" & nYear & "&obs=21&x=22&y=12"
        
        Do While (ie.readyState <> READYSTATE_COMPLETE Or ie.Busy = True)
            DoEvents
        Loop
        
        Set elmCollection = ie.Document.getElementsByTagName("TABLE")
        
        For i = 0 To 12
            out(j, i) = elmCollection.Item(1).Rows.Item(32).all.Item(i).innerText
        Next i
        
        nYear = nYear + 1
    Next j
    
    
    get_weather_data_byIE = out
    
    ie.Quit
    Set ie = Nothing

End Function

'Search by XPath in selenium

Sub get_30year_data(nMethod As Integer)

    Dim resOut() As String
    Dim i, j, k  As Integer
    Dim nYear As Integer
    
    Dim nArea As Integer
       
    nYear = Year(Now()) - 1
    nArea = get_area_code()
    
    
    
    If (nMethod = 1) Then
        resOut = get_weather_data(nYear, nArea)
    Else
        resOut = get_weather_data_bySeleniumXpath(nYear, nArea)
    End If
     
    Application.ScreenUpdating = False
    
    For j = 0 To 29
        Cells(6 + j, 2) = nYear - 29 + j
        For i = 1 To 12
            Cells(6 + j, i + 2) = resOut(j, i)
        Next i
    Next j
    
    Call delete_ignore_error
    Application.ScreenUpdating = True
     
End Sub






