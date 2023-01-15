Attribute VB_Name = "WebScraper_bySelenium"
Option Explicit

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

Sub ChangeFormat()
    
    Dim lang_code As Integer
    Dim str_format As String

    lang_code = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    

    ' 1042 - korean
    ' 1033 - english
    
    If lang_code = 1042 Then
        str_format = "»¡°­"
    Else
         str_format = "Red"
    End If

    Range("B6:N35").Select
    Selection.NumberFormatLocal = "0_);[" & str_format & "](0)"
    Selection.NumberFormatLocal = "0.0_);[" & str_format & "](0.0)"

    Range("B6:B35").Select
    Selection.NumberFormatLocal = "0_);[" & str_format & "](0)"
    
End Sub




Sub clear_30year_data()
    Range("b6:n35").ClearContents
End Sub


Function get_area_code() As Integer
    get_area_code = Sheets("main").Range("local_code")
End Function


Function get_weather_data_bySeleniumXpath(ByVal nYear As Integer, ByVal nArea As Integer) As String()
 
    Dim i, j  As Long
        
    Dim bot As New ChromeDriver
    Dim td As Selenium.WebElement
    
    Dim url As String
    Dim SearchString As String
    Dim out(0 To 30, 0 To 12) As String
    
    Dim flag As Integer
 
    nYear = nYear - 29
    flag = 0
    
    For j = 0 To 29
        Range("u5") = "Working " & j & " ---->  ( " & nYear & " )"
        
        url = "https://www.weather.go.kr/weather/climate/past_table.jsp?stn=" & nArea & "&yy=" & nYear & "&obs=21&x=22&y=12"
        
        With bot
            .AddArgument "--headless"   ''This is the fix
            .Get url
        End With
        
        For i = 2 To 13
            SearchString = "//*[@id=""content_weather""]/table/tbody/tr[32]/td[" & i & "]"
            Set td = bot.FindElementByXPath(SearchString)
            out(j, i - 1) = CStr(td.text)
            
            If CStr(td.text) <> " " Then
                flag = flag + CInt(out(j, i - 1))
            End If
        Next i
        
        If flag = 0 Then
           MsgBox "Data is Empty ..."
           Exit For
        End If
        
        nYear = nYear + 1
        flag = False
        
    Next j
     
    get_weather_data_bySeleniumXpath = out
    
    bot.Close
    Set bot = Nothing
    
End Function


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
    Call ChangeFormat
    Application.ScreenUpdating = True
     
End Sub



