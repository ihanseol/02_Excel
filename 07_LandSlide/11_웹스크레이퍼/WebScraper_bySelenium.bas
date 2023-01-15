Attribute VB_Name = "WebScraper_bySelenium"
Option Explicit

Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
Private cd As Selenium.ChromeDriver


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
        str_format = "빨강"
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





Sub get_weather_data()
    Dim cd As New ChromeDriver
    Dim ddl As Selenium.SelectElement
    
    Dim url As String
    Dim one_string, two_string As String
    Dim sYear, eYear As Integer
    Dim str As String
    
    Range("B2").Value = "30년 " & Range("S8").Value & "데이터, " & Now()
    
    url = "https://data.kma.go.kr/stcs/grnd/grndRnList.do?pgmNo=69"
    Set cd = New Selenium.ChromeDriver
    
    cd.Start
    cd.AddArgument "--headless"
    cd.Window.Maximize
    cd.Get url

    Sleep (1 * 1000)
    

    one_string = "ztree_" & CStr(Range("S10").Value) & "_switch"
    two_string = Range("S8").Value & " (" & CStr(Range("S9").Value) & ")"
    
    Set ddl = cd.FindElementByCss("#dataFormCd").AsSelect
    ddl.SelectByText ("월")
    Sleep (0.5 * 1000)
    
    
    ' ---------------------------------------------------------------
    
    cd.FindElementByCss("#txtStnNm").Click
    Sleep (1 * 1000)
    cd.FindElementByCss("#" & one_string).Click
    Sleep (1 * 1000)
    cd.FindElementByLinkText(two_string).Click
    Sleep (1 * 1000)
    cd.FindElementByLinkText("선택완료").Click
    
    
    ' ---------------------------------------------------------------
    eYear = Year(Now()) - 1
    sYear = eYear - 29
    
    Set ddl = cd.FindElementByCss("#startYear").AsSelect
    ddl.SelectByText (CStr(sYear))
    Sleep (0.5 * 1000)
   
    Set ddl = cd.FindElementByCss("#endYear").AsSelect
    ddl.SelectByText (CStr(eYear))
    Sleep (0.5 * 1000)
    ' ---------------------------------------------------------------
    
    ' Search Button
    ' cd.FindElementByXPath("//*[@id='schForm']/div[2]").Click
    ' copy by selector
    
    cd.FindElementByCss("#schForm > div.wrap_btn > button").Click
    Sleep (1 * 1000)
    
    ' Excel download button
    cd.FindElementByLinkText("Excel").Click
    Sleep (3 * 1000)

    
    
End Sub



