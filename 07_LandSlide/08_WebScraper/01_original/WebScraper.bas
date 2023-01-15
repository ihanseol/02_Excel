Attribute VB_Name = "WebScraper"
Option Explicit

' microsoft internet controls
' microsoft HTML object Library


Sub Extract_TD_text()

    Dim url As String
    Dim ie As InternetExplorer
    Dim HTMLdoc As HTMLDocument
    Dim TDelements As IHTMLElementCollection
    Dim TDelement As HTMLTableCell
    Dim r As Long
    
    'Saved from www vbaexpress com/forum/forumdisplay.php?f=17
    url = "file://C:\VBAExpress_Excel_Forum.html"
    
    Set ie = New InternetExplorer
    
    With ie
        .navigate url
        .Visible = True
    
        'Wait for page to load
        While .Busy Or .readyState <> READYSTATE_COMPLETE: DoEvents: Wend
    
        Set HTMLdoc = .Document
    End With
    
    Set TDelements = HTMLdoc.getElementsByTagName("TD")
    
    Sheet1.Cells.ClearContents
    
    r = 0
    For Each TDelement In TDelements
        'Look for required TD elements - this check is specific to VBA Express forum - modify as required
        If TDelement.className = "alt2" And TDelement.Align = "center" Then
            Sheet1.Range("A1").Offset(r, 0).value = TDelement.innerText
            r = r + 1
        End If
    Next
            
End Sub



'sheet 30 has an error code
'because omitted cell error and i want to delete error

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
                    If (.value = True) Then
                            Select Case Left(.Caption, 1)
                                Case "e" 'etc
                                    get_area_code = Range("r14").value
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


Function get_weather_data(ByVal nYear As Integer, ByVal nArea As Integer) As String()
 
    Dim ie As Object
    Dim strURL As String
    Dim i As Integer
    Dim elmCollection As Object
    
    Dim out(0 To 12) As String
 
    strURL = "https://www.weather.go.kr/weather/climate/past_table.jsp?stn=" & CStr(nArea) & "&yy=" & CStr(nYear) & "&obs=21&x=22&y=12"
    Set ie = CreateObject("InternetExplorer.Application")

    ie.navigate strURL
    ie.Visible = False
    
    Do While (ie.readyState <> READYSTATE_COMPLETE Or ie.Busy = True)
        DoEvents
    Loop
        
    Set elmCollection = ie.Document.getElementsByTagName("TABLE")
    
    For i = 0 To 12
        out(i) = elmCollection.Item(1).Rows.Item(32).all.Item(i).innerText
    Next i
     
    get_weather_data = out

    ie.Quit
    Set ie = Nothing

End Function



'in here is a Current Year

Function get_weather_data2(ByVal nYear As Integer, ByVal nArea As Integer) As String()
 
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
    
    
    get_weather_data2 = out
    
    ie.Quit
    Set ie = Nothing

End Function



' Source - https://stackoverflow.com/questions/57633330/excel-vba-ie-automation-cant-click-on-a-button-that-already-gets-focused-corr

Public Sub test()

    Dim ie As InternetExplorer, url As String, evt As Object
    
    url = "http://finra-markets.morningstar.com/MarketData/Default.jsp?sdkVersion=2.37.0"
    Set ie = New InternetExplorer
    
    With ie
        .Visible = True
        .Navigate2 url
        While .Busy Or .readyState <> 4: DoEvents: Wend
    
        With .Document
            .querySelector("#ms-finra-autocomplete-box").value = "ZUAN.GA"
            Set evt = .createEvent("HTMLEvents")
            evt.initEvent "mousedown", True, False
            .querySelector("[value=GO]").dispatchEvent evt
        End With
    
        Do While .Document.url = url: DoEvents: Loop
    
        MsgBox .Document.url
    End With

End Sub



Sub test_weather_data()

    Dim ie As Object
    Dim strURL As String
    Dim i, j  As Integer
    Dim elm As Object
    Dim nArea, nYear As Integer: nArea = 133: nYear = 2010
    Dim out(0 To 30, 0 To 12) As String
    
    
    Set ie = CreateObject("InternetExplorer.Application")

    ie.Visible = True
    
    
    ie.navigate "https://www.weather.go.kr/weather/climate/past_table.jsp?stn=" & nArea & "&yy=" & nYear & "&obs=21&x=22&y=12"
        
    Application.Wait Now + TimeSerial(0, 0, 3)
        
    Do While (ie.readyState <> READYSTATE_COMPLETE Or ie.Busy = True)
        DoEvents
    Loop
    
    Set elm = ie.Document.getElementById("observation_select2")
    elm.value = "1988"
     
    Set elm = ie.Document.getElementsByName("viewform")
 
   
    ie.Quit
    Set ie = Nothing

End Sub

Sub weather_test()

    Dim resOut() As String
    Dim i, j, k  As Integer
    Dim nArea As Integer
    
    
    Sheet13.OptionButton3.value = True
    
    nArea = get_area_code()
    
    For k = 0 To 29
    
        resOut = get_weather_data(1989 + k, nArea)
        For i = 1 To 12
            Cells(6 + k, i + 2) = resOut(i)
        Next i
    
    Next k
    
    For i = 1 To 12
        Debug.Print resOut(i)
    Next i

End Sub



Sub get_30year_data()

    Dim resOut() As String
    Dim i, j, k  As Integer
    Dim nYear As Integer
    
    Dim nArea As Integer
       
    nYear = Year(Now()) - 1
    nArea = get_area_code()
    
    resOut = get_weather_data2(nYear, nArea)
    
     
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





