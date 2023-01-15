Attribute VB_Name = "GPS"
Option Explicit
Option Base 0


'<error_message> You must use an API key to authenticate each request to Google Maps Platform APIs.
'For additional information, please refer to http://g.co/dev/maps-no-account</error_message>
'
' source - https://rakeion.blog.me/221033846320

Function Gf_GeoAddress(in_val As String) As String
    Dim str_위도 As String
    Dim str_경도 As String
    Dim str As String

    With CreateObject("WinHttp.WinHttpRequest.5.1")
        .Open "GET", "http://maps.googleapis.com/maps/api/geocode/xml?" & "address=" & in_val & "&sensor=false"
        .send
        .WaitForResponse: DoEvents
        str = .responsetext
        
        Debug.Print str
        
        str_위도 = Split(Split(str, "<lat>")(1), "</lat>")(0) '위도
        str_경도 = Split(Split(str, "<lng>")(1), "</lng>")(0) '경도
    End With
    
    Gf_GeoAddress = str_위도 + ", " + str_경도
End Function

Sub test()

    Debug.Print Gf_GeoAddress("대전시 유성구 장대동 278-13")

End Sub


Function Gf_GeoDistance(in_val As String, in_val2 As String) As String
    Dim str1() As String
    Dim str2() As String
    
    Dim dbl1(2) As Double
    Dim dbl2(2) As Double
    
    
    str1 = Split(in_val, ",")
    str2 = Split(in_val2, ",")
    
    dbl1(0) = CDbl(str1(0))
    dbl1(1) = CDbl(str1(1))
    dbl2(0) = CDbl(str2(0))
    dbl2(1) = CDbl(str2(1))
    
    Gf_GeoDistance = Acos(Cos(Radians(90 - dbl1(0))) * Cos(Radians(90 - dbl2(0))) + Sin(Radians(90 - dbl1(0))) * Sin(Radians(90 - dbl2(0))) * Cos(Radians(dbl1(1) - dbl2(1)))) * 6371
End Function


Function Acos(in_val)
    Acos = Application.WorksheetFunction.Acos(in_val)
End Function
    
Function Radians(in_val)
    Radians = Application.WorksheetFunction.Radians(in_val)
End Function






