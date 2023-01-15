Attribute VB_Name = "usage_http"
Private Assert As New Selenium.Assert

Sub Test_Native_HttpRequest()
  ' https://msdn.microsoft.com/en-us/library/ms535874(v=vs.85).aspx
  
  Set http = CreateObject("MSXML2.XMLHTTP")
  http.Open "GET", "https://vortex.data.microsoft.com/collect/v1", False
  http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 6.3; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/47.0.2526.80 Safari/537.36"
  http.setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
  http.Send ""
  Debug.Print http.responseText
End Sub

Sub Test_HttpRequest()
  'https://developer.mozilla.org/fr/docs/Web/API/XMLHttpRequest
  
  Dim drv As New Selenium.FirefoxDriver
  drv.Get "https://vortex.data.microsoft.com"
  
  Const JS_HttpRequest As String = _
    "var r = new XMLHttpRequest();" & _
    "r.open('GET', arguments[0], 0);" & _
    "r.send();" & _
    "return JSON.parse(r.responseText);"
    
  Set result = drv.ExecuteScript(JS_HttpRequest, "https://vortex.data.microsoft.com/collect/v1")
  Assert.Equals 1, result("acc")
  drv.Quit
End Sub

