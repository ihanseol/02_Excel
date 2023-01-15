Attribute VB_Name = "usage_get"

Dim Assert As New Selenium.Assert

Private Sub GetWebPage_Url()
  Dim driver As New FirefoxDriver
  
  ' get the main page
  driver.Get "https://www.google.co.uk"
  Assert.Equals "https://www.google.co.uk/", driver.URL
  
  ' get a sub page
  driver.Get "/intl/en/about/"
  Assert.Equals "https://www.google.co.uk/intl/en/about/", driver.URL
  
  ' get another sub page
  driver.baseUrl = "https://www.google.co.uk/intl/en"
  driver.Get "/policies/privacy"
  Assert.Equals "https://www.google.co.uk/intl/en/policies/privacy/", driver.URL
  
  driver.Quit
End Sub

Private Sub GetWebPage_File()
  Const html As String = _
    "<!DOCTYPE html>" & _
    "<html lang=""en"">" & _
    "<head><title>My title</title></head>" & _
    "<body><h1>My content</h1></body>" & _
    "</html>"
  
  ' Create the html file
  Dim file$
  file = Environ("TEMP") & "\mypage.html"
  Open file For Output As #1
    Print #1, html
  Close #1
  
  ' Open it
  Dim drv As New FirefoxDriver
  drv.Get file
  
  Debug.Assert 0
  drv.Quit
End Sub

Sub GetWebPage_DataScheme()
  Const html As String = _
    "data:text/html;charset=utf-8," & _
    "<!DOCTYPE html>" & _
    "<html lang=""en"">" & _
    "<head><title>My title</title></head>" & _
    "<body><h1>My content</h1></body>" & _
    "</html>"
  
  Dim drv As New FirefoxDriver
  drv.Get html
  
  Debug.Assert 0
  drv.Quit
End Sub

Sub GetWebPage_Javascript()
  Const html As String = _
    "<!DOCTYPE html>" & _
    "<html lang=""en"">" & _
    "<head><title>My title</title></head>" & _
    "<body><h1>My content</h1></body>" & _
    "</html>"
  
  Const JS_WRITEPAGE = _
    "var txt=arguments[0];" & _
    "setTimeout(function(){" & _
    " document.open();" & _
    " document.write(txt);" & _
    " document.close();" & _
    "}, 0);"
  
  Dim drv As New FirefoxDriver
  drv.Get "about:blank"
  drv.ExecuteScript JS_WRITEPAGE, html
  
  Debug.Assert 0
  drv.Quit
End Sub

