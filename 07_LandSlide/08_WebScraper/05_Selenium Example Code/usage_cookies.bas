Attribute VB_Name = "usage_cookies"

Private Assert As New Selenium.Assert


Private Sub Handle_Cookies()
  Dim driver As New FirefoxDriver
  driver.Get "http://admin:admin@the-internet.herokuapp.com/download_secure"
  
  'Get a cookie by name
  Dim cookie As cookie
  Set cookie = driver.Manage.FindCookieByName("rack.session")
  Assert.Equals "the-internet.herokuapp.com", cookie.Domain
  
  'Stop the browser
  driver.Quit
End Sub
