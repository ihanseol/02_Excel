Attribute VB_Name = "usage_send"

Private Sub Handle_Send()
  ' API: https://code.google.com/p/selenium/wiki/JsonWireProtocol

  Dim driver As New FirefoxDriver
  driver.Get "about:blank"
  
  ' Returns all windows handles
  Dim hwnds As List
  Set hwnds = driver.Send("GET", "/window_handles")
  
  ' Returns all links elements
  Dim links As List
  Set links = driver.Send("POST", "/elements", "using", "css selector", "value", "a")
  
  Debug.Assert 0
  driver.Quit
End Sub
