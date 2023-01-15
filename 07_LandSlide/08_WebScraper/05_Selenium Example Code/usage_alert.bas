Attribute VB_Name = "usage_alert"
Private Assert As New Selenium.Assert


Private Sub Handle_Alerts()
  Dim driver As New FirefoxDriver
  driver.SetCapability "unexpectedAlertBehaviour", "ignore"
  driver.Get "http://the-internet.herokuapp.com/javascript_alerts"
  
  ' Display alert
  driver.FindElementByCss("#content li:nth-child(2) button").Click
  
  ' Set the context on the alert dialog
  Set dlg = driver.SwitchToAlert(Raise:=False)
  
  ' Assert an alert is present and the message
  Assert.False dlg Is Nothing, "No alert present!"
  Assert.Equals "I am a JS Confirm", dlg.Text
  
  ' Close alert
  dlg.Accept
  
  driver.Quit
End Sub


'Returns true if an alert is present, false otherwise
' driver: web driver
Private Function IsDialogPresent(driver As WebDriver) As Boolean
  On Error Resume Next
  T = driver.title
  IsDialogPresent = (26 = Err.Number)
End Function
