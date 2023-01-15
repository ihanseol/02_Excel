Attribute VB_Name = "usage_frame"
Private Assert As New Selenium.Assert
Private Keys As New Selenium.Keys


Private Sub Handle_Frames()
  Dim driver As New FirefoxDriver
  driver.Get "http://the-internet.herokuapp.com/nested_frames"
  
  'switch to a child frame
  driver.SwitchToFrame "frame-top"
  Assert.Equals 3, driver.FindElementsByTag("frame").count
  
  'switch to child frame of "frame-top"
  driver.SwitchToFrame "frame-middle"
  Assert.Equals "MIDDLE", driver.FindElementById("content").Text
  
  'switch to the default content
  driver.SwitchToDefaultContent
  Assert.Equals 2, driver.FindElementsByTag("frame").count
  
  'Stop the browser
  driver.Quit
End Sub
