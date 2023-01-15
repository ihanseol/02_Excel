Attribute VB_Name = "usage_window"
Private Assert As New Selenium.Assert

Private Sub Windows_Switch()
  Dim driver As New FirefoxDriver
  
  driver.Get "http://the-internet.herokuapp.com/windows"
  driver.FindElementByCss(".example a").Click
  
  'Switch to a new window
  driver.SwitchToNextWindow
  Assert.Equals "New Window", driver.title
  
  'Switch to the previous activated window
  driver.SwitchToPreviousWindow
  Assert.NotEquals "New Window", driver.title
  
  'Stop the browser
  driver.Quit
End Sub

Private Sub Window_Maximize()
  Dim driver As New FirefoxDriver
  driver.Get "about:blank"
  driver.Window.Maximize
  
  Debug.Assert 0
  driver.Quit
End Sub

Private Sub Window_SetSize()
  Dim driver As New FirefoxDriver
  driver.Get "about:blank"
  driver.Window.SetSize 800, 600
  
  Debug.Assert 0
  driver.Quit
End Sub

Private Sub Windows_Close()
  Dim driver As New FirefoxDriver
  
  driver.Get "http://the-internet.herokuapp.com/windows"
  
  ' Save the main window
  Set winMain = driver.Window
  
  ' Open a new window
  driver.FindElementByCss(".example a").Click
  
  ' Close all the newly opened windows
  For Each win In driver.Windows
    If Not win.Equals(winMain) Then win.Close
  Next
  winMain.Activate
  
  'Stop the browser
  driver.Quit
End Sub


Private Sub Windows_Open_New_Tab()
  Dim driver As New ChromeDriver, Keys As New Keys
  
  driver.Get "http://the-internet.herokuapp.com/windows"
  
  ' Holds the control key while clicking
  driver.FindElementByLinkText("Dropdown").Click Keys.Control
  driver.SwitchToNextWindow
  
  'Stop the browser
  driver.Quit
End Sub


Private Sub Windows_Open_New_One2()
  Dim driver As New ChromeDriver, Keys As New Keys
  
  driver.Get "about:blank"
  
  driver.ExecuteScript "window.open(arguments[0])", "http://www.google.com/"
  driver.SwitchToNextWindow
  
  'Stop the browser
  driver.Quit
End Sub
