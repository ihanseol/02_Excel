Attribute VB_Name = "usage_clipboard"

Dim Keys As New Selenium.Keys

Private Sub Paste_ClipBoard()
  Dim driver As New Selenium.FirefoxDriver
  driver.Get "https://en.wikipedia.org/wiki/Main_Page"
  
  driver.SetClipBoard [b3]
  driver.FindElementById("searchInput").SendKeys Keys.Control, "v"
  
  Debug.Assert 0
  driver.Quit
End Sub
