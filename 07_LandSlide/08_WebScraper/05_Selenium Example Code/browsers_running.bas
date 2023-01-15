Attribute VB_Name = "browsers_running"
' This module shows how to work with a running instance of
' a driver by using the GetObject function.
'
' To do so, create a vbs file with the following code and run it.
'
' Set driver = CreateObject("Selenium.FirefoxDriver")
' driver.Start
' WScript.Echo "Click OK to quit"
'


Public Sub OpenURL()
  Dim driver As WebDriver
  Set driver = GetObject("Selenium.WebDriver")
  driver.Get "https://www.google.co.uk"
End Sub
