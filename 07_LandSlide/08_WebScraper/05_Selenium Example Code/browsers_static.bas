Attribute VB_Name = "browsers_static"
' This module contains is an example on how to use
' the same instance of a browser with different
' procedures.
'

Private driver As New Selenium.FirefoxDriver
Private Assert As New Selenium.Assert
Private Verify As New Selenium.Verify
Private Waiter As New Selenium.Waiter
Private utils As New Selenium.utils
Private Keys As New Selenium.Keys
Private By As New Selenium.By


Public Sub NavigateToURL1()
  driver.Get [Sheet4!B2]
End Sub


Public Sub NavigateToURL2()
  driver.Get [Sheet4!B5]
End Sub


Public Sub QuitDriver()
  driver.Quit
  Set driver = Nothing
End Sub

