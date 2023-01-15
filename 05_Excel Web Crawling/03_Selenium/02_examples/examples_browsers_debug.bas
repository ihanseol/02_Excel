Attribute VB_Name = "browsers_debug"
' This module contains examples on how to work with
' a Chrome based browser that was launched with a
' listening debug port.
'


Private Sub Connect_To_Chrome()
  'Use this command line to launch the browser:
  'chrome.exe -remote-debugging-port=9222
  
  Dim driver As New ChromeDriver
  driver.SetCapability "debuggerAddress", "127.0.0.1:9222"
  driver.Get "https://en.wikipedia.org"
  driver.Quit
End Sub


Private Sub Connect_To_CFE()
  'Chromium Embedded Framework: https://cefbuilds.com/
  'Use this command line to launch the browser:
  'cefclient.exe -remote-debugging-port=9333 -url=data:,
  
  Dim driver As New ChromeDriver
  driver.SetCapability "debuggerAddress", "127.0.0.1:9333"
  driver.Get "https://en.wikipedia.org"
  driver.Quit
End Sub


Private Sub Connect_To_Firefox()
  'Firefox must have the WebDriver extension installed:
  '%USERPROFILE%\AppData\Local\SeleniumBasic\firefoxdriver.xpi
  'To use another port, set the preference webdriver_firefox_port in about:config
  
  Dim driver As New FirefoxDriver
  driver.SetCapability "debuggerAddress", "127.0.0.1:7055"
  driver.Get "https://en.wikipedia.org"
End Sub

