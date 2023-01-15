Attribute VB_Name = "browsers_all"
' This module contains examples on how to work
' with a specific browser.
'


Private Sub Use_Chrome()
  Dim driver As New ChromeDriver
  driver.Get "https://www.google.co.uk"
  driver.Quit
End Sub


Private Sub Use_Firefox()
  Dim driver As New FirefoxDriver
  driver.Get "https://www.google.co.uk"
  driver.Quit
End Sub


Private Sub Use_Opera()
  Dim driver As New OperaDriver
  driver.Get "https://www.google.co.uk"
  driver.Quit
End Sub


Private Sub Use_InternetExplorer()
  Dim driver As New IEDriver
  driver.Get "https://www.google.co.uk"
  driver.Quit
End Sub


Private Sub Use_PhantomJS()
  Dim driver As New PhantomJSDriver
  driver.Get "https://www.google.co.uk"
  driver.Quit
End Sub


Private Sub Use_FirefoxLight()
  ' Firefox Light:
  ' http://sourceforge.net/projects/lightfirefox/
  
  Dim driver As New FirefoxDriver
  driver.SetBinary "C:\...\light.exe"
  driver.Get "https://www.google.co.uk"
  driver.Quit
End Sub


Private Sub Use_CEF()
  ' Chromium Embedded Framework:
  ' https://cefbuilds.com
  
  Dim driver As New ChromeDriver
  driver.SetBinary "C:\...\cefclient.exe"
  driver.AddArgument "url=data:,"
  driver.Get "https://www.google.co.uk"
  driver.Quit
End Sub

