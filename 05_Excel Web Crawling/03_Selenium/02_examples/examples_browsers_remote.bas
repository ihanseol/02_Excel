Attribute VB_Name = "browsers_remote"
' This module contains examples on how to work with
' a browser installed on another station.
'
' Selenium Standalone Server:
'  http://www.seleniumhq.org/download/
'
' Command line to start the server:
'  java -jar selenium-server-standalone-2.47.1.jar
'

Const SERVER = "http://127.0.0.1:4444/wd/hub"

Private Sub Take_ScreenShot_Remotely()
  Dim driver As New WebDriver
  driver.StartRemotely SERVER, "safari"
  'open the page with the URL in Sheet3 in cell A2
  driver.Get [Sheet3!A4]
  
  'Take the screenshoot
  driver.TakeScreenshot().ToExcel [Sheet3!A7]
  
  driver.Quit
End Sub

