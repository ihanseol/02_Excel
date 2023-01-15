Attribute VB_Name = "browsers"
' This module contains examples on how to work
' with a specific browser.
'


Private Sub Use_Chrome()
    Set driver = CreateObject2("Selenium.ChromeDriver")
    driver.Get "https://www.google.co.uk"
    driver.Quit
End Sub


Private Sub Use_Firefox()
    Set driver = CreateObject2("Selenium.FirefoxDriver")
    driver.Get "https://www.google.co.uk"
    driver.Quit
End Sub


Private Sub Use_Opera()
    Set driver = CreateObject2("Selenium.OperaDriver")
    driver.Get "https://www.google.co.uk"
    driver.Quit
End Sub


Private Sub Use_InternetExplorer()
    Set driver = CreateObject2("Selenium.IEDriver")
    driver.Get "https://www.google.co.uk"
    driver.Quit
End Sub


Private Sub Use_PhantomJS()
    Set driver = CreateObject2("Selenium.PhantomJSDriver")
    driver.Get "https://www.google.co.uk"
    driver.Quit
End Sub


Private Sub Use_FirefoxLight()
    ' Firefox Light:
    ' http://sourceforge.net/projects/lightfirefox/
    
    Set driver = CreateObject2("Selenium.FirefoxDriver")
    driver.SetBinary "C:\...\light.exe"
    driver.Get "https://www.google.co.uk"
    driver.Quit
End Sub


Private Sub Use_CEF()
    ' Chromium Embedded Framework:
    '  https://cefbuilds.com/
    
    Set driver = CreateObject2("Selenium.ChromeDriver")
    driver.SetBinary "C:\...\cefclient.exe"
    driver.AddArgument "url=data:,"
    driver.Get "https://www.google.co.uk"
    driver.Quit
End Sub

