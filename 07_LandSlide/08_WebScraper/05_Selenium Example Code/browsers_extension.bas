Attribute VB_Name = "browsers_extension"
' This module contains examples on how to work with
' an extension.
'


Private Sub Use_Chrome_With_Extension()
  ' To download an extension:
  ' http://chrome-extension-downloader.com
  ' To manage the extension preferences:
  ' Developper Tools > Resources > Local Storage > chrome-extension://...
  
  Dim driver As New ChromeDriver
  driver.AddExtension "C:\Users\florent\Downloads\Personal-Blocklist-(by-Google)_v2.6.1.crx"
  driver.SetPreference "plugins.plugins_disabled", Array("Adobe Flash Player")
  driver.Get "chrome-extension://nolijncfnkgaikbjbdaogikpmpbdcdef/manager.html"
  driver.ExecuteScript "localStorage.setItem('blocklist', '[""wikipedia.org""]');"
  
  driver.Get "https://www.google.co.uk"
  driver.Quit
End Sub


Private Sub Use_Firefox_With_Extension()
  ' To download an extension, use a browser other than Firefox
  
  Dim driver As New FirefoxDriver
  driver.AddExtension "C:\Users\florent\Downloads\firebug-2.0.12-fx.xpi"
  driver.SetPreference "extensions.firebug.showFirstRunPage", False
  
  driver.Get "https://www.google.co.uk"
  driver.Quit
End Sub
