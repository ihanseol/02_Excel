Attribute VB_Name = "browsers_profile"
' This module contains examples on how to work with
' a customized profile.
'


Private Sub Use_Chrome_With_Custom_profile_name()
  ' Profiles folder : %APPDATA%\Google\Chrome\Profiles
  ' Note that with Chrome the profile is always persistant
  
  Dim driver As New ChromeDriver
  driver.SetProfile "Selenium"
  
  driver.Get "https://www.google.co.uk"
  driver.Quit
End Sub

Private Sub Use_Chrome_With_Custom_profile_path()
  ' Default profile : %LOCALAPPDATA%\Google\Chrome\User Data
  ' Profiles folder : %APPDATA%\Google\Chrome\Profiles
  ' Note that with Chrome the profile is always persistant
  
  Dim driver As New ChromeDriver
  driver.SetProfile "%LOCALAPPDATA%\Google\Chrome\User Data"
  
  driver.Get "https://www.google.co.uk"
  driver.Quit
End Sub

Private Sub Use_Firefox_With_Custom_profile_name()
  ' To manage firefox profiles: firefox -p
  ' Profiles folder: %APPDATA%\Mozilla\Firefox\Profiles
  ' When persistant is False, the driver works with a copy in the Temp folder.
  
  Dim driver As New FirefoxDriver
  driver.SetProfile "Selenium", persistant:=True
  
  driver.Get "https://www.google.co.uk"
  driver.Quit
End Sub

Private Sub Use_Firefox_With_Custom_profile_path()
  ' To manage the profiles: firefox -p
  ' Profiles folder: %APPDATA%\Mozilla\Firefox\Profiles
  ' When persistant is False, the driver works with a copy in the Temp folder.
  
  Dim driver As New FirefoxDriver
  driver.SetProfile "%APPDATA%\Mozilla\Firefox\Profiles\kfvj49h4.Selenium", persistant:=True
  
  driver.Get "https://www.google.co.uk"
  driver.Quit
End Sub
