Attribute VB_Name = "usage_authentication"
Private Assert As New Selenium.Assert


Private Sub Basic_Authentication_In_URL()
  Dim driver As New IEDriver
  driver.Get "http://admin:admin@the-internet.herokuapp.com/basic_auth"
  
  txt = driver.FindElementByCss(".example p").Text
  Assert.Matches "^Congratulations!", txt
  
  driver.Quit
End Sub


Private Sub AuthenticationDialog_Selenium()
  Dim driver As New IEDriver
  driver.Get "http://the-internet.herokuapp.com/basic_auth"
  
  Dim dlg As Alert: Set dlg = driver.SwitchToAlert(Raise:=False)
  If Not dlg Is Nothing Then
    dlg.SetCredentials "admin", "admin"
    dlg.Accept
  End If
  
  txt = driver.FindElementByCss(".example p").Text
  Assert.Matches "^Congratulations!", txt
  driver.Quit
End Sub


Private Sub AuthenticationDialog_WScript()
  Dim driver As New IEDriver
  driver.Get "http://the-internet.herokuapp.com/basic_auth"
  
  Dim dlg As Alert: Set dlg = driver.SwitchToAlert(Raise:=False)
  If Not dlg Is Nothing Then
    Set wsh = CreateObject("WScript.Shell")
    wsh.SendKeys "admin"
    wsh.SendKeys "{TAB}"
    wsh.SendKeys "admin"
    dlg.Accept
  End If
  
  txt = driver.FindElementByCss(".example p").Text
  Assert.Matches "^Congratulations!", txt
  driver.Quit
End Sub


Private Sub AuthenticationDialog_AutoIt()
  Dim driver As New IEDriver
  driver.Get "http://the-internet.herokuapp.com/basic_auth"
  
  Dim dlg As Alert: Set dlg = driver.SwitchToAlert(Raise:=False)
  If Not dlg Is Nothing Then
    Set aut = CreateObject("AutoItX3.Control")
    aut.Send "admin"
    aut.Send "{TAB}"
    aut.Send "admin"
    dlg.Accept
  End If
  
  txt = driver.FindElementByCss(".example p").Text
  Assert.Matches "^Congratulations!", txt
  driver.Quit
End Sub

