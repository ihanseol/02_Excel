Attribute VB_Name = "usage_wait"
Private Assert As New Selenium.Assert
Private Waiter As New Selenium.Waiter


Private Sub Should_Wait_For_Delegate()
  Dim driver As New FirefoxDriver

  ' without delegate
  While Waiter.Not(WaitDelegate1(), timeout:=2000): Wend
  
  ' without delegate with argument
  While Waiter.Not(WaitDelegate2(driver), timeout:=2000): Wend

  ' with delegate on the driver
  driver.Until AddressOf WaitDelegate1, timeout:=2000
  
  ' with delegate with argument
  Waiter.Until AddressOf WaitDelegate1, driver, timeout:=2000
  
  ' with delegate without argument
  Waiter.Until AddressOf WaitDelegate2, timeout:=2000
End Sub


Private Function WaitDelegate1()
  WaitDelegate1 = True
End Function


Private Function WaitDelegate2(driver As WebDriver)
  WaitDelegate2 = True
End Function
