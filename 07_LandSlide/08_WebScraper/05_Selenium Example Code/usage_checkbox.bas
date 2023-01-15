Attribute VB_Name = "usage_checkbox"
Private Assert As New Selenium.Assert


Private Sub Handle_Checkbox()
  Dim driver As New FirefoxDriver
  driver.Get "http://the-internet.herokuapp.com/checkboxes"
  
  'Get the 2 second checkbox
  Dim cs As WebElement
  Set cb = driver.FindElementByCss("#checkboxes input:nth-of-type(2)")
  
  'Assert the checkbox is checked
  Assert.True cb.IsSelected
  
  'Uncheck the checkbox
  cb.Click
  
  'Assert the checkbox is unchecked
  Assert.False cb.IsSelected
  driver.Quit
End Sub
