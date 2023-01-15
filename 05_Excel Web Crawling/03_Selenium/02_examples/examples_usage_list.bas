Attribute VB_Name = "usage_list"
Private Assert As New Selenium.Assert
Private Keys As New Selenium.Keys


Private Sub Handle_Dropdown_List()
  Dim driver As New FirefoxDriver
  driver.Get "http://the-internet.herokuapp.com/dropdown"
  
  Dim ele As SelectElement
  Set ele = driver.FindElementById("dropdown").AsSelect
  ele.SelectByText "Option 2"
  Assert.Equals "Option 2", ele.SelectedOption.Text
  
  driver.Quit
End Sub



Private Sub Handle_Multi_Select()
  Dim driver As New FirefoxDriver
  driver.Get "http://odyniec.net/articles/multiple-select-fields"
  
  Dim ele As WebElement
  For Each ele In driver.FindElementsByXPath("//select[@name='ingredients[]']/option")
    ele.Click Keys.Control
  Next
  
  driver.Quit
End Sub
