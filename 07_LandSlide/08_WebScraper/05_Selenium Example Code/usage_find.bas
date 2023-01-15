Attribute VB_Name = "usage_find"

Private Assert As New Selenium.Assert

Private Sub Find_By_LinkText()
  Dim drv As New FirefoxDriver
  drv.Get "https://en.wikipedia.org/wiki/Main_Page"

  drv.FindElementByLinkText("Talk").Click
  Assert.Matches "User talk.*", drv.title
  
  drv.Quit
End Sub

Private Sub Find_By_Value()
  Dim drv As New FirefoxDriver
  drv.Get "https://www.google.co.uk"
  
  Dim ele As WebElement
  Set ele = drv.FindElementByXPath("//input[@value='Google Search']")
  Assert.Equals "Google Search", ele.Value
  
  drv.Quit
End Sub

Private Sub Find_By_Partial_Value()
  Dim drv As New FirefoxDriver
  drv.Get "https://www.google.co.uk"
  
  Dim ele As WebElement
  Set ele = drv.FindElementByXPath("//input[contains(@value,'Search')]")
  Assert.Equals "Google Search", ele.Value
  
  drv.Quit
End Sub


Private Sub Find_By_Text()
  Dim drv As New FirefoxDriver
  drv.Get "https://www.google.co.uk"
  
  If drv.FindElementsByXPath("//*[contains(text(),'Google')]").count Then
    Debug.Print "Text is present"
  Else
    Debug.Print "Text is not present"
  End If
  
  drv.Quit
End Sub


Private Sub Text_Exists()
  Const JS_HAS_TEXT = "return document.body.textContent.indexOf(arguments[0]) > -1"

  Dim drv As New FirefoxDriver
  drv.Get "https://www.google.co.uk"
  
  If drv.ExecuteScript(JS_HAS_TEXT, "Google") Then
    Debug.Print "Text is present"
  Else
    Debug.Print "Text is not present"
  End If
  
  drv.Quit
End Sub
