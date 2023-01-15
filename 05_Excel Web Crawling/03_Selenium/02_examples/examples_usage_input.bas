Attribute VB_Name = "usage_input"
Private Assert As New Selenium.Assert
Private Keys As New Selenium.Keys


Private Sub Handle_Input()
  Dim driver As New FirefoxDriver
  driver.Get "https://en.wikipedia.org/wiki/Main_Page"
  
  'get the input box
  Dim ele As WebElement
  Set ele = driver.FindElementById("searchInput")
  
  'set the text
  ele.SendKeys "abc"
  
  'get the text
  Dim txt As String
  txt = ele.Value
  
  'assert text
  Assert.Equals "abc", txt
  
  driver.Quit
End Sub


Private Sub Handle_Input_With_Script()
  Dim driver As New FirefoxDriver
  driver.Get "https://en.wikipedia.org"
  
  driver.FindElementById("searchInput").ExecuteScript _
      "this.value=arguments[0];", "my value"
  
  driver.Quit
End Sub


Private Sub Handle_Input_With_Autoit()
  Set autoit = CreateObject("AutoItX3.Control")
  autoit.Send "{ENTER}"
End Sub


Private Sub Handle_Input_With_WScript()
  Set wsh = CreateObject("WScript.Shell")
  wsh.SendKeys "{ENTER}"
End Sub


Private Sub Handle_TinyMCE_API()
  Dim driver As New FirefoxDriver
  driver.Get "http://the-internet.herokuapp.com/tinymce"
  
  'set a content using javascript
  data_set = "<p><em>12345</em></p>"
  driver.ExecuteScript "tinyMCE.activeEditor.setContent(arguments[0])", data_set
  
  'insert a content using javascript
  data_instert = "<p><em>abcdefg</em></p>"
  driver.ExecuteScript "tinyMCE.activeEditor.insertContent(arguments[0])", data_instert
  
  'read and evaluate a content
  data_read = driver.ExecuteScript("return tinyMCE.activeEditor.getContent()")
  Assert.Equals data_instert & vbLf & data_set, data_read
  
  driver.Quit
End Sub


Private Sub Handle_TinyMCE_Simulate()
  Dim driver As New FirefoxDriver
  driver.Get "http://the-internet.herokuapp.com/tinymce"
  
  Dim btBold As WebElement, btItalic As WebElement, body As WebElement
  
  Set btBold = driver.FindElementByCss(".mce-btn[aria-label='Bold'] button")
  Set btItalic = driver.FindElementByCss(".mce-btn[aria-label='Italic'] button")

  'clear body
  driver.SwitchToFrame 0
  Set body = driver.FindElementByTag("body")
  body.Clear

  'set bold and italic
  driver.SwitchToDefaultContent
  btBold.Click
  btItalic.Click
  
  'set and eval text
  driver.SwitchToFrame 0
  body.SendKeys "abcde12345"
  Assert.Equals "abcde12345", body.Text
  
  driver.Quit
End Sub






