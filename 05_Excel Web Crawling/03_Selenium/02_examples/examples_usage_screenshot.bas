Attribute VB_Name = "usage_screenshot"

Private Sub Take_ScreenShot_Content()
  Dim driver As New FirefoxDriver
  driver.Get "https://en.wikipedia.org/wiki/Main_Page"
  
  'take a screenshot of the page
  Dim img As Image
  Set img = driver.TakeScreenshot()
  
  'save the image in the folder of the workbook
  img.SaveAs ThisWorkbook.Path & "\sc-content.png"
  
  driver.Quit
End Sub


Private Sub Take_ScreenShot_Element()
  Dim driver As New FirefoxDriver
  driver.Get "https://en.wikipedia.org/wiki/Main_Page"
  
  'take a screenshot of an element
  Dim img As Image
  Set img = driver.FindElementById("mp-bottom").TakeScreenshot()
  
  'save the image in the folder of the workbook
  img.SaveAs ThisWorkbook.Path & "\sc-element.png"
  
  driver.Quit
End Sub


Private Sub Take_ScreenShot_Element_Highlight()
  Const JS_ADD_YELLOW_BORDER = "window._eso=this.style.outline;this.style.outline='#FFFF00 solid 5px';"
  Const JS_DEL_YELLOW_BORDER = "this.style.outline=window._eso;"
  
  
  Dim drv As New FirefoxDriver
  drv.Get "https://en.wikipedia.org/wiki/Eurytios_Krater"
  Set ele = drv.FindElementById("searchInput")
  
  ' Apply a yellow outline
  ele.ExecuteScript JS_ADD_YELLOW_BORDER
  
  ' Take the screenshot
  Set img = drv.TakeScreenshot()
  img.SaveAs ThisWorkbook.Path & "\sc-element-highlight.png"
  
  ' Remove the outline
  ele.ExecuteScript JS_DEL_YELLOW_BORDER
  
  drv.Quit
End Sub


Private Sub Take_ScreenShot_Desktop()
  Dim utils As New utils
  Dim driver As New FirefoxDriver
  driver.Get "https://en.wikipedia.org/wiki/Main_Page"
  
  'take a screenshot of the desktop
  Set img = utils.TakeScreenshot()
  
  'save the image in the folder of the workbook
  img.SaveAs ThisWorkbook.Path & "\sc-desktop.png"
  
  driver.Quit
End Sub

