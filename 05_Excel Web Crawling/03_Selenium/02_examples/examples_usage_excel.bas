Attribute VB_Name = "usage_excel"
' This module contains examples to read and write
' data from and to an excel sheet.
'

Private Table As New Selenium.Table
Private Assert As New Selenium.Assert
Private Verify As New Selenium.Verify
Private Keys As New Selenium.Keys


Private Sub Use_Cell_text()
  Dim driver As New FirefoxDriver
  driver.Get "https://en.wikipedia.org/wiki/Main_Page"
  
  'get the input box
  Dim ele As WebElement
  Set ele = driver.FindElementById("searchInput")
  
  'Search the text from the cell A3
  ele.SendKeys [b3]
  ele.Submit
  
  driver.Quit
End Sub

Public Sub VerifyTitles()
  Dim driver As New FirefoxDriver
  
  Dim row
  For Each row In Table.From([Sheet1!A3]).Where("Id > 0")
      'open the page with the link in column "Link"
      driver.Get row("Link")
      
      'Verify the title and set the result in column "Result"
      row("Result") = Verify.Equals(row("ExpectedTitle"), driver.title)
  Next
  
  driver.Quit
End Sub


Public Sub ListLinks()
  Dim driver As New FirefoxDriver
  
  'open the page with the URL in cell A2
  driver.Get [Sheet2!A4]
  
  'get all the href attributes
  Dim links As List
  Set links = driver.FindElementsByTag("a").Attribute("href")
  links.Distinct
  links.Sort
  
  'writes the href values in cell A7
  links.ToExcel [Sheet2!A7]
  
  driver.Quit
End Sub


Public Sub TakeScreenShoot()
  Dim driver As New FirefoxDriver
  
  'open the page with the URL in cell A4
  driver.Get [Sheet3!A4]
  
  'Take the screenshoot in cell A7
  driver.TakeScreenshot().ToExcel [Sheet3!A7]
  
  driver.Quit
End Sub

