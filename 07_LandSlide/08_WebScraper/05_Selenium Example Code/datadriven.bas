Attribute VB_Name = "datadriven"

Public Sub VerifyTitles()
    Dim Table, Verify, row, driver
    
    Set Table = CreateObject2("Selenium.Table")
    Set Verify = CreateObject2("Selenium.Verify")
    Set driver = CreateObject2("Selenium.FirefoxDriver")
    
    For Each row In Table.From([Sheet1!A3]).Where("Id > 0")
        'open the page with the link in column "Link"
        driver.Get row("Link")
        
        'Verify the title and set the result in column "Result"
        row("Result") = Verify.Equals(row("ExpectedTitle"), driver.Title)
    Next
    
    driver.Quit
End Sub

Public Sub ListLinks()
    Dim driver, links
    
    'creates a new instance
    Set driver = CreateObject2("Selenium.FirefoxDriver")
    
    'open the page with the URL in cell A2
    driver.Get [A4]
    
    'get all the href attributes
    Set links = driver.FindElementsByTag("a").Attribute("href")
    links.Distinct
    links.Sort
    
    'writes the href values in cell A7
    links.ToExcel [A7]
    
    driver.Quit
End Sub

Public Sub TakeScreenShoot()
    Set driver = CreateObject2("Selenium.FirefoxDriver")
    
    'open the page with the URL in cell A4
    driver.Get [A4]
    
    'Take the screenshoot in cell A7
    driver.TakeScreenShot().ToExcel [A7]
    
    driver.Quit
End Sub

