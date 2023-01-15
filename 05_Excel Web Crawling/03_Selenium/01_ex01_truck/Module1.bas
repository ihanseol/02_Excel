Attribute VB_Name = "Module1"

Option Explicit


Sub google_search()
    
    Dim p As New parser
    
    Dim row As Integer
    row = 2
    
    Dim bot As WebDriver
    Dim elem As WebElement
    
    Set bot = New WebDriver
    
    
    bot.Start "chrome"
    bot.Get "https://ai.fmcsa.dot.gov/hhg/search.asp"
    

    bot.FindElementByName("DOT").SendKeys Sheet1.Cells(row, 5).Value
    bot.FindElementByName("Submit").Click
    
    'wait until the page arrived
    Do Until InStr(1, bot.Url, "SearchResults") > 0
        Application.Wait (DateAdd("s", 1, Now))
        DoEvents
    Loop
    
    Set elem = bot.FindElementByClass("MiddleTDFMCSA")
    elem.FindElementByTag("a").Click
    
    'wait until the page arrived
    Do Until InStr(1, bot.Url, "SearchDetails") > 0
        Application.Wait (DateAdd("s", 1, Now))
        DoEvents
    Loop
    
    p.text = bot.FindElementByTag("body").Attribute("innerHTML")
    p.position = 1
    p.moveTo ">Telephone</td>"
    p.moveTo "<td "
    p.moveTo ">"
    
    Sheet1.Cells(row, 6).Value = p.getText("&nbsp;")
    
    Stop

End Sub
