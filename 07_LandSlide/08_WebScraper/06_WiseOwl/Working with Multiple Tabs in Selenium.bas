Attribute VB_Name = "Module1"
Option Explicit

Dim cd As Selenium.ChromeDriver

Sub UsingTabs()

    Dim w As Selenium.Window
    Dim NewsItems As Selenium.WebElements
    Dim NewsItem As Selenium.WebElement
    Dim FirstLink As String
    
    Set cd = New Selenium.ChromeDriver
    
    cd.Start
    cd.Get "https://en.wikipedia.org/wiki/Main_Page"
    
    Set NewsItems = cd.FindElementsByCss("#mp-itn > ul > li")
    
    For Each NewsItem In NewsItems
    
        FirstLink = NewsItem.FindElementsByCss("a")(1).Attribute("href")
        
        cd.ExecuteScript "window.open(arguments[0])", FirstLink
        cd.SwitchToNextWindow
        Debug.Print cd.FindElementByCss("#firstHeading").Text
        cd.Window.Close
        cd.SwitchToPreviousWindow

    Next NewsItem
    
End Sub

Sub GoToSpace()

    Dim BaseURL As String
    Dim NextURL As String
    Dim DateOptions As Selenium.WebElements
    Dim DateOption As Selenium.WebElement
    Dim tbls As Selenium.WebElements
    Dim t As Selenium.WebElement
    Dim ws As Worksheet
    
    BaseURL = "https://finance.yahoo.com/quote/SPCE/options"
    
    Set cd = New Selenium.ChromeDriver
    
    cd.Start
    cd.Get BaseURL
    
    cd.FindElementByCss("button").Click
    
    Set DateOptions = cd.FindElementsByCss("option")
    
    Set ws = ThisWorkbook.Worksheets.Add
    
    For Each DateOption In DateOptions
        NextURL = BaseURL & "?date=" & DateOption.Attribute("value")
        
        cd.ExecuteScript "window.open(arguments[0])", NextURL
        cd.SwitchToNextWindow
        
        Set tbls = cd.FindElementsByCss("table")
        
        For Each t In tbls
            t.AsTable.ToExcel ws.Range("A1048576").End(xlUp).Offset(1, 0)
        Next t
        
        cd.Window.Close
        cd.SwitchToPreviousWindow
        
    Next DateOption
    
    ws.Range("A1").Value = Now
    ws.Range("A1").CurrentRegion.EntireColumn.AutoFit
    
End Sub
























