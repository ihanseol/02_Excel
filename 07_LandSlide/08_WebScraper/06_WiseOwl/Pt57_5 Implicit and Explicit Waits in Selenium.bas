Attribute VB_Name = "Module1"
Option Explicit

Private cd As Selenium.ChromeDriver

Sub CairoFlights()

    Set cd = New Selenium.ChromeDriver
    
    cd.Start
    cd.Get "https://www.cairo-airport.com/en-us/Flights/Flight-Information"
    
    Dim FlightTables As Selenium.WebElements
    
    Set FlightTables = cd.FindElementsByCss(".table.table-striped.fl-table")
    
    FlightTables(1).AsTable.ToExcel ThisWorkbook.Worksheets.Add.Range("A2")
    Range("A1").Value = "Arrivals Today"
    
    cd.FindElementByCss("#to").Click
    
    cd.Wait 1000
     
    FlightTables(1).AsTable.ToExcel ThisWorkbook.Worksheets.Add.Range("A2")
    Range("A1").Value = "Arrivals Tomorrow"
    
    cd.FindElementByCss("#hrfDepartures").Click
    
    FlightTables(2).FindElementByCss("#Departures").WaitDisplayed
    
    FlightTables(2).AsTable.ToExcel ThisWorkbook.Worksheets.Add.Range("A2")
    Range("A1").Value = "Departures Today"
    
    cd.FindElementByCss("#to2").Click
    
    cd.Wait 1000
    
    FlightTables(2).AsTable.ToExcel ThisWorkbook.Worksheets.Add.Range("A2")
    Range("A1").Value = "Departures Tomorrow"
    
End Sub

Sub WaitForElements()
    
    Set cd = New Selenium.ChromeDriver
    
    cd.Start
    cd.Get "https://en.wikipedia.org/wiki/Main_Page"
    
    Dim FindBy As New Selenium.By
        
    'Interact with Search Input
    Dim SearchInput As Selenium.WebElement
    
    Set SearchInput = cd.FindElementByCss("#searchInput")
    
    SearchInput.SendKeys "glossop"
        
    'Get list of suggestions
    Dim SuggestionsDiv As Selenium.WebElement
    
    Set SuggestionsDiv = cd.FindElementByCss(".suggestions")
    
    Debug.Print "Suggestions displayed " & SuggestionsDiv.IsDisplayed
    
    SuggestionsDiv.WaitDisplayed
    
    Debug.Print "Suggestions displayed " & SuggestionsDiv.IsDisplayed
    
    Dim Suggestions As Selenium.WebElements
    Dim Suggestion As Selenium.WebElement
    
    Set Suggestions = SuggestionsDiv.FindElementsByCss("a")
    
    For Each Suggestion In Suggestions
        Debug.Print Suggestion.Text, Suggestion.Attribute("href")
    Next Suggestion
    
End Sub
